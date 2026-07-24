from __future__ import annotations

import base64
import io
import json
import math
import mimetypes
import os
import re
import socket
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote

import pandas as pd


ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
GENERATED = ROOT / "generated"
HOST = os.environ.get("FRAME_PARTS_TOOL_HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT") or os.environ.get("FRAME_PARTS_TOOL_PORT", "5508"))

SAP_COLUMNS = ["物料", "物料描述", "需求日期", "需求量", "供应商"]
SRM_COLUMNS = ["计划年份", "计划月份", "供应商编号", "零件编码", "零件名称", "N+1总数"]


def normalize_code(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return re.sub(r"\s+", "", text).upper()


def to_number(value: Any) -> float:
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return 0.0


def read_table(file_bytes: bytes, filename: str) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()
    buffer = io.BytesIO(file_bytes)
    if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}:
        return pd.read_excel(buffer, dtype=object)
    try:
        return pd.read_csv(buffer, dtype=object, encoding="utf-8-sig")
    except UnicodeDecodeError:
        buffer.seek(0)
        return pd.read_csv(buffer, dtype=object, encoding="gbk")


def require_columns(df: pd.DataFrame, required: list[str], label: str) -> None:
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"{label}缺少必要字段：" + "、".join(missing))


def next_month_label(year: int, month: int) -> str:
    if month >= 12:
        return f"{year + 1}-01"
    return f"{year}-{month + 1:02d}"


def month_key(value: Any) -> str:
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return ""
    return f"{dt.year}-{dt.month:02d}"


def item_type(code: str, name: str) -> str:
    # 横梁：中文名称含“横梁”且图号含 T；小件：图号不含 T。
    # 图号含 T 但名称不是横梁的物料不纳入本工具输出。
    has_t = "T" in code.upper()
    if "横梁" in str(name) and has_t:
        return "横梁"
    if not has_t:
        return "小件"
    return "排除"


def build_output_excel(crossbeam: pd.DataFrame, small: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        crossbeam.to_excel(writer, index=False, sheet_name="横梁储备清单")
        small.to_excel(writer, index=False, sheet_name="小件储备清单")
        for sheet in writer.sheets.values():
            sheet.freeze_panes = "A2"
            for column_cells in sheet.columns:
                max_len = max(len(str(cell.value or "")) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 28)
    return buffer.getvalue()


def process_payload(files: dict[str, tuple[str, bytes]], fields: dict[str, str]) -> dict[str, Any]:
    if "sapFile" not in files or "srmFile" not in files:
        raise ValueError("请同时上传 SAP 前三个月用量表和 SRM 未来预测表。")

    crossbeam_threshold = to_number(fields.get("crossbeamThreshold", 20))
    small_threshold = to_number(fields.get("smallThreshold", 30))
    crossbeam_reserve_rate = to_number(fields.get("crossbeamReserveRate", 30)) / 100
    small_reserve_rate = to_number(fields.get("smallReserveRate", 30)) / 100

    sap_name, sap_bytes = files["sapFile"]
    srm_name, srm_bytes = files["srmFile"]
    sap = read_table(sap_bytes, sap_name).fillna("")
    srm = read_table(srm_bytes, srm_name).fillna("")
    require_columns(sap, SAP_COLUMNS, "SAP表")
    require_columns(srm, SRM_COLUMNS, "SRM表")

    sap = sap.copy()
    sap["图号"] = sap["物料"].map(normalize_code)
    sap["物料名称"] = sap["物料描述"].astype(str).str.strip()
    sap["供应商代码"] = sap["供应商"].astype(str).str.strip()
    sap["月份"] = sap["需求日期"].map(month_key)
    sap["月用量"] = sap["需求量"].map(to_number)
    sap = sap[(sap["图号"] != "") & (sap["月份"] != "")]

    recent_months = sorted(sap["月份"].unique())[-3:]
    if len(recent_months) < 3:
        raise ValueError("SAP表中可识别的月份少于3个月，请检查需求日期字段。")

    pivot = (
        sap[sap["月份"].isin(recent_months)]
        .pivot_table(index="图号", columns="月份", values="月用量", aggfunc="sum", fill_value=0)
        .reset_index()
    )
    for m in recent_months:
        if m not in pivot.columns:
            pivot[m] = 0

    info = (
        sap.sort_values(["图号", "月份"])
        .groupby("图号", as_index=False)
        .agg({"物料名称": "first", "供应商代码": "first"})
    )
    usage = info.merge(pivot, on="图号", how="left")
    usage["前三个月月均用量"] = usage[recent_months].sum(axis=1) / 3
    usage["物料类型"] = usage.apply(lambda row: item_type(row["图号"], row["物料名称"]), axis=1)

    srm = srm.copy()
    srm["图号"] = srm["零件编码"].map(normalize_code)
    srm["SRM供应商代码"] = srm["供应商编号"].astype(str).str.strip()
    srm["下月预测需求"] = srm["N+1总数"].map(to_number)
    years = [int(to_number(v)) for v in srm["计划年份"].tolist() if to_number(v) > 0]
    months = [int(to_number(v)) for v in srm["计划月份"].tolist() if to_number(v) > 0]
    target_month = next_month_label(max(set(years), key=years.count), max(set(months), key=months.count)) if years and months else "下个月"
    srm_sum = (
        srm[srm["图号"] != ""]
        .groupby("图号", as_index=False)
        .agg({"下月预测需求": "sum", "SRM供应商代码": "first"})
    )

    merged = usage.merge(srm_sum, on="图号", how="left")
    merged["下月预测需求"] = merged["下月预测需求"].fillna(0)
    merged["供应商代码"] = merged["供应商代码"].where(merged["供应商代码"] != "", merged["SRM供应商代码"].fillna(""))
    merged["输出月份"] = target_month
    def reserve_quantity(row: pd.Series) -> int:
        if row["物料类型"] == "横梁":
            rate = crossbeam_reserve_rate
        elif row["物料类型"] == "小件":
            rate = small_reserve_rate
        else:
            rate = 0
        return int(math.ceil(to_number(row["前三个月月均用量"]) * rate))

    merged["下月储备数量"] = merged.apply(reserve_quantity, axis=1)

    def monthly_usage_pass(row: pd.Series) -> bool:
        if row["物料类型"] == "横梁":
            threshold = crossbeam_threshold
        elif row["物料类型"] == "小件":
            threshold = small_threshold
        else:
            return False
        return all(to_number(row[m]) > threshold for m in recent_months)

    merged["每月用量是否达标"] = merged.apply(monthly_usage_pass, axis=1)
    merged["是否纳入储备"] = (
        (merged["物料类型"].isin(["横梁", "小件"]))
        & merged["每月用量是否达标"]
        & (merged["下月预测需求"] > 0)
    )

    output_cols = [
        "输出月份",
        "物料类型",
        "图号",
        "物料名称",
        "供应商代码",
        recent_months[0],
        recent_months[1],
        recent_months[2],
        "前三个月月均用量",
        "下月预测需求",
        "下月储备数量",
    ]
    final = merged[merged["是否纳入储备"]].copy()
    final = final.sort_values(["物料类型", "供应商代码", "图号"])
    final["前三个月月均用量"] = final["前三个月月均用量"].round(2)
    crossbeam = final[final["物料类型"] == "横梁"][output_cols].copy()
    small = final[final["物料类型"] == "小件"][output_cols].copy()

    GENERATED.mkdir(exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    filename = f"frame_parts_reserve_{target_month}_{stamp}.xlsx"
    display_name = f"车架散件储备清单_{target_month}_{stamp}.xlsx"
    file_bytes = build_output_excel(crossbeam, small)
    output_path = GENERATED / filename
    output_path.write_bytes(file_bytes)

    return {
        "summary": {
            "sapRows": int(len(sap)),
            "srmRows": int(len(srm)),
            "months": recent_months,
            "targetMonth": target_month,
            "crossbeamCount": int(len(crossbeam)),
            "smallCount": int(len(small)),
            "totalReserveQty": int(final["下月储备数量"].sum()) if len(final) else 0,
            "crossbeamThreshold": crossbeam_threshold,
            "smallThreshold": small_threshold,
            "crossbeamReserveRate": crossbeam_reserve_rate,
            "smallReserveRate": small_reserve_rate,
        },
        "crossbeam": crossbeam.to_dict(orient="records"),
        "small": small.to_dict(orient="records"),
        "download": f"/download/{filename}",
        "downloadName": display_name,
    }


def parse_multipart(content_type: str, body: bytes) -> tuple[dict[str, tuple[str, bytes]], dict[str, str]]:
    match = re.search(r"boundary=(.+)", content_type)
    if not match:
        raise ValueError("请求缺少 multipart boundary。")
    boundary = ("--" + match.group(1).strip().strip('"')).encode()
    files: dict[str, tuple[str, bytes]] = {}
    fields: dict[str, str] = {}
    for part in body.split(boundary):
        part = part.strip(b"\r\n")
        if not part or part == b"--":
            continue
        header_bytes, _, data = part.partition(b"\r\n\r\n")
        headers = header_bytes.decode("utf-8", errors="ignore")
        data = data.rstrip(b"\r\n")
        name_match = re.search(r'name="([^"]+)"', headers)
        if not name_match:
            continue
        name = name_match.group(1)
        filename_match = re.search(r'filename="([^"]*)"', headers)
        if filename_match and filename_match.group(1):
            files[name] = (filename_match.group(1), data)
        else:
            fields[name] = data.decode("utf-8", errors="ignore")
    return files, fields


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        return

    def send_json(self, status: int, payload: dict[str, Any]) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        request_path = self.path.split("?", 1)[0]
        if request_path.startswith("/download/"):
            filename = unquote(request_path.rsplit("/", 1)[-1])
            target = (GENERATED / filename).resolve()
            if not str(target).startswith(str(GENERATED.resolve())) or not target.exists():
                self.send_error(404)
                return
            raw = target.read_bytes()
            display_name = filename
            if filename.startswith("frame_parts_reserve_"):
                display_name = filename.replace("frame_parts_reserve_", "车架散件储备清单_", 1)
            encoded = quote(display_name)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f"attachment; filename=\"reserve.xlsx\"; filename*=UTF-8''{encoded}")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        if request_path == "/":
            request_path = "/index.html"
        target = (STATIC / request_path.lstrip("/")).resolve()
        if not str(target).startswith(str(STATIC.resolve())) or not target.exists():
            self.send_error(404)
            return
        raw = target.read_bytes()
        self.send_response(200)
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if content_type.startswith("text/") or content_type == "application/javascript":
            content_type = f"{content_type}; charset=utf-8"
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self) -> None:
        if self.path != "/api/process":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            files, fields = parse_multipart(self.headers.get("Content-Type", ""), self.rfile.read(length))
            self.send_json(200, process_payload(files, fields))
        except Exception as exc:
            self.send_json(400, {"error": str(exc)})


def main() -> None:
    local_ip = socket.gethostbyname(socket.gethostname())
    print(f"本机访问：http://127.0.0.1:{PORT}")
    print(f"局域网访问：http://{local_ip}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()


if __name__ == "__main__":
    main()

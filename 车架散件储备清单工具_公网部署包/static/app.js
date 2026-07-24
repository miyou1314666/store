const form = document.getElementById("toolForm");
const resultPanel = document.getElementById("resultPanel");
const downloadLink = document.getElementById("downloadLink");
const ruleCard = document.getElementById("ruleCard");

document.documentElement.classList.add("notranslate");
document.documentElement.setAttribute("translate", "no");
document.body.setAttribute("translate", "no");

if (ruleCard) {
  ruleCard.replaceChildren();
  const title = document.createElement("h2");
  const text = document.createElement("textarea");
  title.textContent = "当前规则";
  text.id = "ruleText";
  text.readOnly = true;
  text.tabIndex = -1;
  text.value = "横梁：名称含“横梁”且图号含 T。\n小件：图号不含 T。\n筛选：三个自然月每月用量均超过对应阈值，且 N+1 需求大于 0。\n计算：三个月平均用量 × 对应储备比例。";
  ruleCard.classList.add("notranslate");
  ruleCard.setAttribute("translate", "no");
  ruleCard.append(title, text);
}

function bindDrop(zoneSelector, inputId, nameId) {
  const zone = document.querySelector(zoneSelector);
  const input = document.getElementById(inputId);
  const name = document.getElementById(nameId);

  input.addEventListener("change", () => {
    name.textContent = input.files[0] ? input.files[0].name : "拖拽文件到这里，或点击选择";
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.remove("dragover");
    });
  });

  zone.addEventListener("drop", (event) => {
    const file = event.dataTransfer.files[0];
    if (!file) return;
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    name.textContent = file.name;
  });
}

bindDrop('[data-zone="sap"]', "sapFile", "sapName");
bindDrop('[data-zone="srm"]', "srmFile", "srmName");

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const button = form.querySelector(".primary");
  button.disabled = true;
  button.textContent = "处理中...";

  try {
    const response = await fetch("/api/process", {
      method: "POST",
      body: new FormData(form),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "处理失败");
    render(payload);
  } catch (error) {
    alert(error.message);
  } finally {
    button.disabled = false;
    button.textContent = "生成储备清单";
  }
});

function fmt(value) {
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/\.00$/, "");
  }
  return value ?? "";
}

function renderRows(tableId, rows, months) {
  const table = document.getElementById(tableId);
  const thead = table.querySelector("thead");
  const tbody = document.querySelector(`#${tableId} tbody`);
  const columns = [
    ["输出月份", "输出月份"],
    ["图号", "图号"],
    ["物料名称", "名称"],
    ["供应商代码", "供应商"],
    ...months.map((month) => [month, `${month}用量`]),
    ["前三个月月均用量", "月均用量"],
    ["下月预测需求", "N+1预测"],
    ["下月储备数量", "储备数量"],
  ];

  thead.innerHTML = "";
  tbody.innerHTML = "";
  const headRow = document.createElement("tr");
  columns.forEach(([, label]) => {
    const th = document.createElement("th");
    th.textContent = label;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);

  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = columns.length;
    td.textContent = "无符合条件的图号";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach(([key]) => {
      const td = document.createElement("td");
      td.textContent = fmt(row[key]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function render(payload) {
  const summary = payload.summary;
  document.getElementById("crossbeamCount").textContent = summary.crossbeamCount;
  document.getElementById("smallCount").textContent = summary.smallCount;
  document.getElementById("reserveQty").textContent = summary.totalReserveQty;
  document.getElementById("targetMonth").textContent = summary.targetMonth;
  const resultMeta = document.getElementById("resultMeta");
  resultMeta.classList.add("notranslate");
  resultMeta.setAttribute("translate", "no");
  resultMeta.textContent =
    `SAP识别 ${summary.sapRows} 行，SRM识别 ${summary.srmRows} 行；参与计算月份：${summary.months.join("、")}；输出月份：${summary.targetMonth}；横梁比例 ${Math.round(summary.crossbeamReserveRate * 100)}%，小件比例 ${Math.round(summary.smallReserveRate * 100)}%。`;
  renderRows("crossbeamTable", payload.crossbeam || [], summary.months || []);
  renderRows("smallTable", payload.small || [], summary.months || []);
  document.getElementById("crossbeamShown").textContent = `（全部 ${payload.crossbeam?.length || 0} 条）`;
  document.getElementById("smallShown").textContent = `（全部 ${payload.small?.length || 0} 条）`;
  downloadLink.replaceChildren(document.createTextNode("下载 Excel 清单"));
  downloadLink.classList.add("notranslate");
  downloadLink.setAttribute("translate", "no");
  downloadLink.href = payload.download;
  downloadLink.download = payload.downloadName || "车架散件储备清单.xlsx";
  resultPanel.classList.remove("hidden");
  resultPanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

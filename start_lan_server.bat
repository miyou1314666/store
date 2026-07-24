@echo off
set FRAME_PARTS_TOOL_HOST=0.0.0.0
set FRAME_PARTS_TOOL_PORT=5555
echo Starting frame parts reserve tool on LAN...
echo Local: http://127.0.0.1:5555
echo Other computers should open your Wi-Fi IPv4 address, for example:
echo http://10.111.47.229:5555
echo.
echo If other computers cannot open it, allow Python through Windows Firewall.
echo Keep this window open while using the tool.
"C:\Users\miyou\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "C:\Users\miyou\Documents\Codex\2026-07-23\new-chat\work\frame-parts-tool\app.py"
pause

@echo off
REM VBA MCP Server Pro - HTTP/SSE Transport Launcher
REM Quick start script for Windows

echo ========================================
echo VBA MCP Server Pro - HTTP/SSE
echo ========================================
echo.

REM Get Windows IP
echo Your IP addresses:
ipconfig | findstr /R "IPv4"
echo.

REM Activate venv and start server
echo Starting server...
echo.
echo Press Ctrl+C to stop the server
echo.

call venv\Scripts\activate.bat
venv\Scripts\vba-mcp-server-pro-http.exe --host 0.0.0.0 --port 8000 --log-level INFO

pause

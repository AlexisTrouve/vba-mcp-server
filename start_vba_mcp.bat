@echo off
REM VBA MCP Pro Server Launcher
REM This script sets up the environment and starts the MCP server

cd /d C:\Users\alexi\Documents\projects\vba-mcp-monorepo

REM Set Python path to include all packages
set PYTHONPATH=C:\Users\alexi\Documents\projects\vba-mcp-monorepo\packages\core\src;C:\Users\alexi\Documents\projects\vba-mcp-monorepo\packages\lite\src;C:\Users\alexi\Documents\projects\vba-mcp-monorepo\packages\pro\src

REM Start the MCP server
python -m vba_mcp_pro.server

REM If server exits, pause to see errors
pause

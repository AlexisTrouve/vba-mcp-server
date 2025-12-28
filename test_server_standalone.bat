@echo off
REM Test VBA MCP Pro Server - Standalone
REM Ce script teste que le serveur MCP se lance correctement

echo ========================================
echo Test VBA MCP Pro Server (Standalone)
echo ========================================
echo.

cd /d %~dp0

REM Set Python path
set PYTHONPATH=%~dp0packages\core\src;%~dp0packages\lite\src;%~dp0packages\pro\src

echo [1/3] Setting environment...
echo PYTHONPATH=%PYTHONPATH%
echo.

echo [2/3] Testing server import...
python -c "from vba_mcp_pro.server import app; print('[OK] Server imported successfully')" 2>test_error.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to import server
    type test_error.txt
    del test_error.txt
    pause
    exit /b 1
)
echo.

echo [3/3] Listing MCP tools...
python -c "import asyncio; from vba_mcp_pro.server import app; tools = asyncio.run(app.list_tools()); print(f'[OK] Server has {len(tools)} tools:'); [print(f'  - {t.name}') for t in tools]" 2>test_error.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to list tools
    type test_error.txt
    del test_error.txt
    pause
    exit /b 1
)
echo.

del test_error.txt 2>nul

echo ========================================
echo [SUCCESS] Server is working!
echo ========================================
echo.
echo Next step: Configure Claude Code
echo See: START_HERE.md
echo.
pause

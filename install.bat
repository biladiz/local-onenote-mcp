@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   Local OneNote MCP Server - Installation
echo ============================================
echo.

:: 1. Check Python
echo [1/4] Checking Python installation...
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Install from https://www.python.org/downloads/
    exit /b 1
)
for /f "tokens=2 delims= " %%v in ('python --version 2^>&1') do set PYVER=%%v
for /f "tokens=1,2 delims=." %%a in ("%PYVER%") do (set PYMAJOR=%%a & set PYMINOR=%%b)
echo        Found Python %PYVER%
if %PYMAJOR% lss 3 ( echo [ERROR] Python 3.10+ required. & exit /b 1 )
if %PYMAJOR% equ 3 if %PYMINOR% lss 10 ( echo [ERROR] Python 3.10+ required. & exit /b 1 )
echo        [OK] Python version is compatible.
echo.

:: 2. Create venv
echo [2/4] Creating virtual environment...
set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%.venv"
if exist "%VENV_DIR%\Scripts\python.exe" (
    echo        Already exists, skipping.
) else (
    python -m venv "%VENV_DIR%"
    if %errorlevel% neq 0 ( echo [ERROR] Failed to create venv. & exit /b 1 )
    echo        [OK] Created at %VENV_DIR%
)
echo.

:: 3. Install deps
echo [3/4] Installing dependencies...
"%VENV_DIR%\Scripts\pip.exe" install --upgrade pip >nul 2>&1
"%VENV_DIR%\Scripts\pip.exe" install -r "%SCRIPT_DIR%requirements.txt"
if %errorlevel% neq 0 ( echo [ERROR] Dependency install failed. & exit /b 1 )
echo        [OK] Dependencies installed.
echo.

:: 4. Generate MCP config
echo [4/4] Generating MCP configuration...
set "PYTHON_PATH=%VENV_DIR%\Scripts\python.exe"
set "MCP_SCRIPT=%SCRIPT_DIR%onenote_pro_mcp_ps.py"
set "JSON_PYTHON=!PYTHON_PATH:\=\\!"
set "JSON_SCRIPT=!MCP_SCRIPT:\=\\!"
(
    echo {
    echo   "mcpServers": {
    echo     "local-onenote-mcp": {
    echo       "command": "!JSON_PYTHON!",
    echo       "args": ["!JSON_SCRIPT!"],
    echo       "disabled": false
    echo     }
    echo   }
    echo }
) > "%SCRIPT_DIR%mcp_config_sample.json"
echo        [OK] Config written to mcp_config_sample.json
echo.

echo ============================================
echo   Done! Next: paste mcp_config_sample.json
echo   into your MCP client config, then restart.
echo ============================================
endlocal

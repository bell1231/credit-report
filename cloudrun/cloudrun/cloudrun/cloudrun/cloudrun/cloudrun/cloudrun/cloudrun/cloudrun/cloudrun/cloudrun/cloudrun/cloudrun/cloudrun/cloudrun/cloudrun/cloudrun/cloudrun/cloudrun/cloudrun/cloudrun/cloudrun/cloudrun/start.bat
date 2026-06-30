@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo   企业授信报告系统 V2 启动中...
echo   企查查 QCC MCP 6大Server
echo ========================================
echo.
echo [1/2] 检查依赖...
if not exist "node_modules" (
    echo 正在安装依赖...
    call npm install
) else (
    echo 依赖已就绪
)
echo.
echo [2/2] 启动服务...
echo.
start http://localhost:3000
node server.js
pause

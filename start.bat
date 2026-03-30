@echo off
chcp 65001 >nul
echo ====================================
echo   舱单数据处理系统
echo ====================================
echo.
echo 正在启动服务器...
echo 服务器地址: http://localhost:5000
echo 按 Ctrl+C 停止服务器
echo.
python app.py

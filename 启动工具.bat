@echo off
rem 启动文档格式转换工具（无黑色命令行窗口）
where pythonw >nul 2>nul && (
    start "" pythonw "%~dp0doc_tool.py"
    exit /b
)
set "PYW=%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe"
if exist "%PYW%" (
    start "" "%PYW%" "%~dp0doc_tool.py"
    exit /b
)
echo 错误：未找到 Python，请先运行"安装依赖.bat"安装 Python 环境
pause

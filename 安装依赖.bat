@echo off
chcp 65001 >nul
echo 正在安装/检查依赖库（首次使用运行一次即可），请稍候...
python -m pip install --upgrade pypandoc-binary xhtml2pdf pymupdf markdown python-docx sv-ttk pillow windnd
echo.
echo 全部完成！以后直接双击桌面上的「文档格式转换工具」即可。
pause

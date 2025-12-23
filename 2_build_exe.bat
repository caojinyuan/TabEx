@echo off
chcp 65001 > nul
echo ======================================
echo TabExplorer 打包工具
echo ======================================
echo.

REM 检查pyinstaller是否安装
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [步骤 1/3] 正在安装 PyInstaller...
    pip install pyinstaller
) else (
    echo [步骤 1/3] PyInstaller 已安装
)

echo.
echo [步骤 2/3] 正在打包程序...
echo 这可能需要几分钟时间，请耐心等待...
echo.

REM 打包成单个exe文件，直接输出到当前目录
pyinstaller --onefile --windowed --name TabExplorer --distpath . TabEx.py

echo.
echo [步骤 3/3] 清理临时文件...
REM 可选：删除build文件夹和spec文件
rmdir /s /q build
del TabExplorer.spec

echo.
echo ======================================
echo 打包完成！
echo 可执行文件位置: TabExplorer.exe
echo ======================================
echo.
echo 注意事项：
echo 1. bookmarks.json 和 pinned_tabs.json 会自动在exe同目录创建
echo 2. 首次运行时会自动创建这些配置文件
echo.
pause

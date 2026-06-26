@echo off
chcp 65001 > nul
echo ======================================
echo 安装 TabExplorer 运行依赖
echo ======================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python 环境，请先安装 Python 3.9 或更高版本。
    pause
    exit /b 1
)

python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请检查上面的错误信息。
    pause
    exit /b 1
)

echo.
echo 依赖安装完成！现在可以双击 1_TabEx.bat 启动程序。
pause



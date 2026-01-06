@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

set VERSION=2.4

echo ======================================
echo TabExplorer v%VERSION% 打包工具
echo ======================================
echo.

REM 检查Python环境
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python环境！
    echo 请先安装Python 3.9或更高版本
    pause
    exit /b 1
)

REM 检查pyinstaller是否安装
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [步骤 1/4] 正在安装 PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller 安装失败！
        pause
        exit /b 1
    )
) else (
    echo [步骤 1/4] PyInstaller 已安装
)

echo.
echo [步骤 2/4] 清理旧文件...
if exist TabExplorer.exe (
    echo 删除旧的 TabExplorer.exe
    del /q TabExplorer.exe
)
if exist build (
    echo 删除旧的 build 目录
    rmdir /s /q build
)
if exist dist (
    echo 删除旧的 dist 目录
    rmdir /s /q dist
)
if exist *.spec (
    echo 删除旧的 spec 文件
    del /q *.spec
)

echo.
echo [步骤 3/4] 正在打包程序...
echo 这可能需要几分钟时间，请耐心等待...
echo.

REM 打包成单个exe文件，直接输出到当前目录
pyinstaller --onefile --windowed --name TabExplorer --distpath . TabEx.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败！请检查错误信息
    pause
    exit /b 1
)

echo.
echo [步骤 4/4] 清理临时文件...
if exist build rmdir /s /q build
if exist TabExplorer.spec del /q TabExplorer.spec

echo.
echo ======================================
echo 打包完成！版本: v%VERSION%
echo ======================================
echo.
echo 可执行文件: TabExplorer.exe
if exist TabExplorer.exe (
    for %%A in (TabExplorer.exe) do echo 文件大小: %%~zA 字节
)
echo.
echo 注意事项：
echo 1. bookmarks.json 和 pinned_tabs.json 会自动在exe同目录创建
echo 2. 首次运行时会自动创建这些配置文件
echo 3. 如需发布，建议配合 README.md 一起打包
echo.
echo 下一步：
echo - 测试: 双击 TabExplorer.exe 运行
echo - 发布: 上传到 GitHub Releases
echo.
pause

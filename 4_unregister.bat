@echo off
chcp 65001 > nul
echo ======================================
echo 取消 TabExplorer 系统注册
echo ======================================
echo.
echo 此操作将移除右键菜单中的 TabExplorer 选项
echo.
pause

echo.
echo 正在从注册表移除...
echo.

REM 创建临时注册表文件
set "REG_FILE=%TEMP%\TabExplorer_unregister.reg"

(
echo Windows Registry Editor Version 5.00
echo.
echo ; 移除文件夹右键菜单
echo [-HKEY_CLASSES_ROOT\Directory\shell\TabExplorer]
echo [-HKEY_CLASSES_ROOT\Directory\shell\openwithTabExplorer]
echo.
echo ; 移除文件右键菜单
echo [-HKEY_CLASSES_ROOT\*\shell\TabExplorer]
echo.
echo ; 移除文件夹背景右键菜单
echo [-HKEY_CLASSES_ROOT\Directory\Background\shell\TabExplorer]
echo [-HKEY_CLASSES_ROOT\Directory\Background\shell\openwithTabExplorer]
echo [-HKEY_CLASSES_ROOT\Directory\Background\shell\openTabExplorerhere]
echo.
echo ; 移除驱动器右键菜单
echo [-HKEY_CLASSES_ROOT\Drive\shell\TabExplorer]
echo [-HKEY_CLASSES_ROOT\Drive\shell\openwithTabExplorer]
) > "%REG_FILE%"

REM 导入注册表
regedit /s "%REG_FILE%"

if %errorlevel% equ 0 (
    echo.
    echo ======================================
    echo 移除成功！
    echo ======================================
    echo.
    echo TabExplorer 已从右键菜单中移除
    echo.
) else (
    echo.
    echo 移除失败！请以管理员身份运行此脚本
    echo.
)

REM 清理临时文件
del "%REG_FILE%"

pause

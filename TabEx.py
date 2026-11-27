"""TabExplorer 启动入口；界面与功能代码位于 tabexplorer/ 目录。"""

# 版本号唯一来源：程序、打包脚本 2_build_exe.bat 与旧版本的检查更新都读取这一行。
APP_VERSION = "4.22"


if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == '--preview':
        from tabexplorer.preview import main
    else:
        from tabexplorer.app import main
    main()

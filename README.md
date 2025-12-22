# TabEx
Add Chrome-like tab functionality to explorer, replacing Clover and ExTab.

实现功能：
1.书签管理
2.标签功能，标签固定
3.导航树
4.双击空白处，回退上一级目录
5.拖拽文件夹到标签区，自动打开文件夹
6.生成独立运行文件 TabExplorer.exe
7.添加面包屑导航栏


运行环境：
首次运行前需要双击 0_install_requirements.bat ， 安装必要软件。
然后双击运行 1_TabEx.bat 。

开机自启动方式：
将 1_TabEx.bat 的快捷方式放到 shell:startup 目录

我本地环境：
python 3.9.6

打包步骤：
pip install pyinstaller
pyinstaller --onefile --windowed --name TabExplorer TabEx.py
或者
2_build_exe.bat

# TabEx
Add Chrome-like tab functionality to explorer, replacing Clover and ExTab.

实现功能：
1.书签管理
2.标签功能，标签固定
3.导航树
4.双击空白处，回退上一级目录
5.拖拽文件夹到标签区，自动打开文件夹


界面布局（中文说明）

- 窗口：主窗口为 `MainWindow`，主体是左右分割的 `QSplitter`。
- 左侧导航：`QTreeView`（`dir_tree`）+ `QFileSystemModel`，显示盘符与目录，宽度有限制（min/max）。
- 右侧内容：垂直布局包含 `path_bar`（路径栏，`QLineEdit`）、嵌入的原生 Explorer (`QAxWidget` 控件 `Shell.Explorer`)、隐藏的 `drive_list`（盘符列表，`QListWidget`）以及底部的 `blank` 空白条（`QLabel`）。
- 标签区：`QTabWidget`（`tab_widget`），每个标签为 `FileExplorerTab`，支持关闭、固定（pin）、右键菜单（固定、添加书签等）。角落有 `⬆️`（上一级）和 `➕`（新建标签）按钮。
- 菜单与书签：顶部 `menuBar()` 由 `BookmarkManager` 从 `bookmarks.json` 生成书签菜单，包含“书签管理”。
- 路径同步：通过定时器读取 `QAxWidget` 的 `LocationURL` 来同步 `current_path`、`path_bar` 并触发左侧目录树展开。
- 返回上一级的触发源：路径栏双击、底部 `blank` 双击、上一级按钮、标签区角落双击，以及程序化调用。为避免误触（双击文件/文件夹同时触发返回），实现了多层判断：按下时读取 ActiveX `SelectedItems()`、延迟检查选中数、在 Windows 下优先用原生 ListView HitTest（需 `pywin32`），并对比导航前后路径；只有确认为空白位置时才执行“返回上一级”。

（以上为界面与交互的简要说明，已同步到代码 `TabEx.py` 中的相关实现。）


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

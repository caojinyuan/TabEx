# TabEx
Add Chrome-like tab functionality to explorer, replacing Clover and ExTab.

## 📥 下载

**最新版本**: 从 [Releases](../../releases/latest) 页面下载 `TabExplorer.exe`

无需安装Python环境，下载即用！

---

## ✨ 特色功能
✅ 多标签页浏览（每个标签页独立历史记录）
✅ 前进/后退导航（支持浏览历史记录）
✅ 标签页固定（带 📌 图标，开机自动恢复）
✅ 标签页拖拽排序（鼠标拖拽调整标签页顺序，固定标签保持在最左侧）
✅ 拖拽文件夹打开新标签页
✅ 书签管理（支持文件夹层级、上移/下移、编辑/删除）
✅ 面包屑导航（支持点击层级跳转、双击编辑）
✅ 网络路径支持（UNC 路径 \\server\share）
✅ 特殊 shell 路径支持（回收站、桌面、此电脑等）
✅ 左侧目录树（自动展开本地盘符，隐藏网络驱动器）
✅ 双击空白区域智能操作（标签栏空白→新建标签页，文件区空白→返回上级）
✅ **单实例模式**（右键菜单打开文件夹时在已有窗口中打开新标签页，避免启动多个进程）
✅ **强大的搜索功能**（支持文件名、文件内容、文件夹名搜索）

注册为右键菜单：
运行 `3_register_as_default.bat` 注册到资源管理器右键菜单（需管理员权限）
支持：
- 右键文件夹 → "open with TabExplorer"
- 右键文件 → "open with TabExplorer"（打开文件所在文件夹）
- 文件夹空白处右键 → "open TabExplorer here"
- 右键驱动器 → "open with TabExplorer"

卸载右键菜单：运行 `4_unregister.bat`

运行环境：
首次运行前需要双击 0_install_requirements.bat ， 安装必要软件。
然后双击运行 1_TabEx.bat 。

开机自启动方式：
将 1_TabEx.bat 的快捷方式放到 shell:startup 目录

我本地环境：
python 3.9.6






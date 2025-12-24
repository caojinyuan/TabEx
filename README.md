# TabEx
Add Chrome-like tab functionality to explorer, replacing Clover and ExTab.

## 📥 下载

**最新版本**: 从 [Releases](../../releases/latest) 页面下载 `TabExplorer.exe`

无需安装Python环境，下载即用！

---

## ✨ 特色功能
✅ **完整的Windows资源管理器体验**（每个标签页嵌入真实的explorer.exe窗口，包含完整的导航窗格、地址栏、工具栏、搜索框等所有原生功能）
✅ 多标签页浏览（每个标签页独立路径）
✅ 标签页固定（带 📌 图标，开机自动恢复）
✅ 标签页拖拽排序（鼠标拖拽调整标签页顺序，固定标签保持在最左侧）
✅ 拖拽文件夹打开新标签页
✅ 书签管理（支持文件夹层级、上移/下移、编辑/删除）
✅ 默认书签（🖥️此电脑、🗔桌面、🗑️回收站、🚀启动项、⬇️下载）
✅ 网络路径支持（UNC 路径 \\server\share）
✅ 特殊 shell 路径支持（回收站、桌面、此电脑等）
✅ 双击标签栏空白区域新建标签页
✅ **单实例模式**（右键菜单打开文件夹时在已有窗口中打开新标签页，避免启动多个进程）
✅ **强大的搜索功能**（支持文件名、文件内容、文件夹名搜索，支持文件类型过滤）
✅ **自动接管资源管理器**（默认开启，类似 Clover，自动将新打开的资源管理器窗口转为 TabExplorer 标签页）
✅ **设置界面**（可配置接管资源管理器功能，设置持久化保存）

---

## 🚀 完整的Windows资源管理器体验

TabExplorer 最大的特色是将**真正的Windows资源管理器**嵌入到每个标签页中！

**与传统文件管理器的区别**：
- ❌ 传统方案：使用简化的ActiveX控件，功能受限
- ✅ TabExplorer：嵌入完整的explorer.exe窗口，100%原生体验

**每个标签页都包含完整的Explorer功能**：
- 📂 左侧导航窗格（快速访问、OneDrive、此电脑等）
- 🔍 顶部搜索框（实时搜索）
- 📍 地址栏（快速导航、路径复制）
- 🎨 工具栏/Ribbon界面（所有Windows资源管理器功能）
- 👁️ 各种视图模式（大图标、列表、详细信息等）
- 🏷️ Git Tortoise等覆盖图标完美支持

---

## 🚀 自动接管资源管理器

TabExplorer **默认自动启动**接管功能，像 Clover 一样将系统资源管理器窗口转为标签页！

**工作原理**：
- 自动检测新打开的Explorer窗口
- 直接嵌入到TabExplorer的新标签页中
- 原窗口无缝转换，不丢失任何内容

**控制方式**：
- 菜单栏 **"⚙️ 设置"** → **"接管资源管理器"** 复选框可随时开关
- 设置自动保存到 config.json，重启后保持
- 默认已启用，无需手动操作

**适用场景**：
- 按 `Win + E` 打开资源管理器
- 双击桌面文件夹
- 从开始菜单打开文件夹
- VSCode等软件跳转打开文件夹
- 任何方式打开的Explorer窗口

---

## 📋 注册为右键菜单

运行 `3_register_as_default.bat` 注册到资源管理器右键菜单（需管理员权限）
支持：
- 右键文件夹 → "open with TabExplorer"
- 右键文件 → "open with TabExplorer"（打开文件所在文件夹）
- 文件夹空白处右键 → "open TabExplorer here"
- 右键驱动器 → "open with TabExplorer"

卸载右键菜单：运行 `4_unregister.bat`

---

## 🏃 运行说明

运行环境：
python 3.9.6
首次运行前需要双击 `0_install_requirements.bat` ， 安装必要软件。
然后双击运行 `1_TabEx.bat` 。

开机自启动方式：
将 `1_TabEx.bat` 的快捷方式放到 shell:startup 目录








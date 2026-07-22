# 全局搜索缓存（LRU缓存，最多缓存50个搜索结果）

from collections import OrderedDict
import hashlib
import os

# 应用版本号（单一来源）：窗口标题与打包脚本 2_build_exe.bat 均引用此处。
# 修改版本时只改这一行；2_build_exe.bat 会自动解析。
APP_VERSION = "3.62"


# TabEx i18n module
# Usage: tr("中文") returns English when language is "en", else returns the Chinese original.
# This module is inserted near the top of TabEx.py before any UI code.

_app_language = "zh"  # default; overridden from config.json at startup

def _set_app_language(lang):
    global _app_language
    _app_language = lang if lang in ("zh", "en") else "zh"

def tr(zh_text):
    """Return translated string. Falls back to zh_text if no translation found."""
    if _app_language != "en":
        return zh_text
    return _LANG_EN.get(zh_text, zh_text)

_LANG_EN = {
    # ── Common buttons / actions ──────────────────────────────────────────
    "关闭": "Close",
    "后退": "Back",
    "前进": "Forward",
    "关闭标签页": "Close Tab",
    "新建标签页": "New Tab",
    "上级目录": "Parent Folder",
    "恢复关闭的标签页": "Reopen Closed Tab",
    "刷新": "Refresh",
    "确定": "OK",
    "取消": "Cancel",
    "删除": "Delete",
    "复制": "Copy",
    "编辑": "Edit",
    "保存": "Save",
    "设置": "Settings",
    "上移": "Move Up",
    "下移": "Move Down",
    "重命名": "Rename",
    "无": "None",
    "成功": "Success",
    "错误": "Error",
    "提示": "Notice",
    "不支持": "Not Supported",
    "警告": "Warning",
    "格式错误": "Format Error",
    "已打开": "Opened",
    "已停止": "Stopped",
    "已删除": "Deleted",
    "已启动": "Started",
    "已保存": "Saved",
    "已选择": "Selected",  # alias
    "未选择": "Not Selected",
    "保存失败": "Save Failed",
    "导入失败": "Import Failed",
    "导出失败": "Export Failed",
    "导入成功": "Imported",
    "导出成功": "Exported",
    "添加失败": "Add Failed",
    "打开失败": "Open Failed",
    # ── Dialogs ───────────────────────────────────────────────────────────
    "书签": "Bookmarks",
    "书签管理": "Bookmarks",
    "书签管理器": "Bookmark Manager",
    "书签名称": "Bookmark Name",
    "书签栏": "Bookmark Bar",
    "名称": "Name",
    "路径": "Path",
    "路径错误": "Path Error",
    "选择匹配项": "Select Match",
    "类型": "Type",
    "完整路径": "Full Path",
    "文件夹": "Folder",
    "文件": "File",
    # ── Search dialog ─────────────────────────────────────────────────────
    "文件名": "File Name",
    "修改日期": "Modified",
    "大小": "Size",
    "搜索:": "Search:",
    "搜索": "Search",
    "搜索路径:": "Search Path:",
    "文件类型:": "File Type:",
    "就绪": "Ready",
    "🔍 搜索": "🔍 Search",
    "⏹ 停止": "⏹ Stop",
    "搜索文件名": "Search File Name",
    "搜索文件内容": "Search Content",
    "区分大小写": "Case Sensitive",
    "全词匹配": "Whole Word",
    "使用 Everything (极速)": "Use Everything (Ultra-fast)",
    "轻量模式(更快)": "Lightweight (Faster)",
    "输入搜索关键词...": "Enter search keyword...",
    "输入要搜索的文件夹路径...": "Enter folder path to search...",
    "例如: *.c,*.h,*.xml (留空表示搜索所有类型)": "e.g. *.c,*.h,*.xml (empty = all types)",
    "区分大小写匹配文件名和文件内容": "Case-sensitive match for filename and content",
    "仅匹配完整单词，避免命中更长字符串的一部分": "Whole-word match only, avoid partial string hits",
    "勾选后搜索结果将优先显示核心路径信息，可能省略修改时间/大小": "Prioritize path info; may omit modified date/size",
    "使用Everything搜索引擎\n路径: {}\n只搜索文件名，速度极快": "Using Everything engine\nPath: {}\nFilename search only, blazing fast",
    "未检测到Everything，请从 https://www.voidtools.com/ 下载安装": "Everything not found. Install from https://www.voidtools.com/",
    "搜索 - {}": "Search - {}",
    "正在加载缓存结果...": "Loading cached results...",
    "已停止": "Stopped",
    "列举文件中{}... (最多显示{}个结果)": "Listing files {}... (max {} results)",
    "搜索中... (最多显示{}个结果)": "Searching... (max {} results)",
    "搜索中... 已扫描 {} 个文件，找到 {} 个结果": "Searching... scanned {} files, {} found",
    "搜索完成（缓存），共显示 {} 个结果": "Search complete (cache), {} results",
    "搜索完成，共找到 {} 个结果（扫描了 {} 个文件）": "Search complete, {} results ({} files scanned)",
    "搜索完成（已限制），显示前 {} 个结果（扫描了 {} 个文件）⚠️": "Search limited, showing first {} results ({} files scanned) ⚠️",
    "Everything搜索完成，共找到 {} 个结果{}": "Everything search done, {} results{}",
    # ── Search log messages ───────────────────────────────────────────────
    "[Search] 开始搜索路径: {}": "[Search] Start path: {}",
    "[Search] 搜索关键词: {}": "[Search] Keyword: {}",
    "[Search] 搜索文件名: {}, 搜索内容: {}": "[Search] File name: {}, Content: {}",
    "[Search] 文件类型过滤: {}": "[Search] File type filter: {}",
    "[Search] 搜索完成，共扫描 {} 个文件，找到 {} 个结果": "[Search] Done, scanned {} files, {} results",
    "[Search] 搜索被中断": "[Search] Search interrupted",
    "[Search] 搜索被中断（文件循环）": "[Search] Interrupted (file loop)",
    "[Search] 搜索缓存已清除": "[Search] Cache cleared",
    "[Search] 文件被类型过滤跳过: {}": "[Search] File skipped by type filter: {}",
    "[Search] 无法读取文件 {}: {}": "[Search] Cannot read file {}: {}",
    "[Search] 读取文件失败 {} (编码 {}): {}": "[Search] Read failed {} (encoding {}): {}",
    "[Search] 跳过 {} 个二进制文件（不搜索内容）": "[Search] Skipped {} binary files",
    "[Search] ⚠️ 队列溢出 {} 次（部分结果未显示）": "[Search] ⚠️ Queue overflow {} times (some results hidden)",
    "[Search] 搜索缓存已清除": "[Search] Search cache cleared",
    '[Search] UI更新已调度（使用队列）': '[Search] UI update queued',
    '[Search] ⚠️ 队列满，最终状态更新失败': '[Search] ⚠️ Queue full, final status update failed',
    'Using Everything搜索引擎...': 'Using Everything search engine...',
    # ── Context menu (search results) ────────────────────────────────────
    "打开": "Open",
    "打开所在目录": "Open Containing Folder",
    "用系统默认程序打开": "Open with Default App",
    "用记事本打开": "Open with Notepad",
    "用 Notepad++ 打开": "Open with Notepad++",
    "选择其他应用...": "Open with...",
    "未检测到 Notepad++": "Notepad++ not detected",
    "该搜索结果不是可打开的文件": "This result is not an openable file",
    "当前系统不支持打开\u201c选择其他应用\u201d对话框": "This system does not support 'Open with' dialog",
    "无法使用 {} 打开文件: {}": "Cannot open file with {}: {}",
    "无法使用系统默认程序打开文件: {}": "Cannot open file with default app: {}",
    '无法打开\u201c选择其他应用\u201d对话框: {}': "Cannot open 'Open with' dialog: {}",
    "当前选中项不可打开": "Selected item cannot be opened",
    "当前选中项不是可打开的文件": "Selected item is not an openable file",
    # ── BookmarkDialog ────────────────────────────────────────────────────
    "名称": "Name",
    "路径": "Path",
    "关闭": "Close",
    "无可用书签文件夹": "No bookmark folders available",
    "请先在 bookmarks.json 中添加至少一个文件夹。": "Add at least one folder in bookmarks.json first.",
    "选择书签文件夹": "Select Bookmark Folder",
    "请选择父文件夹：": "Select parent folder:",
    "请输入书签名称：": "Enter bookmark name:",
    "未能添加书签，请检查父文件夹。": "Failed to add bookmark. Check parent folder.",
    "不支持的书签": "Unsupported Bookmark",
    "🗑️ 删除书签": "🗑️ Delete Bookmark",
    "已删除": "Deleted",
    "书签 '{}' 已删除": "Bookmark '{}' deleted",
    "书签已保存": "Bookmarks saved",
    "书签已成功导出到:\n{}": "Bookmarks exported to:\n{}",
    "成功导入 {} 个书签项": "Imported {} bookmark items",
    "选中的书签/文件夹已删除": "Selected bookmark/folder deleted",
    "请先选择要上移的书签或文件夹。": "Select a bookmark or folder to move up.",
    "请先选择要下移的书签或文件夹。": "Select a bookmark or folder to move down.",
    "请先选择要删除的书签或文件夹。": "Select a bookmark or folder to delete.",
    "请先选择要编辑的书签或文件夹。": "Select a bookmark or folder to edit.",
    "请先选择要重命名的书签或文件夹。": "Select a bookmark or folder to rename.",
    "暂不支持打开此类型书签: {}": "Cannot open bookmark type: {}",
    "请输入文件夹名称：": "Enter folder name:",
    "请输入新名称：": "Enter new name:",
    "请输入新路径：": "Enter new path:",
    "编辑书签": "Edit Bookmark",
    "编辑文件夹": "Edit Folder",
    "保存当前书签顺序和层级": "Save current bookmark order and hierarchy",
    "💾 保存": "💾 Save",
    "📑 添加书签": "📑 Add Bookmark",
    "📤 导出": "📤 Export",
    "📥 导入": "📥 Import",
    "导入书签": "Import Bookmarks",
    "导出书签": "Export Bookmarks",
    "从JSON文件导入书签": "Import Bookmarks from JSON",
    "导出书签到JSON文件": "Export Bookmarks to JSON",
    "导入方式": "Import Method",
    "已自动选择合并模式，将导入内容追加到现有书签。": "Auto-selected merge mode; imported bookmarks appended.",
    "导入的文件不是有效的JSON格式": "File is not valid JSON",
    "导入的文件中没有书签内容": "No bookmark content in file",
    "导入的文件格式不正确，必须包含 'bookmark_bar' 节点": "Invalid format; must contain 'bookmark_bar' node",
    # ── QuickFindResultsDialog ────────────────────────────────────────────
    "选择匹配项": "Select Match",
    "名称": "Name",
    "类型": "Type",
    "完整路径": "Full Path",
    "确定": "OK",
    "取消": "Cancel",
    "文件夹": "Folder",
    "文件": "File",
    "切换到同级文件夹": "Switch to sibling folder",
    "无子文件夹": "No subfolders",
    # ── SettingsDialog ────────────────────────────────────────────────────
    "路径栏分隔符设置": "Path Bar Separator",
    "路径栏拷贝分隔符:": "Copy Separator:",
    "设置从路径栏拷贝时使用的分隔符": "Separator used when copying from path bar",
    "Explorer监听设置": "Explorer Monitor",
    "监听新Explorer窗口": "Monitor New Explorer Windows",
    "状态栏显示 CPU/内存占用": "Show CPU/Memory in Status Bar",
    "内存": "Mem",
    "在状态栏右侧实时显示整机 CPU 与内存占用，每2秒刷新": "Show system-wide CPU and memory on the right of the status bar, refreshed every 2s",
    "监听间隔（秒）:": "Monitor Interval (s):",
    "检查新Explorer窗口的时间间隔，更长的间隔降低CPU占用": "Interval for checking new Explorer windows; longer = less CPU",
    "（推荐: 2.0秒）": "(Recommended: 2.0s)",
    "调试设置": "Debug Settings",
    "启用调试输出（输出到终端）": "Enable Debug Output (to terminal)",
    "启用后将在终端输出调试信息，用于开发和问题排查": "Outputs debug info to terminal for development",
    "启用 Explorer Monitor 调试输出": "Enable Explorer Monitor Debug Output",
    "单独控制 Explorer Monitor 的日志输出（需要先启用调试输出）": "Control Explorer Monitor logging (requires debug output enabled)",
    "启用资源快照日志": "Enable Resource Snapshot Log",
    "定时写入 runtime_health.log，用于观察长期运行时的内存和线程趋势": "Periodically write runtime_health.log for long-term monitoring",
    "资源快照间隔（分钟）:": "Snapshot Interval (min):",
    "资源快照日志写入周期，建议 5 分钟或更长": "Snapshot write period; 5 min or more recommended",
    "分钟": "min",
    "打开资源日志": "Open Resource Log",
    "打开 runtime_health.log；如果日志尚未生成，则打开所在目录": "Open runtime_health.log, or its directory if not yet created",
    "标签页设置": "Tab Settings",
    "关闭时缓存当前标签页，下次启动时恢复": "Cache tabs on close, restore on next launch",
    "关闭软件时保存非固定标签，下次启动时自动恢复（不包括固定标签）": "Save non-pinned tabs on close; restore at next launch",
    "鼠标手势设置": "Mouse Gestures",
    "启用鼠标手势（按住右键画线）": "Enable mouse gestures (hold right button and draw)",
    "在文件区域按住鼠标右键画线即可触发导航操作；关闭后右键恢复为普通右键菜单": "Hold the right mouse button in the file area and draw to trigger navigation; disabling restores the normal right-click menu",
    "← 向左：后退\n→ 向右：前进\n↓ 向下：关闭当前标签页\n↑ 向上：打开新标签页\n↑↓ 上下：刷新\n↓↑ 下上：返回上级目录\n↓→ 下右：恢复关闭的标签页": "← Left: Back\n→ Right: Forward\n↓ Down: Close current tab\n↑ Up: Open new tab\n↑↓ Up-Down: Refresh\n↓↑ Down-Up: Go to parent folder\n↓→ Down-Right: Reopen closed tab",
    "开机启动设置": "Auto-start",
    "开机自动启动 TabExplorer": "Auto-start TabExplorer at login",
    "在 Windows 启动时自动运行 TabExplorer.exe": "Run TabExplorer.exe automatically at Windows startup",
    "Git 工具设置": "Git Tools",
    "显示 TortoiseGit 快捷按钮（标题栏）": "Show TortoiseGit buttons (title bar)",
    "在标题栏显示 Git Log 和 Git Commit 快捷按钮": "Show Git Log and Commit buttons in title bar",
    "默认终端:": "Default Terminal:",
    "路径栏输入 terminal 或 term 时使用的默认终端": "Default terminal when typing 'terminal' in the path bar",
    "快捷方式设置": "Launcher Settings",
    "启用标题栏启动区（可拖拽应用、快捷方式或脚本并点击启动）": "Enable title bar launcher (drag apps/shortcuts/scripts to launch)",
    "拖拽应用、快捷方式或脚本到标题栏 Git 左侧区域，后续可一键启动": "Drag apps/shortcuts/scripts to the title bar for one-click launch",
    "快捷键设置": "Hotkey Settings",
    "Ctrl+T - 新建标签页": "Ctrl+T - New Tab",
    "Ctrl+W - 关闭当前标签页": "Ctrl+W - Close Tab",
    "Ctrl+Shift+T - 恢复关闭的标签页": "Ctrl+Shift+T - Reopen Tab",
    "Ctrl+Tab / Ctrl+Shift+Tab - 切换标签页": "Ctrl+Tab / Ctrl+Shift+Tab - Switch Tab",
    "Ctrl+F - 打开搜索对话框": "Ctrl+F - Open Search",
    "Ctrl+G - 检索当前目录文件夹/文件名": "Ctrl+G - Quick Find in Current Dir",
    "Alt+Left/Right - 前进/后退": "Alt+Left/Right - Back/Forward",
    "Alt+Up - 返回上级目录": "Alt+Up - Go Up",
    "F5 - 刷新当前路径": "F5 - Refresh",
    "Ctrl+D - 添加当前路径到书签": "Ctrl+D - Add Bookmark",
    "Alt+Z - 复制选中文件名（含后缀）": "Alt+Z - Copy File Name (with ext)",
    "Alt+X - 复制文件路径\\文件名": "Alt+X - Copy File Path",
    "💡 提示：取消勾选可禁用对应的快捷键": "💡 Tip: Uncheck to disable a shortcut",
    "常规": "General",
    "手势与快捷键": "Gestures & Hotkeys",
    "工具集成": "Tools",
    "AI 助手": "AI Assistant",
    "高级": "Advanced",
    "AI 助手设置": "AI Assistant Settings",
    "启用 AI 助手（显示标题栏机器人按钮🤖）": "Enable AI Assistant (show 🤖 button in title bar)",
    "── 请选择预设服务商 ──": "── Select AI Provider ──",
    "Groq（免费·极速·推荐）": "Groq (Free · Ultra-fast · Recommended)",
    "SiliconFlow 硅基流动（免费额度·国内快）": "SiliconFlow (Free quota · Fast in China)",
    "DeepSeek（注册送额度·中文强）": "DeepSeek (Free credits · Strong Chinese)",
    "Google Gemini（免费版）": "Google Gemini (Free tier)",
    "OpenRouter（含永久免费模型）": "OpenRouter (Includes free models)",
    "本地 LM Studio（无需Key）": "Local LM Studio (No key required)",
    "本地 Ollama（无需Key）": "Local Ollama (No key required)",
    "── 自定义（手动填写下方） ──": "── Custom (fill in manually) ──",
    "免费注册获取Key: https://console.groq.com/keys": "Get free key: https://console.groq.com/keys",
    "免费注册获取Key: https://cloud.siliconflow.cn": "Get free key: https://cloud.siliconflow.cn",
    "注册获取Key: https://platform.deepseek.com/api_keys": "Get key: https://platform.deepseek.com/api_keys",
    "免费获取Key: https://aistudio.google.com/app/apikey": "Get free key: https://aistudio.google.com/app/apikey",
    "注册获取Key: https://openrouter.ai/keys": "Get key: https://openrouter.ai/keys",
    "启动 LM Studio → Local Server 后使用": "Start LM Studio → Local Server, then use",
    "安装 Ollama 并运行模型后使用": "Install Ollama, run a model, then use",
    "快速选择:": "Quick Select:",
    "选择预设后自动填充地址和模型名，然后只需粘贴对应的 API Key": "Select preset to auto-fill URL and model; just paste the API key",
    "API 地址:": "API URL:",
    "例: https://api.groq.com/openai/v1": "e.g. https://api.groq.com/openai/v1",
    "填写 OpenAI 兼容 API 的基础地址（不含 /chat/completions）": "OpenAI-compatible API base URL (without /chat/completions)",
    "API 密钥:": "API Key:",
    "粘贴从服务商网站获取的 Key（本地模型可留空）": "Paste the API key (leave empty for local models)",
    "模型名称:": "Model:",
    "例: llama-3.3-70b-versatile": "e.g. llama-3.3-70b-versatile",
    "系统提示词（留空使用默认）:": "System Prompt (empty = built-in default):",
    "留空则使用内置提示词（支持 [OPEN_DIR:] 和 [RUN_SCRIPT:] 指令）": "Leave empty to use built-in prompt",
    "面板宽度 (px):": "Panel Width (px):",
    "点击链接在浏览器中打开 GitHub Releases 页面": "Click to open GitHub Releases in browser",
    '检查更新: <a href="https://github.com/caojinyuan/TabEx/releases">https://github.com/caojinyuan/TabEx/releases</a>':
        'Check for updates: <a href="https://github.com/caojinyuan/TabEx/releases">https://github.com/caojinyuan/TabEx/releases</a>',
    "语言 / Language:": "语言 / Language:",
    # ── Auto-start toasts ─────────────────────────────────────────────────
    "已启用开机自动启动": "Auto-start enabled",
    "已禁用开机自动启动": "Auto-start disabled",
    "设置开机启动失败: {}": "Failed to set auto-start: {}",
    "未找到 TabExplorer.exe，请确保程序已正确安装": "TabExplorer.exe not found. Please ensure it is installed.",
    # ── Resource log ──────────────────────────────────────────────────────
    "资源日志": "Resource Log",
    "日志尚未生成，已打开日志目录": "Log not yet generated; opened log directory",
    "日志目录不存在": "Log directory does not exist",
    "无法打开资源日志: {}": "Cannot open resource log: {}",
    # ── Toast messages ────────────────────────────────────────────────────
    "请至少选择一种搜索类型": "Please select at least one search type",
    "请输入搜索路径": "Please enter a search path",
    "不支持搜索特殊路径（shell:）": "Special shell: paths are not supported for search",
    "路径不存在: {}": "Path does not exist: {}",
    "路径不存在:\n{}": "Path does not exist:\n{}",
    "路径不是文件夹:\n{}": "Path is not a folder:\n{}",
    "控制面板已在新窗口打开": "Control Panel opened in new window",
    "无法打开控制面板: {}": "Cannot open Control Panel: {}",
    "当前路径无效": "Invalid current path",
    "当前目录不是 Git 仓库，未找到 .git": "Current dir is not a Git repo (.git not found)",
    "未找到 TortoiseGit，请确认已安装 TortoiseGit\n下载地址: https://tortoisegit.org/download/": "TortoiseGit not found.\nDownload: https://tortoisegit.org/download/",
    "无法打开 TortoiseGit Log: {}": "Cannot open TortoiseGit Log: {}",
    "无法打开 TortoiseGit Commit: {}": "Cannot open TortoiseGit Commit: {}",
    "当前路径无效，无法启动终端": "Invalid path; cannot open terminal",
    "未找到 Git Bash，请确认已安装 Git for Windows": "Git Bash not found. Install Git for Windows.",
    "未找到可用的 Git Bash 可执行文件": "No Git Bash executable found",
    "无法打开 Git Bash: {}": "Cannot open Git Bash: {}",
    "当前系统不支持打开计算器": "Calculator is not supported on this system",
    "无法打开计算器: {}": "Cannot open calculator: {}",
    "当前路径无效，无法打开默认终端": "Invalid path; cannot open default terminal",
    "当前路径无效，无法打开命令行": "Invalid path; cannot open command line",
    "当前路径无效，无法打开 PowerShell": "Invalid path; cannot open PowerShell",
    "当前路径无效，无法打开 Git Bash": "Invalid path; cannot open Git Bash",
    "无法打开 cmd: {}": "Cannot open cmd: {}",
    "无法打开 PowerShell: {}": "Cannot open PowerShell: {}",
    "无法打开命令行: {}": "Cannot open command line: {}",
    "无法打开默认终端: {}": "Cannot open default terminal: {}",
    "无法打开浏览器: {}": "Cannot open browser: {}",
    "当前为特殊路径，无法定位到 cmd 或 PowerShell": "Special path; cannot open cmd or PowerShell",
    "正在加载大文件夹...": "Loading large folder...",
    "复制成功": "Copied",
    "文件名: {}": "File name: {}",
    "路径: {}": "Path: {}",
    "未选中文件，也无法获取路径栏地址": "No file selected and path bar address unavailable",
    "当前没有可用的标签页路径": "No active tab path available",
    "当前路径不支持快捷检索": "Path does not support quick find",
    "当前目录无效": "Invalid current directory",
    "快捷定位": "Quick Find",
    "请输入要检索的文件或文件夹关键字：": "Enter keyword to search for file or folder:",
    "请输入搜索关键词": "Please enter a keyword",
    "快捷检索失败: {}": "Quick find failed: {}",
    "当前目录下未找到包含\"{}\"的文件或文件夹名": "No file/folder matching \"{}\" in current directory",
    "当前标签自动刷新已冻结": "Auto-refresh frozen for current tab",
    "当前标签自动刷新已恢复": "Auto-refresh resumed for current tab",
    "当前标签页路径无效": "Current tab path is invalid",
    "设置已更新": "Settings Updated",
    "无法嵌入该窗口，已尝试用系统资源管理器打开。\n{}": "Cannot embed window; opened with Explorer instead.\n{}",
    "无法启动快捷方式: {}": "Cannot launch shortcut: {}",
    "无法打开选中项: {}": "Cannot open selection: {}",
    "快捷方式不存在，已从列表移除": "Shortcut not found; removed from list",
    "拖拽保存失败: {}": "Drag-save failed: {}",
    "检查更新": "Check for Updates",
    "已在浏览器中打开更新页面": "Update page opened in browser",
    "请先打开一个文件夹": "Please open a folder first",
    "已在当前目录选中{}: {}": "Selected {} in current dir: {}",
    "启动文件夹不存在: {}": "Startup folder does not exist: {}",
    "Explorer窗口监听已{}\n{}": "Explorer window monitoring {}\n{}",
    # ── Status bar ────────────────────────────────────────────────────────
    "就绪": "Ready",
    "共 {} 项": "{} items",
    "（共 {} 项）": "({} items)",
    "已选 {} 项{}": "Selected {}{}",
    "，已省略大小统计": ", size stats omitted",
    "，修改时间 {}": ", modified {}",
    "分支: {}": "Branch: {}",
    '✔ 无更改': '✔ No changes',
    '<span style="color:#2e7d32;font-weight:bold">✔ 无更改</span>': '<span style="color:#2e7d32;font-weight:bold">✔ No changes</span>',
    "暂存(Add)": "Staged (Add)",
    "修改(Commit)": "Modified (Commit)",
    "未跟踪(待Add)": "Untracked (Pending Add)",
    "，元数据降级 {} 条": ", {} items degraded",
    # ── Main window toolbar / menus ───────────────────────────────────────
    "后退 (Alt+←)": "Back (Alt+←)",
    "前进 (Alt+→)": "Forward (Alt+→)",
    "新建标签页 (Ctrl+T)": "New Tab (Ctrl+T)",
    "恢复关闭的标签页 (Ctrl+Shift+T)": "Reopen Tab (Ctrl+Shift+T)",
    "搜索当前文件夹 (Ctrl+F)": "Search Current Folder (Ctrl+F)",
    "书签管理": "Bookmarks",
    "设置": "Settings",
    "AI 助手面板 (Ctrl+Shift+A)": "AI Assistant (Ctrl+Shift+A)",
    "打开 TortoiseGit 日志": "Open TortoiseGit Log",
    "打开 TortoiseGit 提交窗口": "Open TortoiseGit Commit",
    "在当前标签页路径打开 Git Bash": "Open Git Bash here",
    "在当前标签页路径打开 cmd": "Open cmd here",
    "在当前标签页路径打开 PowerShell": "Open PowerShell here",
    "打开计算器": "Open Calculator",
    # ── Title bar launcher ────────────────────────────────────────────────
    "拖入应用或快捷方式": "Drop app or shortcut",
    "+ 拖入应用或快捷方式": "+ Drop app or shortcut",
    "移除该快捷方式": "Remove shortcut",
    # ── Tab context menu ──────────────────────────────────────────────────
    "📌 取消固定": "📌 Unpin",
    "📌 固定": "📌 Pin",
    "🔨 取消固定": "🔨 Unpin",
    "🔖 添加书签": "🔖 Add Bookmark",
    "▶ 恢复自动刷新": "▶ Resume Auto-refresh",
    "⏸ 冻结自动刷新": "⏸ Freeze Auto-refresh",
    "自动刷新": "Auto-refresh",
    "当前标签自动刷新已冻结": "Auto-refresh frozen",
    "当前标签自动刷新已恢复": "Auto-refresh resumed",
    # ── File operations ───────────────────────────────────────────────────
    "新建文件夹": "New Folder",
    "编辑": "Edit",
    # ── AI panel ─────────────────────────────────────────────────────────
    "🤖 AI 助手": "🤖 AI Assistant",
    "清空": "Clear",
    "当前目录: {}": "Current Dir: {}",
    "当前目录: —": "Current Dir: —",
    "输入问题… (Enter 发送，Shift+Enter 换行)": "Ask a question… (Enter to send, Shift+Enter for newline)",
    "发 送": "Send",
    "复制": "Copy",
    "全选": "Select All",
    "清空聊天": "Clear Chat",
    "👤 你": "👤 You",
    "🖥 系统": "🖥 System",
    "⏳ AI 思考中…": "⏳ AI thinking…",
    "⏳ AI 推理中{}…": "⏳ AI reasoning{}…",
    "⏳ AI 第 {} 轮推理中…": "⏳ AI round {} reasoning…",
    "⚠️ 检测到重复工具调用，强制要求给出结论…": "⚠️ Repeated tool call detected, forcing conclusion…",
    "空路径": "Empty path",
    "❌ 请先在 设置 → AI 助手 中填写 API 地址": "❌ Please fill in API URL in Settings → AI Assistant",
    "❌ 请求失败: {}": "❌ Request failed: {}",
    "⏸ 已取消删除: {}": "⏸ Delete cancelled: {}",
    "⏸ 已取消覆盖文件: {}": "⏸ File overwrite cancelled: {}",
    "⏸ 已取消运行脚本: {}": "⏸ Script run cancelled: {}",
    "确认运行脚本": "Confirm Run Script",
    "确认覆盖文件": "Confirm Overwrite File",
    "目录（及其所有内容）": "directory (and all contents)",
    "确认删除": "Confirm Delete",
    "全量覆盖会丢失未读内容。": "Full overwrite will discard unread content.",
    "请改用 [PATCH_FILE: 路径|旧文本|新文本] 进行局部修改。": "Use [PATCH_FILE: path|old|new] for partial edits instead.",
    # ── AI tool result strings ────────────────────────────────────────────
    "[LIST_DIR 结果] 目录 {}:\n{}": "[LIST_DIR result] Directory {}:\n{}",
    "[当前目录: {}]\n": "[Current dir: {}]\n",
    "📁 目录列表: {}\n{}": "📁 Directory listing: {}\n{}",
    "📂 打开目录: {}": "📂 Open dir: {}",
    "📄 文件内容（{}-{}/{}字节，{}）: {}": "📄 File content ({}-{}/{}B, {}): {}",
    "✅ 已写入文件: {}": "✅ Written: {}",
    "✅ 已创建目录: {}": "✅ Created dir: {}",
    "✅ 已删除{}: {}": "✅ Deleted {}: {}",
    "✅ 已启动脚本: {}": "✅ Script started: {}",
    "✅ 已在当前标签切换到目录: {}": "✅ Switched to dir: {}",
    "✅ 已打开目录: {}": "✅ Opened dir: {}",
    "✅ 补丁成功: {}": "✅ Patch applied: {}",
    "❌ PATCH_FILE 格式: [PATCH_FILE: 路径|旧文本|新文本]": "❌ PATCH_FILE format: [PATCH_FILE: path|old|new]",
    "❌ 写入失败: {}": "❌ Write failed: {}",
    "❌ 写入失败: {}（{}）": "❌ Write failed: {} ({})",
    "❌ 写入失败: 格式应为 [WRITE_FILE: 路径|内容]": "❌ Write failed: format must be [WRITE_FILE: path|content]",
    "❌ 列目录失败: {}": "❌ List dir failed: {}",
    "❌ 列目录失败: {}（{}）": "❌ List dir failed: {} ({})",
    "❌ 创建目录失败: {}": "❌ Create dir failed: {}",
    "❌ 创建目录失败: {}（{}）": "❌ Create dir failed: {} ({})",
    "❌ 删除失败: {}": "❌ Delete failed: {}",
    "❌ 删除失败: {}（{}）": "❌ Delete failed: {} ({})",
    "❌ 打开目录失败: {}（{}）": "❌ Open dir failed: {} ({})",
    "❌ 文件不存在: {}": "❌ File not found: {}",
    "❌ 无法读取文件: {}": "❌ Cannot read file: {}",
    "❌ 目录不存在: {}": "❌ Dir not found: {}",
    "❌ 补丁失败: {}": "❌ Patch failed: {}",
    "❌ 补丁失败: {}（{}）": "❌ Patch failed: {} ({})",
    "❌ 补丁失败：在 {} 中未找到目标文本": "❌ Patch failed: target text not found in {}",
    "❌ 读取失败: {}": "❌ Read failed: {}",
    "❌ 读取失败: {}（{}）": "❌ Read failed: {} ({})",
    "❌ 路径不存在: {}": "❌ Path not found: {}",
    "❌ 运行脚本失败: {}（{}）": "❌ Script failed: {} ({})",
    "⚠️ 文件太大（{:.1f}MB），无法读取": "⚠️ File too large ({:.1f}MB) to read",
    "⚠️ 目录不存在: {}": "⚠️ Dir not found: {}",
    "⚠️ 脚本不存在: {}": "⚠️ Script not found: {}",
    "\n⚠️ 文件过大，仅读取前 {} 字符，剩余 {} 字节未读": "\n⚠️ File too large; read first {} chars, {} bytes remaining",
    "\n\n【已截断 {} 个字符】": "\n\n[Truncated {} chars]",
    "（这是文件，请用 READ_FILE）": "(This is a file; use READ_FILE)",
    "（这是目录，请用 LIST_DIR 列目录，或直接指定具体 .c/.h 文件路径）": "(This is a directory; use LIST_DIR or specify a .c/.h file path)",
    "运行: {}": "Run: {}",
    "超出当前目录范围: {}": "Outside current directory scope: {}",
    "路径非法: {}": "Invalid path: {}",
    "路径使用 Windows 格式。": "Use Windows path format.",
    "路径使用 Windows 格式，例如 D:\\project\\src。": "Use Windows path format, e.g. D:\\project\\src.",
    "路径使用 Windows 格式，例如 D:\\project\\src。\n": "Use Windows path format, e.g. D:\\project\\src.\n",
    "保存聊天记录失败: {}": "Failed to save chat history: {}",
    "加载聊天记录失败: {}": "Failed to load chat history: {}",
    "删除聊天记录文件失败: {}": "Failed to delete chat history file: {}",
    "保存失败: {}": "Save failed: {}",
    "第 {} 轮": "Round {}",
    "ℹ 当前标签已在该目录: {}": "ℹ Current tab is already at: {}",
    "ℹ 系统": "ℹ System",
    # ── AI system prompts ─────────────────────────────────────────────────
    "你是一个智能文件管理助手。": "You are an intelligent file management assistant.",
    "你是一个智能文件管理助手，帮助用户管理文件系统、打开目录和运行脚本。\n":
        "You are an intelligent file management assistant that helps users manage the file system, open directories, and run scripts.\n",
    "你是一个智能文件管理助手，正在通过工具调用完成用户任务。\n":
        "You are an intelligent file management assistant completing user tasks via tool calls.\n",
    "⚠️ 只有当用户明确要求'切换/打开目录'时，才在回复末尾添加：[OPEN_DIR: 目录完整路径]。":
        "⚠️ Only when the user explicitly asks to 'switch/open a directory', append: [OPEN_DIR: full/path].",
    "普通问答不要输出任何操作指令。\n": "Do not output any operation commands for general Q&A.\n",
    "⚠️ 只有当用户明确要求运行脚本时，才添加：[RUN_SCRIPT: 脚本完整路径]。\n":
        "⚠️ Only when the user explicitly asks to run a script, append: [RUN_SCRIPT: full/script/path].\n",
    "⚠️ 仅当用户明确要求时，才可使用以下指令：\n": "⚠️ Use the following commands ONLY when explicitly requested:\n",
    "[READ_FILE: 文件路径] 读取文件（系统自动分段读取大文件，无需手动指定偏移）；\n":
        "[READ_FILE: path] Read file (system auto-paginates large files; no offset needed);\n",
    "[PATCH_FILE: 路径|旧文本|新文本] 局部修改已有文件（安全，仅替换目标内容段）；\n":
        "[PATCH_FILE: path|old_text|new_text] Partial edit of existing file (safe, replaces only target section);\n",
    "⚠️ 修改已有文件时必须使用 PATCH_FILE，禁止用 WRITE_FILE 覆盖已有大文件；\n":
        "⚠️ Use PATCH_FILE for existing files; do NOT use WRITE_FILE to overwrite large existing files;\n",
    "⚠️ PATCH_FILE 使用规范：\n": "⚠️ PATCH_FILE usage rules:\n",
    "⚠️ PATCH_FILE 最小改动规范：\n": "⚠️ PATCH_FILE minimal-diff rules:\n",
    "  1. 旧文本只需包含被修改的行及前后各1-2行（足够唯一定位即可），不要复制整个函数体；\n":
        "  1. old_text only needs the changed line(s) plus 1-2 lines of context (enough to uniquely locate); don't copy whole functions;\n",
    "  1. 旧文本只取被修改行及前后各1-2行（能唯一定位即可），不要复制整个函数；\n":
        "  1. old_text: only the changed line(s) + 1-2 context lines (uniquely locatable); don't copy whole functions;\n",
    "  2. 对同一文件做多个 PATCH_FILE 时，第二个 patch 的旧文本必须是第一个 patch 应用后的实际内容；\n":
        "  2. For multiple PATCHes on the same file, each old_text must reflect content after previous patches;\n",
    "  2. 对同一文件连续多个 PATCH_FILE 时，后一个的旧文本必须反映前一个 patch 已应用后的文件内容；\n":
        "  2. For sequential PATCHes, each old_text must reflect the file state after prior patches;\n",
    "  3. 最小改动原则：新文本只改动必要的行，不要重写周围未变更的代码。\n":
        "  3. Minimal change principle: new_text only changes necessary lines; don't rewrite unchanged surrounding code.\n",
    "  3. 新文本只改动必要的行，保留其余未变行，使 diff 最小化。\n":
        "  3. new_text changes only necessary lines; keep unchanged lines to minimize diff.\n",
    "[WRITE_FILE: 文件路径|文件内容] 仅用于创建全新文件（已有文件禁止使用）；\n":
        "[WRITE_FILE: path|content] Only for creating brand-new files (forbidden for existing files);\n",
    "[LIST_DIR: 目录路径] 列出目录；\n": "[LIST_DIR: path] List directory;\n",
    "[MKDIR: 目录路径] 创建目录；\n": "[MKDIR: path] Create directory;\n",
    "[DELETE: 路径] 删除文件或目录（需用户确认）。\n": "[DELETE: path] Delete file or directory (requires user confirmation).\n",
    "可以在同一回复中包含多个操作命令。": "Multiple operation commands may appear in the same reply.",
    "你可以自由使用以下工具指令：\n": "You may freely use these tool commands:\n",
    "[READ_FILE: 路径] 读文件（优先读 .c/.h 源代码）；": "[READ_FILE: path] Read file (prefer .c/.h source);\n",
    "[PATCH_FILE: 路径|旧文本|新文本] 局部修改文件（修改已有文件时必须用此，禁止用 WRITE_FILE 覆盖已有大文件）；":
        "[PATCH_FILE: path|old|new] Partial file edit (required for existing files; never use WRITE_FILE to overwrite);",
    "[LIST_DIR: 路径] 列目录（仅在不知道源码位置时使用）；": "[LIST_DIR: path] List dir (use only when source location is unknown);",
    "[OPEN_DIR: 路径] [MKDIR: 路径] [WRITE_FILE: 路径|内容]（仅用于创建新文件） [DELETE: 路径]。\n":
        "[OPEN_DIR: path] [MKDIR: path] [WRITE_FILE: path|content] (new files only) [DELETE: path].\n",
    "ℹ️ 工作原则：得到目录列表后尽快选择源代码文件直接阅读，":
        "ℹ️ Workflow: after listing a directory, immediately select and read source files,",
    "而不要反复展开子目录；修改文件时优先使用 PATCH_FILE 安全替换具体代码段。\n":
        " rather than repeatedly expanding subdirectories; prefer PATCH_FILE for safe targeted edits.\n",
    "以上是工具调用结果。请高效利用工具，": "The above are tool call results. Use tools efficiently:",
    "优先阅读源代码文件（.c/.h）而非反复列目录，": " prefer reading source files (.c/.h) over repeatedly listing dirs,",
    "并尽快基于收集到的信息给出具体结论或优化建议。": " and provide concrete conclusions or optimization suggestions promptly.",
    "以上是最后一批工具调用结果。不要再调用任何工具指令，": "The above are the last tool call results. Do not call any more tools.",
    "必须直接基于已收集的所有信息给出具体的结论、优化建议或解决方案。":
        " Give specific conclusions, optimization suggestions, or solutions based on all collected information.",
    "不能将任何工具指令放入回复，": "Do not include any tool commands in the reply;",
    "必须直接基于已收集的全部信息给出具体的结论、优化建议或解决方案。":
        " give specific conclusions, suggestions, or solutions based on all collected information.",
    "请直接基于已收集的全部信息给出具体的结论、优化建议或解决方案。":
        "Provide specific conclusions, suggestions, or solutions based on all collected information.",
    "以上是工具调用结果。要高效利用工具，": "The above are tool call results. Use tools efficiently:",
    "你是一个智能文件管理助手，帮助用户管理文件系统。\n": "You are a smart file management assistant helping users manage the filesystem.\n",
    "不要将任何工具指令放入回复，": "Do not include any tool commands in the reply,",
    # ── Misc app messages ─────────────────────────────────────────────────
    "[App] 已加载固定标签页": "[App] Pinned tabs loaded",
    "[App] 已恢复上次激活的标签页": "[App] Last active tab restored",
    "[App] 没有缓存和固定标签，添加默认主目录标签": "[App] No cached/pinned tabs; adding default home tab",
    "[App] 程序关闭，已清除搜索缓存": "[App] App closed, search cache cleared",
    "[App] 跳过缓存标签（已固定）: {}": "[App] Skipped cached tab (pinned): {}",
    "[App] 跳过缓存标签（已打开）: {}": "[App] Skipped cached tab (already open): {}",
    " - 窗口恢复中": " - Restoring...",
    "{} - 窗口恢复中": "{} - Restoring...",
    # ── Windows special folders (used for shell path detection) ───────────
    "控制面板": "Control Panel",
    '控制面板': "Control Panel",
    '控制面板 - ': "Control Panel - ",
    '控制面板\\': "Control Panel\\",
    '\\控制面板\\': "\\Control Panel\\",
    "回收站": "Recycle Bin",
    "此电脑": "This PC",
    "我的电脑": "My Computer",
    "网络": "Network",
    "启动文件夹": "Startup Folder",
    "启动项": "Startup Items",
    "开机启动项": "Startup Items",
    "用户帐户": "User Accounts",
    "程序和功能": "Programs and Features",
    "系统": "System",
    "设备管理器": "Device Manager",
    "网络和共享中心": "Network and Sharing Center",
    "记事本": "Notepad",
    # ── Misc ──────────────────────────────────────────────────────────────
    "(空目录)": "(Empty directory)",
    "💡 提示：可以拖动书签和文件夹调整顺序和层级，调整后点击【保存】按钮保存更改":
        "💡 Tip: Drag bookmarks/folders to reorder; click 【Save】 to apply",
    "保存聊天记录失败: {}": "Failed to save chat: {}",
    "加载聊天记录失败: {}": "Failed to load chat: {}",
    "❌ 流式响应被服务器提前关闭（上下文可能过长）。\n":
        "❌ Streaming response closed early by server (context may be too long).\n",
    "❌ 流式响应被服务器提前关闭（上下文可能过长，请清空聊天记录重试）。\n":
        "❌ Streaming response closed early (context too long; please clear chat and retry).\n",
    "未找到 {}": "{} not found",
    "暂不支持打开此类型书签: {}": "Cannot open bookmark type: {}",
    "已在当前目录选中{}: {}": "Selected {} in current dir: {}",
}



import json as _json_lang, sys as _sys_lang
try:
    _cfg_lang_path = os.path.join(
        os.path.dirname(os.path.abspath(
            _sys_lang.executable if getattr(_sys_lang, 'frozen', False) else __file__
        )), 'config.json')
    with open(_cfg_lang_path, 'r', encoding='utf-8') as _lf:
        _app_language = _json_lang.load(_lf).get('language', 'zh')
except Exception:
    _app_language = 'zh'
del _json_lang, _sys_lang



def get_app_base_dir():
    """Return a stable base directory for persistent app data."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_app_data_path(*parts):
    return os.path.join(get_app_base_dir(), *parts)

# 常用中英文目录名对照表
_COMMON_PATH_PAIRS = [
    ("Users", "用户"),
    ("Documents", "文档"),
    ("Desktop", "桌面"),
    ("Downloads", "下载"),
    ("Pictures", "图片"),
    ("Music", "音乐"),
    ("Videos", "视频"),
    ("Favorites", "收藏夹"),
    ("OneDrive", "OneDrive"),
    ("AppData", "AppData"),
    ("Roaming", "Roaming"),
    ("Local", "Local"),
    ("Public", "Public"),
]

def translate_common_path(path):
    """
    尝试将路径中的中英文常用目录名互相转换，递归尝试所有组合，返回第一个存在的路径。
    若找不到存在的路径，则返回原始 path。
    """
    if not path or os.path.exists(path):
        return path
    # 只处理本地绝对路径
    norm_path = os.path.normpath(path)
    # 分割为各级目录
    parts = norm_path.split(os.sep)
    # 记录所有可替换的目录位置
    replace_indices = []
    replace_options = []
    for i, part in enumerate(parts):
        opts = [part]
        for en, zh in _COMMON_PATH_PAIRS:
            if part == en:
                opts.append(zh)
            elif part == zh:
                opts.append(en)
        if len(opts) > 1:
            replace_indices.append(i)
            replace_options.append(opts)
    # 如果没有可替换的，直接返回原始
    if not replace_indices:
        return path
    # 递归生成所有组合
    from itertools import product
    for combo in product(*replace_options):
        new_parts = parts[:]
        for idx, val in zip(replace_indices, combo):
            new_parts[idx] = val
        candidate = os.sep.join(new_parts)
        if os.path.exists(candidate):
            return candidate
    # 没找到存在的路径，返回原始
    return path

# 性能优化配置常量
MAX_SEARCH_CACHE_SIZE = 50  # 搜索缓存最大数量
MAX_SEARCH_RESULTS = 1000000  # 单次搜索最大结果数
MAX_CACHED_RESULTS_PER_QUERY = 5000  # 单次搜索最多缓存结果数（控制内存占用）
CONTENT_SEARCH_CHUNK_SIZE = 10 * 1024 * 1024  # 内容搜索分块大小（10MB）
CONTENT_SEARCH_MAX_BYTES_PER_FILE = 64 * 1024 * 1024  # 单文件最多扫描64MB，避免超大文件拖慢整体
CONTENT_SEARCH_IN_MEMORY_THRESHOLD = 2 * 1024 * 1024  # 小文件（<=2MB）一次性读入内存后多编码匹配
CONTENT_SEARCH_ENCODINGS = ['utf-8', 'utf-8-sig', 'utf-16', 'utf-16-le', 'utf-16-be', 'gbk', 'gb2312', 'latin-1']
SEARCH_RESULT_QUEUE_MAXSIZE = 3000  # 搜索结果队列容量（降低高吞吐时溢出概率）
SEARCH_RESULT_BATCH_BASE = 100  # 搜索线程默认批量发送大小
SEARCH_RESULT_BATCH_MIN = 50  # 低积压时最小批量发送大小
SEARCH_RESULT_BATCH_MAX = 400  # 高积压时最大批量发送大小
SEARCH_METADATA_DEGRADE_ENABLED = True  # 队列高压时降级元数据(stat)获取
SEARCH_METADATA_DEGRADE_QUEUE_RATIO = 0.75  # 触发降级的队列占用比例
MAX_CLOSED_TABS_HISTORY = 20  # 关闭标签页历史最大数量（从10增加到20）
MAX_SEARCH_HISTORY = 30  # 搜索历史最大数量（从20增加到30）
MAX_NAVIGATION_HISTORY = 50  # 导航历史最大数量
MAX_ACTIVE_TOASTS = 5  # 同时显示的提示数量上限
MAX_CHAT_HISTORY_MESSAGES = 80   # AI 聊天历史最大消息数
MAX_CHAT_MESSAGE_CHARS = 12000   # 单条 AI 消息最大长度，防止历史长期膨胀
MAX_CONTEXT_TOTAL_CHARS = 40000  # 发给 API 的全部历史消息总字符上限（约 10K token），超出则丢弃最老消息
HOUSEKEEPING_INTERVAL_MS = 5 * 60 * 1000  # 低频运行时清理周期（5分钟）
HOUSEKEEPING_GC_EVERY_N = 3  # 每 N 次清理执行一次 gc.collect()
SESSION_SNAPSHOT_INTERVAL_MS = 15000  # 崩溃恢复兜底：定期写入当前会话快照
SESSION_SNAPSHOT_DEBOUNCE_MS = 1200  # 标签/路径变化后的会话快照防抖时间
SESSION_SNAPSHOT_MIN_INTERVAL_MS = 8000  # 事件驱动快照最小间隔，防止 DirPoll/FileWatcher 高频触发写盘
MAX_HEALTH_LOG_BYTES = 1024 * 1024  # runtime_health.log 超过此大小即轮转为 .1，避免长期运行无限增长
APP_INTERNAL_CHANGE_FILENAMES = {
    'config.json',
    'config.json.tmp',
    'bookmarks.json',
    'chat_history.json',
    'runtime_health.log',
    'runtime_health.log.1',
}

# 大文件夹异步加载配置
LARGE_FOLDER_THRESHOLD = 1000  # 超过此数量文件视为大文件夹
FOLDER_CHECK_TIMEOUT = 500  # 文件夹检查超时时间(ms)
ASYNC_LOAD_ENABLED = True  # 是否启用异步加载
# 慢盘（网络/UNC/映射盘）导航兜底超时：后台解析成功但 NavigateComplete2 因网络中断
# 始终不触发时，用此超时解除“导航中”锁定并隐藏 loading，避免标签永久卡在加载态。
ASYNC_NAV_TIMEOUT_MS = 20000
# 兜底轮询快照逐项 stat 的上限：超大目录每 8s 在 UI 线程逐项 stat 会造成周期性卡顿，
# 超过此上限后停止逐项统计，改用“总项目数 + 目录自身 mtime”兜底检测增删/重命名。
DIR_SNAPSHOT_MAX_ENTRIES = 5000


def _compute_dir_snapshot(path, ignore_check=None):
    """计算目录轻量元数据快照 (count, latest_mtime_ns, size_sum, name_hash)，失败返回 None。

    纯 I/O + 计算，不触碰任何 Qt 对象，可安全在后台线程（QRunnable）中执行。
    ignore_check(path, name) 用于忽略应用自身写出的配置/日志文件（可为 None）。"""
    try:
        count = 0
        latest_mtime_ns = 0
        size_sum = 0
        name_hash = 0
        truncated = False
        with os.scandir(path) as entries:
            for entry in entries:
                if ignore_check is not None and ignore_check(path, entry.name):
                    continue
                count += 1
                if count > DIR_SNAPSHOT_MAX_ENTRIES:
                    # 超大目录：停止逐项 stat，仍继续累计总数，配合目录自身 mtime 兜底
                    truncated = True
                    continue
                try:
                    is_dir = entry.is_dir(follow_symlinks=False)
                    stat = entry.stat(follow_symlinks=False)
                    mtime_ns = getattr(stat, 'st_mtime_ns', None)
                    if mtime_ns is None:
                        mtime_ns = int(stat.st_mtime * 1_000_000_000)
                    if mtime_ns > latest_mtime_ns:
                        latest_mtime_ns = mtime_ns
                    if not is_dir:
                        size_sum += getattr(stat, 'st_size', 0)
                    name_hash ^= hash((entry.name, is_dir))
                except Exception:
                    continue
        if truncated:
            # 折叠目录自身 mtime，保证超出上限部分的增删/重命名仍能触发刷新
            try:
                dir_mtime_ns = os.stat(path).st_mtime_ns
                if dir_mtime_ns > latest_mtime_ns:
                    latest_mtime_ns = dir_mtime_ns
            except Exception:
                pass
        return (count, latest_mtime_ns, size_sum, name_hash)
    except Exception:
        return None

STATUS_SELECTION_METADATA_LIMIT = 20  # 多选超过阈值时跳过逐项大小统计
SHORTCUT_POLL_ACTIVE_MS = 60  # 主窗口激活时快捷键轮询频率（需足够快以捕捉"同时按住"的短促组合键）
SHORTCUT_POLL_INACTIVE_MS = 500  # 主窗口非激活时快捷键轮询频率
# ── 主线程 COM 轮询抗高负载保护 ───────────────────────────────────────────────
# LocationURL 是同步跨进程 COM 调用，无超时。CPU 饱和时其延迟会飙升，阻塞 UI 线程，
# 且卡顿时 Qt 定时器事件堆积、线程一空就爆发式触发，形成无法恢复的“死亡螺旋”。
COM_POLL_SLOW_MS = 180        # 单次 LocationURL 调用超过此耗时即判定系统高负载
COM_POLL_STRESS_BACKOFF_MS = 3000  # 高负载期间将 COM 轮询间隔退避到此值，给 UI 线程喘息
COM_POLL_MIN_GAP_MS = 50      # 挂钟防抖：两次实际轮询的最小真实间隔，吸收卡顿后的爆发触发
COM_POLL_HARD_DEADLINE_MS = 250  # QAx LocationURL 看门狗硬超时：超时即放弃读取并用缓存值
# 状态栏资源占用颜色预警阈值（百分比）：低于 WARN 绿色，WARN~CRIT 橙色，>=CRIT 红色
RESOURCE_WARN_PERCENT = 75
RESOURCE_CRIT_PERCENT = 90
SEARCH_RESULT_TYPE_COL_WIDTH = 90
SEARCH_RESULT_DATE_COL_WIDTH = 155
SEARCH_RESULT_SIZE_COL_WIDTH = 100
STATUS_UPDATE_DEFER_MS = 80
STATUS_TRACKING_INTERVAL_MS = 220
STATUS_TRACKING_WINDOW_MS = 1400
TITLE_SHORTCUT_EXTENSIONS = ('.lnk', '.exe', '.bat', '.cmd', '.ps1')
SUPPORTED_TERMINAL_TOOLS = ('cmd', 'powershell', 'git-bash')


def apply_runtime_performance_config(perf_cfg=None):
    """将配置中的性能参数应用到运行时常量（带边界校验）。"""
    global CONTENT_SEARCH_CHUNK_SIZE
    global CONTENT_SEARCH_MAX_BYTES_PER_FILE
    global CONTENT_SEARCH_IN_MEMORY_THRESHOLD
    global SEARCH_RESULT_QUEUE_MAXSIZE
    global SEARCH_RESULT_BATCH_BASE
    global SEARCH_RESULT_BATCH_MIN
    global SEARCH_RESULT_BATCH_MAX
    global SEARCH_METADATA_DEGRADE_ENABLED
    global SEARCH_METADATA_DEGRADE_QUEUE_RATIO

    if not isinstance(perf_cfg, dict):
        return

    def _clamp_int(val, default_val, min_val, max_val):
        try:
            iv = int(val)
        except Exception:
            return default_val
        if iv < min_val:
            return min_val
        if iv > max_val:
            return max_val
        return iv

    def _clamp_float(val, default_val, min_val, max_val):
        try:
            fv = float(val)
        except Exception:
            return default_val
        if fv < min_val:
            return min_val
        if fv > max_val:
            return max_val
        return fv

    def _to_bool(val, default_val):
        if isinstance(val, bool):
            return val
        if isinstance(val, (int, float)):
            return bool(val)
        if isinstance(val, str):
            v = val.strip().lower()
            if v in ('1', 'true', 'yes', 'on'):
                return True
            if v in ('0', 'false', 'no', 'off'):
                return False
        return default_val

    CONTENT_SEARCH_CHUNK_SIZE = _clamp_int(
        perf_cfg.get("content_search_chunk_size", CONTENT_SEARCH_CHUNK_SIZE),
        CONTENT_SEARCH_CHUNK_SIZE,
        256 * 1024,
        64 * 1024 * 1024,
    )
    CONTENT_SEARCH_MAX_BYTES_PER_FILE = _clamp_int(
        perf_cfg.get("content_search_max_bytes_per_file", CONTENT_SEARCH_MAX_BYTES_PER_FILE),
        CONTENT_SEARCH_MAX_BYTES_PER_FILE,
        1 * 1024 * 1024,
        1024 * 1024 * 1024,
    )
    CONTENT_SEARCH_IN_MEMORY_THRESHOLD = _clamp_int(
        perf_cfg.get("content_search_in_memory_threshold", CONTENT_SEARCH_IN_MEMORY_THRESHOLD),
        CONTENT_SEARCH_IN_MEMORY_THRESHOLD,
        64 * 1024,
        16 * 1024 * 1024,
    )
    SEARCH_RESULT_QUEUE_MAXSIZE = _clamp_int(
        perf_cfg.get("search_result_queue_maxsize", SEARCH_RESULT_QUEUE_MAXSIZE),
        SEARCH_RESULT_QUEUE_MAXSIZE,
        200,
        100000,
    )
    SEARCH_RESULT_BATCH_BASE = _clamp_int(
        perf_cfg.get("search_result_batch_base", SEARCH_RESULT_BATCH_BASE),
        SEARCH_RESULT_BATCH_BASE,
        20,
        2000,
    )
    SEARCH_RESULT_BATCH_MIN = _clamp_int(
        perf_cfg.get("search_result_batch_min", SEARCH_RESULT_BATCH_MIN),
        SEARCH_RESULT_BATCH_MIN,
        10,
        1000,
    )
    SEARCH_RESULT_BATCH_MAX = _clamp_int(
        perf_cfg.get("search_result_batch_max", SEARCH_RESULT_BATCH_MAX),
        SEARCH_RESULT_BATCH_MAX,
        20,
        5000,
    )
    SEARCH_METADATA_DEGRADE_ENABLED = _to_bool(
        perf_cfg.get("search_metadata_degrade_enabled", SEARCH_METADATA_DEGRADE_ENABLED),
        SEARCH_METADATA_DEGRADE_ENABLED,
    )
    SEARCH_METADATA_DEGRADE_QUEUE_RATIO = _clamp_float(
        perf_cfg.get("search_metadata_degrade_queue_ratio", SEARCH_METADATA_DEGRADE_QUEUE_RATIO),
        SEARCH_METADATA_DEGRADE_QUEUE_RATIO,
        0.1,
        0.98,
    )

    # 保证内存阈值不大于单文件扫描上限
    if CONTENT_SEARCH_IN_MEMORY_THRESHOLD > CONTENT_SEARCH_MAX_BYTES_PER_FILE:
        CONTENT_SEARCH_IN_MEMORY_THRESHOLD = CONTENT_SEARCH_MAX_BYTES_PER_FILE
    if SEARCH_RESULT_BATCH_MIN > SEARCH_RESULT_BATCH_MAX:
        SEARCH_RESULT_BATCH_MIN = SEARCH_RESULT_BATCH_MAX
    if SEARCH_RESULT_BATCH_BASE < SEARCH_RESULT_BATCH_MIN:
        SEARCH_RESULT_BATCH_BASE = SEARCH_RESULT_BATCH_MIN
    elif SEARCH_RESULT_BATCH_BASE > SEARCH_RESULT_BATCH_MAX:
        SEARCH_RESULT_BATCH_BASE = SEARCH_RESULT_BATCH_MAX

    debug_print(
        "[Config] Performance applied:",
        f"chunk={CONTENT_SEARCH_CHUNK_SIZE}",
        f"max_file={CONTENT_SEARCH_MAX_BYTES_PER_FILE}",
        f"in_memory={CONTENT_SEARCH_IN_MEMORY_THRESHOLD}",
        f"queue={SEARCH_RESULT_QUEUE_MAXSIZE}",
        f"batch_base={SEARCH_RESULT_BATCH_BASE}",
        f"batch_min={SEARCH_RESULT_BATCH_MIN}",
        f"batch_max={SEARCH_RESULT_BATCH_MAX}",
        f"meta_degrade={SEARCH_METADATA_DEGRADE_ENABLED}",
        f"meta_ratio={SEARCH_METADATA_DEGRADE_QUEUE_RATIO}",
    )

class SearchCache:
    """搜索结果缓存，使用LRU策略"""
    def __init__(self, max_size=50):
        self.cache = OrderedDict()
        self.max_size = max_size
    
    def get_key(self, search_path, keyword, search_filename, search_content, file_types, force_metadata_degrade=False, match_case=False, match_whole_word=False):
        """生成缓存键"""
        key_str = f"{search_path}|{keyword}|{search_filename}|{search_content}|{file_types}|{force_metadata_degrade}|{match_case}|{match_whole_word}"
        return hashlib.md5(key_str.encode()).hexdigest()
    
    def get(self, key):
        """获取缓存结果"""
        if key in self.cache:
            # 移到末尾（最近使用）
            self.cache.move_to_end(key)
            return self.cache[key]
        return None
    
    def put(self, key, value):
        """存储缓存结果"""
        if key in self.cache:
            self.cache.move_to_end(key)
        else:
            self.cache[key] = value
            # 如果超过最大缓存数，删除最旧的
            if len(self.cache) > self.max_size:
                self.cache.popitem(last=False)
    
    def clear(self):
        """清空缓存"""
        self.cache.clear()

# 全局搜索缓存实例（使用配置常量）
_search_cache = SearchCache(max_size=MAX_SEARCH_CACHE_SIZE)

from PyQt5.QtWidgets import QDialog, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QPushButton
# 多层结构书签弹窗
class BookmarkDialog(QDialog):
    def __init__(self, bookmark_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("书签"))
        self.resize(500, 600)
        self.bookmark_manager = bookmark_manager
        layout = QVBoxLayout(self)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels([tr("名称"), tr("路径")])
        layout.addWidget(self.tree)
        self.populate_tree()
        self.tree.itemDoubleClicked.connect(self.on_item_double_clicked)
        close_btn = QPushButton(tr("关闭"))
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

    def populate_tree(self):
        self.tree.clear()
        tree = self.bookmark_manager.get_tree()
        for root_name, root in tree.items():
            self.add_node(None, root)

    def add_node(self, parent_item, node):
        if node.get('type') == 'folder':
            item = QTreeWidgetItem([node.get('name', ''), ''])
            if parent_item:
                parent_item.addChild(item)
            else:
                self.tree.addTopLevelItem(item)
            for child in node.get('children', []):
                self.add_node(item, child)
        elif node.get('type') == 'url':
            item = QTreeWidgetItem([node.get('name', ''), node.get('url', '')])
            if parent_item:
                parent_item.addChild(item)
            else:
                self.tree.addTopLevelItem(item)

    def on_item_double_clicked(self, item, column):
        url = item.text(1)
        if url and url.startswith('file:///'):
            from urllib.parse import unquote
            local_path = unquote(url[8:])
            if os.name == 'nt' and local_path.startswith('/'):
                local_path = local_path[1:]
            local_path2 = translate_common_path(local_path)
            if os.path.exists(local_path2):
                self.accept()
                if self.parent() and hasattr(self.parent(), 'add_new_tab'):
                    self.parent().add_new_tab(local_path2)
            else:
                show_toast(self, tr("路径错误"), tr("路径不存在: {}").format(local_path2), level="warning")


class QuickFindResultsDialog(QDialog):
    def __init__(self, matched_paths, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("选择匹配项"))
        self.resize(680, 420)
        self.selected_path = None

        layout = QVBoxLayout(self)

        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels([tr("名称"), tr("类型"), tr("完整路径")])
        self.table.setRowCount(len(matched_paths))
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.itemDoubleClicked.connect(self._accept_current_selection)

        for row, path in enumerate(matched_paths):
            name_item = QTableWidgetItem(os.path.basename(path))
            type_item = QTableWidgetItem(tr("文件夹") if os.path.isdir(path) else tr("文件"))
            path_item = QTableWidgetItem(path)
            name_item.setData(Qt.UserRole, path)
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, type_item)
            self.table.setItem(row, 2, path_item)

        if matched_paths:
            self.table.selectRow(0)

        layout.addWidget(self.table)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        ok_btn = QPushButton(tr("确定"), self)
        cancel_btn = QPushButton(tr("取消"), self)
        ok_btn.clicked.connect(self._accept_current_selection)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

    def _accept_current_selection(self):
        current_row = self.table.currentRow()
        if current_row < 0:
            return
        current_item = self.table.item(current_row, 0)
        if current_item is None:
            return
        self.selected_path = current_item.data(Qt.UserRole)
        if self.selected_path:
            self.accept()

# 自定义委托：在文件名列实现省略号在开头
from PyQt5.QtWidgets import QStyledItemDelegate, QTableView, QAbstractItemView
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QPainter

class ElideLeftDelegate(QStyledItemDelegate):
    """自定义委托，文本过长时在开头显示省略号"""
    def paint(self, painter, option, index):
        if index.column() == 0:  # 只对第一列（文件名列）应用
            painter.save()
            # 获取完整文本
            text = index.data(Qt.DisplayRole)
            # 使用字体度量计算省略文本
            fm = painter.fontMetrics()
            elided_text = fm.elidedText(text, Qt.ElideLeft, option.rect.width() - 10)
            # 绘制文本
            painter.drawText(option.rect.adjusted(5, 0, -5, 0), Qt.AlignLeft | Qt.AlignVCenter, elided_text)
            painter.restore()
        else:
            super().paint(painter, option, index)


class SearchResultsTableModel(QAbstractTableModel):
    _HEADER_KEYS = ["文件名", "类型", "修改日期", "大小"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows = []

    @property
    def HEADERS(self):
        return [tr(k) for k in self._HEADER_KEYS]

    def rowCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self._rows)

    def columnCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self.HEADERS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row = self._rows[index.row()]
        column = index.column()

        if role == Qt.DisplayRole:
            if column == 0:
                return row.get('name', '')
            if column == 1:
                return row.get('file_type', '')
            if column == 2:
                return row.get('date', '')
            if column == 3:
                return row.get('size', '')
        elif role == Qt.ToolTipRole:
            if column == 0:
                return row.get('full_path') or row.get('path', '')
            return row.get('path', '')
        elif role == Qt.UserRole:
            return row.get('path', '')
        elif role == Qt.TextAlignmentRole and column == 0:
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal and 0 <= section < len(self.HEADERS):
            return self.HEADERS[section]
        return super().headerData(section, orientation, role)

    def sort(self, column, order=Qt.AscendingOrder):
        key_map = {
            0: lambda row: str(row.get('name', '')).lower(),
            1: lambda row: str(row.get('file_type', '')).lower(),
            2: lambda row: (row.get('sort_date_ts') is None, row.get('sort_date_ts') if row.get('sort_date_ts') is not None else 0, str(row.get('date', ''))),
            3: lambda row: (row.get('sort_size_bytes') is None, row.get('sort_size_bytes') if row.get('sort_size_bytes') is not None else -1, str(row.get('size', ''))),
        }
        key_fn = key_map.get(column)
        if key_fn is None or len(self._rows) <= 1:
            return
        self.layoutAboutToBeChanged.emit()
        self._rows.sort(key=key_fn, reverse=(order == Qt.DescendingOrder))
        self.layoutChanged.emit()

    def clear(self):
        self.beginResetModel()
        self._rows = []
        self.endResetModel()

    def append_results(self, rows):
        if not rows:
            return 0
        start_row = len(self._rows)
        end_row = start_row + len(rows) - 1
        self.beginInsertRows(QModelIndex(), start_row, end_row)
        normalized_rows = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            row_copy = dict(row)
            row_copy.setdefault('sort_date_ts', None)
            row_copy.setdefault('sort_size_bytes', None)
            normalized_rows.append(row_copy)
        self._rows.extend(normalized_rows)
        self.endInsertRows()
        return len(normalized_rows)

    def path_for_row(self, row):
        if 0 <= row < len(self._rows):
            return self._rows[row].get('path', '')
        return ''

# Everything 搜索引擎集成
def detect_everything():
    """检测系统中是否安装了Everything"""
    import shutil
    # 检查Everything命令行工具es.exe是否在PATH中
    es_path = shutil.which('es.exe')
    if es_path:
        return es_path
    
    # 检查常见安装路径
    common_paths = [
        r'C:\Program Files\Everything\es.exe',
        r'C:\Program Files (x86)\Everything\es.exe',
        os.path.expandvars(r'%PROGRAMFILES%\Everything\es.exe'),
        os.path.expandvars(r'%PROGRAMFILES(X86)%\Everything\es.exe'),
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
    
    return None


def detect_notepad_plus_plus():
    """检测系统中是否安装了 Notepad++。"""
    import shutil

    exe_path = shutil.which('notepad++.exe')
    if exe_path:
        return exe_path

    common_paths = [
        r'C:\Program Files\Notepad++\notepad++.exe',
        r'C:\Program Files (x86)\Notepad++\notepad++.exe',
        os.path.expandvars(r'%PROGRAMFILES%\Notepad++\notepad++.exe'),
        os.path.expandvars(r'%PROGRAMFILES(X86)%\Notepad++\notepad++.exe'),
        os.path.expandvars(r'%LOCALAPPDATA%\Programs\Notepad++\notepad++.exe'),
    ]
    for path in common_paths:
        if os.path.exists(path):
            return path
    return None


def format_file_size(size_bytes):
    """将字节数格式化为带单位的人类可读字符串。"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    if size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


def is_text_file(file_path, sample_size=1024):
    """智能检测文件是否为文本文件（读取前N字节检测）。

    模块级实现：避免在每次搜索时重新定义闭包，且可独立复用/测试。
    """
    try:
        with open(file_path, 'rb') as f:
            sample = f.read(sample_size)
            if not sample:
                return True  # 空文件视为文本文件

            # UTF-16/UTF-32 BOM 视为文本，避免被 NULL 字节规则误判
            if sample.startswith((b'\xff\xfe', b'\xfe\xff', b'\xff\xfe\x00\x00', b'\x00\x00\xfe\xff')):
                return True

            # NULL 字节几乎总是二进制特征
            if b'\x00' in sample:
                return False

            # 仅统计 ASCII 控制字符（排除 \t\n\r），避免把 UTF-8 非 ASCII 文本误判为二进制
            control_count = 0
            for byte in sample:
                if byte < 9 or (13 < byte < 32):
                    control_count += 1

            # 控制字符占比过高，判定为二进制
            if control_count / len(sample) > 0.1:
                return False

            return True
    except Exception:
        return False  # 无法读取则视为二进制

# 搜索对话框
class SearchDialog(QDialog):    
    def __init__(self, search_path, parent=None, search_history=None):
        super().__init__(parent)
        self.setWindowTitle(tr("搜索 - {}").format(search_path))
        # 设置为可调整大小，并显示最小化/最大化按钮
        self.setWindowFlags(
            Qt.Dialog
            | Qt.WindowTitleHint
            | Qt.WindowCloseButtonHint
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowSystemMenuHint
        )
        # 关闭时立即销毁 C++ 对象（释放所有 Qt 子控件占用的内存）
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.resize(800, 500)  # 初始大小，但允许调整
        self.search_path = search_path
        self.main_window = parent
        self.search_thread = None
        self.is_searching = False
        self.search_history = search_history or []  # 搜索历史列表
        
        # 检测Everything
        self.everything_path = detect_everything()
        self.notepad_plus_plus_path = detect_notepad_plus_plus()
        debug_print(f"[Search] Everything detected: {self.everything_path}")
        
        # 线程安全的结果队列（限制大小防止内存溢出）
        import queue
        self.result_queue = queue.Queue(maxsize=SEARCH_RESULT_QUEUE_MAXSIZE)  # 结果队列容量
        self.ui_update_timer = None
        self.queue_overflow_count = 0  # 队列溢出计数
        self._queue_idle_ticks = 0
        
        # 结果限制配置（使用虚拟滚动优化，支持更多结果）
        self.max_results = 1000000  # 最多显示100万个结果（虚拟滚动优化）
        self.current_result_count = 0
        self.batch_insert_size = 500  # 批量插入大小
        
        layout = QVBoxLayout(self)
        
        # 搜索选项区域
        search_options = QHBoxLayout()
        search_options.setSpacing(5)  # 设置控件间距为5像素
        
        # 搜索关键词（改为QComboBox支持历史记录）
        search_label = QLabel(tr("搜索:"))
        search_label.setFixedWidth(40)  # 固定标签宽度
        search_options.addWidget(search_label)
        from PyQt5.QtWidgets import QComboBox
        self.search_input = QComboBox()
        self.search_input.setEditable(True)
        self.search_input.setInsertPolicy(QComboBox.NoInsert)  # 不自动插入新条目
        self.search_input.setMinimumWidth(300)  # 设置最小宽度300像素
        self.search_input.lineEdit().setPlaceholderText(tr("输入搜索关键词..."))
        self.search_input.lineEdit().returnPressed.connect(self.start_search)
        # 填充历史记录
        if self.search_history:
            self.search_input.addItems(self.search_history)
        search_options.addWidget(self.search_input, 1)  # 添加stretch factor，让搜索框可以拉伸
        
        # 搜索按钮
        self.search_btn = QPushButton(tr("🔍 搜索"))
        self.search_btn.clicked.connect(self.start_search)
        search_options.addWidget(self.search_btn)
        
        # 停止按钮
        self.stop_btn = QPushButton(tr("⏹ 停止"))
        self.stop_btn.clicked.connect(self.stop_search)
        self.stop_btn.setEnabled(False)
        search_options.addWidget(self.stop_btn)
        
        layout.addLayout(search_options)
        
        # 搜索路径输入框（可编辑）
        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel(tr("搜索路径:")))
        self.path_input = QLineEdit(search_path)
        self.path_input.setStyleSheet("QLineEdit { color: #0066cc; font-weight: bold; padding: 5px; }")
        self.path_input.setPlaceholderText(tr("输入要搜索的文件夹路径..."))
        path_layout.addWidget(self.path_input)
        layout.addLayout(path_layout)
        
        # 搜索类型选择
        type_options = QHBoxLayout()
        self.search_filename_cb = QCheckBox(tr("搜索文件名"))
        self.search_filename_cb.setChecked(True)
        type_options.addWidget(self.search_filename_cb)
        
        self.search_content_cb = QCheckBox(tr("搜索文件内容"))
        self.search_content_cb.setChecked(True)  # 默认也选中
        type_options.addWidget(self.search_content_cb)

        self.match_case_cb = QCheckBox(tr("区分大小写"))
        self.match_case_cb.setToolTip(tr("区分大小写匹配文件名和文件内容"))
        type_options.addWidget(self.match_case_cb)

        self.match_whole_word_cb = QCheckBox(tr("全词匹配"))
        self.match_whole_word_cb.setToolTip(tr("仅匹配完整单词，避免命中更长字符串的一部分"))
        type_options.addWidget(self.match_whole_word_cb)
        
        # Everything搜索选项
        self.use_everything_cb = QCheckBox(tr("使用 Everything (极速)"))
        if self.everything_path:
            self.use_everything_cb.setChecked(True)  # 如果有Everything，默认启用
            self.use_everything_cb.setToolTip(tr("使用Everything搜索引擎\n路径: {}\n只搜索文件名，速度极快").format(self.everything_path))
        else:
            self.use_everything_cb.setEnabled(False)
            self.use_everything_cb.setToolTip(tr("未检测到Everything，请从 https://www.voidtools.com/ 下载安装"))
        self.use_everything_cb.stateChanged.connect(self.on_everything_toggled)
        type_options.addWidget(self.use_everything_cb)

        # 手动轻量模式：强制降级元数据以提升吞吐
        self.force_lightweight_cb = QCheckBox(tr("轻量模式(更快)"))
        self.force_lightweight_cb.setChecked(False)
        self.force_lightweight_cb.setToolTip(tr("勾选后搜索结果将优先显示核心路径信息，可能省略修改时间/大小"))
        type_options.addWidget(self.force_lightweight_cb)
        
        type_options.addStretch(1)
        layout.addLayout(type_options)
        
        # 文件类型过滤
        file_type_layout = QHBoxLayout()
        file_type_layout.addWidget(QLabel(tr("文件类型:")))
        self.file_type_input = QLineEdit()
        self.file_type_input.setPlaceholderText(tr("例如: *.c,*.h,*.xml (留空表示搜索所有类型)"))
        self.file_type_input.setText("*.c,*.h,*.xdm,*.arxml,*.xml")  # 默认值
        self.file_type_input.setStyleSheet("QLineEdit { padding: 5px; }")
        file_type_layout.addWidget(self.file_type_input)
        layout.addLayout(file_type_layout)
        
        # 状态标签
        self.status_label = QLabel(tr("就绪"))
        layout.addWidget(self.status_label)
        
        # 结果表格
        self.result_model = SearchResultsTableModel(self)
        self.result_list = QTableView()
        self.result_list.setModel(self.result_model)
        self.result_list.horizontalHeader().setStretchLastSection(False)
        self.result_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.result_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.result_list.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.result_list.horizontalHeader().setSectionResizeMode(3, QHeaderView.Fixed)
        self.result_list.setColumnWidth(1, SEARCH_RESULT_TYPE_COL_WIDTH)
        self.result_list.setColumnWidth(2, SEARCH_RESULT_DATE_COL_WIDTH)
        self.result_list.setColumnWidth(3, SEARCH_RESULT_SIZE_COL_WIDTH)
        self.result_list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.result_list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.result_list.setWordWrap(False)
        self.result_list.doubleClicked.connect(self.on_result_double_clicked)
        self.result_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_list.customContextMenuRequested.connect(self.show_result_context_menu)
        # 启用排序功能
        self.result_list.setSortingEnabled(True)
        # 设置自定义委托，让文件名列的省略号显示在开头
        self.result_list.setItemDelegateForColumn(0, ElideLeftDelegate(self.result_list))
        # 设置行高和网格线
        self.result_list.verticalHeader().setDefaultSectionSize(24)  # 设置默认行高为24像素
        self.result_list.setShowGrid(True)  # 显示网格线
        self.result_list.setAlternatingRowColors(True)  # 启用交替行颜色
        # 设置表头样式
        self.result_list.setStyleSheet("""
            QHeaderView::section {
                background-color: #E0E0E0;
                padding: 4px;
                border: 1px solid #C0C0C0;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.result_list)
        
        # 启动UI更新定时器
        from PyQt5.QtCore import QTimer
        self.ui_update_timer = QTimer(self)
        self.ui_update_timer.timeout.connect(self.update_ui_from_queue)
        self.ui_update_timer.start(80)  # 搜索时动态提速，空闲时降频

    def _drain_result_queue(self):
        if not self.result_queue:
            return
        while True:
            try:
                self.result_queue.get_nowait()
            except Exception:
                break

    def _release_search_resources(self):
        self.is_searching = False
        self.queue_overflow_count = 0
        self._queue_idle_ticks = 0

        if self.ui_update_timer:
            self.ui_update_timer.stop()

        # 等待后台搜索线程优雅退出，避免关闭后仍占用磁盘IO
        search_thread = getattr(self, 'search_thread', None)
        if (search_thread and search_thread.is_alive()
                and search_thread is not threading.current_thread()):
            search_thread.join(timeout=2.0)

        self._drain_result_queue()

        if hasattr(self, 'result_model') and self.result_model:
            self.result_model.clear()
        self.current_result_count = 0
        self.search_thread = None

    def closeEvent(self, event):
        self._release_search_resources()
        super().closeEvent(event)

    def _ensure_ui_update_timer(self, interval=None):
        if not self.ui_update_timer:
            return
        if interval is not None and self.ui_update_timer.interval() != interval:
            self.ui_update_timer.setInterval(interval)
        if not self.ui_update_timer.isActive():
            self.ui_update_timer.start()
    
    def update_ui_from_queue(self):
        """从队列中取出结果并更新UI（在主线程中调用，批量优化）"""
        try:
            pending = 0
            try:
                pending = self.result_queue.qsize()
            except Exception:
                pending = 0

            # 自适应轮询：有积压时提速，空闲时降频，减少空转开销
            target_interval = 80
            if self.is_searching:
                if pending > 600:
                    target_interval = 20
                elif pending > 120:
                    target_interval = 35
                else:
                    target_interval = 60
            else:
                target_interval = 220 if pending == 0 else 80
            if self.ui_update_timer.interval() != target_interval:
                self.ui_update_timer.setInterval(target_interval)

            # 批量处理结果（一次处理最多200个，加快队列消费）
            batch_results = []
            batch_count = 0
            # 根据队列积压动态调整消费批次，减少结果爆发时的UI延迟
            max_batch = 200
            try:
                if pending > 800:
                    max_batch = 500
                elif pending > 300:
                    max_batch = 350
            except Exception:
                pass
            
            while batch_count < max_batch:
                try:
                    item = self.result_queue.get_nowait()
                    
                    if item['type'] == 'result':
                        batch_results.append(item)
                        batch_count += 1
                    elif item['type'] == 'result_batch':
                        items = item.get('items') or []
                        if items:
                            batch_results.extend(items)
                            batch_count += len(items)
                    elif item['type'] == 'status':
                        self.status_label.setText(item['text'])
                    elif item['type'] == 'button':
                        if item['button'] == 'search':
                            self.search_btn.setEnabled(item['enabled'])
                        elif item['button'] == 'stop':
                            self.stop_btn.setEnabled(item['enabled'])
                    elif item['type'] == 'enable_sorting':
                        # 搜索完成后启用排序
                        self.result_list.setSortingEnabled(True)
                except Exception:
                    break  # 队列为空
            
            # 批量添加结果到表格（性能优化）
            if batch_results:
                self._append_results_to_table(batch_results)
                self._queue_idle_ticks = 0
            elif not self.is_searching:
                self._queue_idle_ticks += 1
                if self._queue_idle_ticks > 3 and pending == 0 and self.ui_update_timer.isActive():
                    self.ui_update_timer.stop()
                elif self._queue_idle_ticks > 1 and self.ui_update_timer.interval() != 250:
                    self.ui_update_timer.setInterval(250)
                
        except Exception as e:
            debug_print(f"[Search] UI update error: {e}")

    def _append_results_to_table(self, results):
        """批量渲染搜索结果到表格。"""
        from PyQt5.QtCore import Qt

        if not results:
            return

        remaining_capacity = self.max_results - self.current_result_count
        if remaining_capacity <= 0:
            return
        rows_to_add = results[:remaining_capacity]

        self.result_list.setUpdatesEnabled(False)
        added_count = self.result_model.append_results(rows_to_add)
        self.current_result_count += added_count
        self.result_list.setUpdatesEnabled(True)
    
    def add_search_result(self, item):
        """添加单条搜索结果（通过队列，线程安全）。"""
        self.result_queue.put({'type': 'result', **item})

    def add_search_results_batch(self, items, timeout=0.5):
        """批量添加搜索结果到队列，减少高并发下的队列争用。"""
        if not items:
            return
        self.result_queue.put({'type': 'result_batch', 'items': items}, timeout=timeout)
    
    def on_everything_toggled(self, state):
        """当Everything选项切换时"""
        if state:
            # 启用Everything时，禁用文件内容搜索（Everything只支持文件名）
            self.search_content_cb.setChecked(False)
            self.search_content_cb.setEnabled(False)
        else:
            # 禁用Everything时，恢复文件内容搜索选项
            self.search_content_cb.setEnabled(True)
    
    def clear_search_cache(self):
        """清除所有搜索缓存（内部使用，软件关闭时自动调用）"""
        global _search_cache
        _search_cache.clear()
        debug_print(tr("[Search] 搜索缓存已清除"))
    
    def start_search(self):
        keyword = self.search_input.currentText().strip()  # 改用currentText获取输入或选中的文本
        keyword_is_empty = not keyword
        
        # 仅在有关键词时才添加到历史记录
        if not keyword_is_empty:
            if self.main_window and hasattr(self.main_window, 'add_search_history'):
                self.main_window.add_search_history(keyword)
                # 更新下拉列表
                self.search_input.clear()
                if hasattr(self.main_window, 'search_history'):
                    self.search_input.addItems(self.main_window.search_history)
                # 设置当前文本为刚刚搜索的关键词
                self.search_input.setCurrentText(keyword)
        
        # 空关键词时：禁用内容搜索（搜索空内容无意义），确保文件名搜索开启
        do_search_filename = self.search_filename_cb.isChecked()
        do_search_content = self.search_content_cb.isChecked() and not keyword_is_empty
        if keyword_is_empty and not do_search_filename:
            do_search_filename = True  # 空关键词时自动启用文件名搜索
        
        if not do_search_filename and not do_search_content:
            show_toast(self, tr("提示"), tr("请至少选择一种搜索类型"), level="warning")
            return
        
        # 获取并验证搜索路径
        search_path = self.path_input.text().strip()
        if not search_path:
            show_toast(self, tr("提示"), tr("请输入搜索路径"), level="warning")
            return
        
        # 检查路径是否存在
        if not os.path.exists(search_path):
            show_toast(self, tr("路径错误"), tr("路径不存在:\n{}").format(search_path), level="warning")
            return
        
        # 检查是否是目录
        if not os.path.isdir(search_path):
            show_toast(self, tr("路径错误"), tr("路径不是文件夹:\n{}").format(search_path), level="warning")
            return
        
        # 检查是否是特殊路径（不支持搜索）
        if search_path.startswith('shell:'):
            show_toast(self, tr("不支持"), tr("不支持搜索特殊路径（shell:）"), level="warning")
            return
        
        # 更新搜索路径
        self.search_path = search_path
        
        # 获取文件类型过滤
        file_types = self.file_type_input.text().strip()
        
        # 检查缓存
        global _search_cache
        force_metadata_degrade = self.force_lightweight_cb.isChecked()
        cache_key = _search_cache.get_key(
            search_path, keyword,
            do_search_filename,
            do_search_content,
            file_types,
            force_metadata_degrade,
            self.match_case_cb.isChecked(),
            self.match_whole_word_cb.isChecked(),
        )
        cached_results = _search_cache.get(cache_key)
        
        if cached_results is not None:
            # 使用缓存结果
            debug_print(f"[Search] 使用缓存结果，共 {len(cached_results)} 个")
            self.result_model.clear()
            self.current_result_count = 0
            self.status_label.setText(tr("正在加载缓存结果..."))
            
            # 批量添加缓存结果
            sorting_enabled = self.result_list.isSortingEnabled()
            self.result_list.setSortingEnabled(False)

            self._append_results_to_table(cached_results)
            
            self.result_list.setSortingEnabled(sorting_enabled)
            shown_count = min(len(cached_results), self.max_results)
            self.status_label.setText(tr("搜索完成（缓存），共显示 {} 个结果").format(shown_count))
            if self.ui_update_timer and self.ui_update_timer.isActive():
                self.ui_update_timer.stop()
            return
        
        # 清空之前的结果
        self.result_model.clear()
        self.current_result_count = 0  # 重置计数器
        
        # 搜索期间完全禁用排序（性能优化）
        self.result_list.setSortingEnabled(False)
        
        self.is_searching = True
        self.search_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        if keyword_is_empty:
            file_types_hint = f"（{file_types}）" if file_types else ""
            self.status_label.setText(tr("列举文件中{}... (最多显示{}个结果)").format(file_types_hint, self.max_results))
        else:
            self.status_label.setText(tr("搜索中... (最多显示{}个结果)").format(self.max_results))
        if self.ui_update_timer:
            self._queue_idle_ticks = 0
            self._ensure_ui_update_timer(20)
        
        # 在后台线程执行搜索
        import threading
        use_everything = self.use_everything_cb.isChecked() if self.everything_path else False
        match_case = self.match_case_cb.isChecked()
        match_whole_word = self.match_whole_word_cb.isChecked()
        self.search_thread = threading.Thread(
            target=self.do_search,
            args=(keyword, do_search_filename, do_search_content, file_types, cache_key, use_everything, force_metadata_degrade, match_case, match_whole_word)
        )
        self.search_thread.daemon = True
        self.search_thread.start()
    
    def stop_search(self):
        self.is_searching = False
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText(tr("已停止"))
        if self.ui_update_timer:
            self._queue_idle_ticks = 0
            self._ensure_ui_update_timer(220)
    
    def search_with_everything(self, keyword, search_path, file_types="", match_case=False, match_whole_word=False):
        """使用Everything进行搜索"""
        import subprocess
        import re

        def _matches_filename(name):
            if match_whole_word:
                flags = 0 if match_case else re.IGNORECASE
                return bool(re.search(rf'\b{re.escape(keyword)}\b', name, flags))
            if match_case:
                return keyword in name
            return keyword.lower() in name.lower()
        
        try:
            # 构建Everything命令
            cmd = [self.everything_path]
            
            # 如果指定了搜索路径，添加路径过滤
            if search_path and os.path.exists(search_path):
                # Everything使用path:语法指定路径
                search_pattern = f'path:"{search_path}" {keyword}'
            else:
                search_pattern = keyword
            
            # 添加文件类型过滤
            if file_types:
                extensions = []
                for ft in file_types.split(','):
                    ft = ft.strip()
                    if ft.startswith('*.'):
                        ft = ft[2:]  # 移除 *.
                    extensions.append(f'ext:{ft}')
                if extensions:
                    search_pattern += ' ' + ' | '.join(extensions)
            
            cmd.append(search_pattern)
            
            # 执行Everything搜索
            if _DEBUG_MODE:
                debug_print(f"[Everything] Executing: {' '.join(cmd)}")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore',
                timeout=30
            )
            
            if result.returncode == 0:
                # 解析结果（每行一个文件路径）
                lines = result.stdout.strip().split('\n')
                results = []
                for line in lines:
                    line = line.strip()
                    # es.exe 输出本身来自索引，避免逐条 exists 造成大量额外 I/O
                    if line and _matches_filename(os.path.basename(line)):
                        results.append(line)
                
                debug_print(f"[Everything] Found {len(results)} results")
                return results
            else:
                debug_print(f"[Everything] Error: {result.stderr}")
                return []
        
        except subprocess.TimeoutExpired:
            debug_print("[Everything] Search timeout")
            return []
        except Exception as e:
            debug_print(f"[Everything] Error: {e}")
            return []
    
    def do_search(self, keyword, search_filename, search_content, file_types="", cache_key=None, use_everything=False, force_metadata_degrade=False, match_case=False, match_whole_word=False):
        import re

        metadata_degrade_count = 0
        whole_word_pattern = None
        whole_word_pattern_bytes = None
        if match_whole_word:
            whole_word_flags = 0 if match_case else re.IGNORECASE
            whole_word_pattern = re.compile(rf'\b{re.escape(keyword)}\b', whole_word_flags)
            if keyword.isascii():
                byte_flags = 0 if match_case else re.IGNORECASE
                whole_word_pattern_bytes = re.compile(rb'\b' + re.escape(keyword.encode('ascii', errors='ignore')) + rb'\b', byte_flags)

        def _matches_text(value):
            if not isinstance(value, str) or not value:
                return False
            if whole_word_pattern is not None:
                return bool(whole_word_pattern.search(value))
            if match_case:
                return keyword in value
            return keyword.lower() in value.lower()

        def _matches_bytes(value):
            if not value:
                return False
            if whole_word_pattern_bytes is not None:
                return bool(whole_word_pattern_bytes.search(value))
            if match_case:
                return keyword_bytes in value
            return keyword_bytes_lower in value.lower()

        def _should_degrade_metadata():
            if force_metadata_degrade:
                return True
            if not SEARCH_METADATA_DEGRADE_ENABLED:
                return False
            try:
                pending = self.result_queue.qsize()
                queue_max = max(1, self.result_queue.maxsize)
                return (pending / queue_max) >= SEARCH_METADATA_DEGRADE_QUEUE_RATIO
            except Exception:
                return False

        # 如果使用Everything搜索
        if use_everything and self.everything_path:
            self.result_queue.put({'type': 'status', 'text': tr('Using Everything搜索引擎...')})
            
            try:
                results = self.search_with_everything(keyword, self.search_path, file_types, match_case=match_case, match_whole_word=match_whole_word)
                
                if not self.is_searching:
                    return
                
                # 将Everything结果添加到显示队列
                batch_size = 200
                batch_items = []
                for file_path in results:
                    if not self.is_searching:
                        break
                    
                    try:
                        basename = os.path.basename(file_path)
                        name_without_ext, file_ext = os.path.splitext(basename)
                        path_without_ext = os.path.join(os.path.dirname(file_path), name_without_ext)
                        file_type = file_ext[1:].upper() if file_ext else tr("无")
                        sort_date_ts = None
                        sort_size_bytes = None
                        if _should_degrade_metadata():
                            metadata_degrade_count += 1
                            mtime = "-"
                            size_str = "-"
                        else:
                            try:
                                stat_info = os.stat(file_path)
                                mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                                size_bytes = stat_info.st_size
                                sort_date_ts = stat_info.st_mtime
                                sort_size_bytes = size_bytes
                                size_str = format_file_size(size_bytes)
                            except Exception:
                                mtime = "-"
                                size_str = "-"

                        batch_items.append({
                            'path': file_path,
                            'name': f"📄 {path_without_ext}",
                            'full_path': f"📄 {file_path}",
                            'file_type': file_type,
                            'date': mtime,
                            'size': size_str,
                            'sort_date_ts': sort_date_ts,
                            'sort_size_bytes': sort_size_bytes,
                        })
                        if len(batch_items) >= batch_size:
                            self.add_search_results_batch(batch_items, timeout=0.5)
                            batch_items = []
                    except Exception:
                        pass

                if batch_items:
                    try:
                        self.add_search_results_batch(batch_items, timeout=0.5)
                    except Exception:
                        pass
                
                # 搜索完成
                final_count = len(results)
                degrade_note = tr("，元数据降级 {} 条").format(metadata_degrade_count) if metadata_degrade_count > 0 else ""
                self.result_queue.put({'type': 'status', 'text': tr("Everything搜索完成，共找到 {} 个结果{}").format(final_count, degrade_note)})
                
            except Exception as e:
                self.result_queue.put({'type': 'error', 'text': f'Everything搜索错误: {str(e)}'})
            
            return
        
        # 原有的搜索逻辑
        found_count = 0
        keyword_lower = keyword.lower()
        keyword_is_ascii = keyword.isascii()
        keyword_bytes = keyword.encode('ascii', errors='ignore') if keyword_is_ascii else b''
        keyword_bytes_lower = keyword_lower.encode('ascii', errors='ignore') if keyword_is_ascii else b''
        results_buffer = []  # 结果缓冲区
        base_buffer_size = SEARCH_RESULT_BATCH_BASE
        max_cached_results = MAX_CACHED_RESULTS_PER_QUERY
        all_results = [] if cache_key else None  # 仅在需要缓存时保存结果
        ext_encoding_cache = {}  # 扩展名 -> 最近成功编码（减少重复试错）

        def _adaptive_buffer_size():
            """根据结果队列积压动态调整发送批次，降低高压场景争用。"""
            try:
                pending = self.result_queue.qsize()
                queue_max = max(1, self.result_queue.maxsize)
                ratio = pending / queue_max
                if ratio >= 0.75:
                    return min(SEARCH_RESULT_BATCH_MAX, base_buffer_size * 3)
                if ratio >= 0.4:
                    return min(SEARCH_RESULT_BATCH_MAX, base_buffer_size * 2)
                if ratio <= 0.1:
                    return max(SEARCH_RESULT_BATCH_MIN, base_buffer_size // 2)
            except Exception:
                pass
            return base_buffer_size

        def _flush_results_buffer(force=False):
            nonlocal results_buffer
            if not results_buffer:
                return
            if not force:
                threshold = _adaptive_buffer_size()
                if len(results_buffer) < threshold:
                    return
            try:
                self.add_search_results_batch(results_buffer, timeout=0.5)
            except Exception:
                self.queue_overflow_count += len(results_buffer)
            results_buffer = []

        # 结果限制（防止内存溢出和UI卡死）
        max_results = self.max_results
        results_limited = False
        
        # 二进制文件扩展名黑名单（这些文件肯定不搜索内容）
        binary_file_extensions = {
            # 可执行文件
            'exe', 'dll', 'so', 'dylib', 'bin', 'com', 'app',
            # 归档/压缩文件
            'zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'xz', 'iso', 'dmg',
            # 图片文件
            'jpg', 'jpeg', 'png', 'gif', 'bmp', 'ico', 'svg', 'webp', 'tiff', 'psd', 'ai',
            # 音频文件
            'mp3', 'wav', 'flac', 'aac', 'ogg', 'wma', 'm4a',
            # 视频文件
            'mp4', 'avi', 'mkv', 'mov', 'wmv', 'flv', 'webm', 'mpeg', 'mpg',
            # Office文件（二进制格式）
            'doc', 'xls', 'ppt', 'docx', 'xlsx', 'pptx', 'pdf',
            # 数据库文件
            'db', 'sqlite', 'mdb', 'accdb',
            # 其他二进制
            'obj', 'o', 'a', 'lib', 'pyc', 'pyo', 'class', 'jar', 'war',
        }

        # 文本扩展名白名单：命中时跳过二进制探测，减少一次额外文件读取
        text_file_extensions = {
            'txt', 'md', 'rst', 'log', 'ini', 'cfg', 'conf', 'toml', 'yaml', 'yml', 'json', 'xml',
            'csv', 'tsv', 'sql', 'bat', 'ps1', 'sh', 'c', 'h', 'cpp', 'hpp', 'cc', 'cs', 'java',
            'py', 'js', 'ts', 'jsx', 'tsx', 'html', 'htm', 'css', 'scss', 'less', 'go', 'rs', 'php',
            'rb', 'swift', 'kt', 'm', 'mm', 'vue', 'svelte', 'dockerfile', 'gitignore', 'arxml', 'xdm'
        }
        
        # 对于文件内容搜索，使用编译的正则表达式可能更快（可选优化）
        # 但Python的内置字符串搜索已经很快，这里保持简单
        
        # 解析文件类型过滤（支持*.ext格式，逗号分隔）
        file_extensions = set()
        if file_types:
            for ft in file_types.split(','):
                ft = ft.strip()
                if ft.startswith('*.'):
                    file_extensions.add(ft[2:].lower())  # 去掉*.，只保留扩展名
                elif ft.startswith('.'):
                    file_extensions.add(ft[1:].lower())  # 去掉.，只保留扩展名
                elif ft:
                    file_extensions.add(ft.lower())  # 直接使用输入的扩展名
        
        # 调试信息：输出搜索路径
        debug_print(tr("[Search] 开始搜索路径: {}").format(self.search_path))
        debug_print(tr("[Search] 搜索关键词: {}").format(keyword))
        debug_print(tr("[Search] 搜索文件名: {}, 搜索内容: {}").format(search_filename, search_content))
        debug_print(tr("[Search] 文件类型过滤: {}").format(file_extensions if file_extensions else '所有类型'))
        
        try:
            scanned_files = 0
            folder_count = 0
            skipped_binary_files = 0  # 跳过的二进制文件数
            last_status_update_ms = int(time.time() * 1000)
            for root, dirs, files in os.walk(self.search_path):
                if not self.is_searching:
                    debug_print(tr("[Search] 搜索被中断"))
                    break
                
                folder_count += 1
                
                # 搜索文件夹名（空关键词时跳过目录，因为目录无扩展名无法匹配文件类型过滤）
                if search_filename and keyword:
                    for dirname in dirs:
                        if not self.is_searching:
                            break
                        
                        # 检查是否达到结果限制
                        if found_count >= max_results:
                            results_limited = True
                            break
                        
                        # 使用Python内置的字符串搜索（已优化）
                        if _matches_text(dirname):
                            found_count += 1
                            dir_path = os.path.join(root, dirname)
                            
                            # 获取文件夹信息
                            if _should_degrade_metadata():
                                metadata_degrade_count += 1
                                mtime = "-"
                                size_str = "-"
                                sort_date_ts = None
                            else:
                                try:
                                    stat_info = os.stat(dir_path)
                                    mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                                    size_str = "-"  # 文件夹不显示大小
                                    sort_date_ts = stat_info.st_mtime
                                except Exception:
                                    mtime = "-"
                                    size_str = "-"
                                    sort_date_ts = None
                            
                            result_item = {
                                'path': dir_path,
                                'name': f"📁 {dirname}",
                                'full_path': f"📁 {dir_path}",
                                'file_type': tr('文件夹'),
                                'date': mtime,
                                'size': size_str,
                                'sort_date_ts': sort_date_ts,
                                'sort_size_bytes': None,
                            }
                            results_buffer.append(result_item)
                            if all_results is not None and len(all_results) < max_cached_results:
                                all_results.append(result_item)  # 保存到缓存列表
                            
                            # 批量更新UI（队列满时等待）
                            _flush_results_buffer()
                
                # 检查是否达到结果限制
                if found_count >= max_results:
                    results_limited = True
                    break
                
                # 搜索文件名和文件内容
                for filename in files:
                    if not self.is_searching:
                        debug_print(tr("[Search] 搜索被中断（文件循环）"))
                        break

                    filename_lower = filename.lower()
                    name_without_ext, ext = os.path.splitext(filename)
                    file_ext = ext[1:].lower() if ext else ''
                    
                    # 检查是否达到结果限制
                    if found_count >= max_results:
                        results_limited = True
                        break
                    
                    # 优化：如果只搜索文件名，快速过滤不匹配的文件
                    if search_filename and not search_content:
                        if not _matches_text(filename):
                            continue  # 文件名不匹配，跳过
                    
                    # 检查文件类型过滤
                    if file_extensions and file_ext not in file_extensions:
                        # 调试：显示被过滤的文件（仅对特定文件名）
                        if 'TstMgr' in filename or scanned_files < 5:
                            debug_print(tr("[Search] 文件被类型过滤跳过: {}").format(filename))
                        continue  # 跳过不匹配的文件类型
                    
                    scanned_files += 1
                    
                    # 状态更新节流：每300个文件必更新；否则每32个文件检查一次时间阈值
                    should_update_status = False
                    if scanned_files % 300 == 0:
                        should_update_status = True
                    elif scanned_files % 32 == 0:
                        now_ms = int(time.time() * 1000)
                        if (now_ms - last_status_update_ms) >= 500:
                            should_update_status = True

                    if should_update_status:
                        status_text = tr("搜索中... 已扫描 {} 个文件，找到 {} 个结果").format(scanned_files, found_count)
                        try:
                            self.result_queue.put({'type': 'status', 'text': status_text}, timeout=0.1)
                            last_status_update_ms = int(time.time() * 1000)
                        except Exception:
                            pass  # 超时后继续搜索
                    
                    file_path = os.path.join(root, filename)
                    matched = False
                    match_type = ""
                    
                    # 搜索文件名（Python内置优化）
                    if search_filename and _matches_text(filename):
                        matched = True
                        match_type = "📄"
                    
                    # 搜索文件内容（智能检测文本文件）
                    if search_content and not matched:
                        # 1. 首先检查黑名单（明确的二进制文件）
                        if file_ext in binary_file_extensions:
                            skipped_binary_files += 1
                            continue

                        # 2. 预取文件大小：空文件不可能命中关键词，直接跳过；同时复用size避免重复stat
                        try:
                            file_size = os.path.getsize(file_path)
                        except OSError:
                            continue
                        if file_size == 0:
                            continue

                        # 3. 文本白名单直接通过；其余文件走探测
                        if file_ext not in text_file_extensions and not is_text_file(file_path):
                            skipped_binary_files += 1
                            continue
                        
                        try:
                            # 文件大小已在上方预取（file_size）

                            # 分块读取与单文件扫描上限
                            chunk_size = CONTENT_SEARCH_CHUNK_SIZE
                            max_scan_bytes = CONTENT_SEARCH_MAX_BYTES_PER_FILE
                            in_memory_threshold = CONTENT_SEARCH_IN_MEMORY_THRESHOLD

                            # ASCII关键词快速路径：直接按字节匹配，跳过多编码解码
                            if keyword_is_ascii and keyword_bytes:
                                read_limit = min(file_size, max_scan_bytes)
                                if read_limit <= in_memory_threshold:
                                    with open(file_path, 'rb') as bf:
                                        raw_content = bf.read(read_limit)
                                    # 兼容 UTF-16(无BOM) 等含 NULL 字节文本：移除 NULL 后再匹配一次
                                    raw_no_null = raw_content.replace(b'\x00', b'')
                                    if (
                                        _matches_bytes(raw_content)
                                        or _matches_bytes(raw_no_null)
                                    ):
                                        matched = True
                                        match_type = "📄"
                                else:
                                    overlap_bytes = max(1, len(keyword_bytes) * 2)
                                    scanned_bytes = 0
                                    with open(file_path, 'rb') as bf:
                                        while True:
                                            if scanned_bytes >= max_scan_bytes:
                                                break
                                            chunk = bf.read(chunk_size)
                                            if not chunk:
                                                break
                                            scanned_bytes += len(chunk)
                                            chunk_no_null = chunk.replace(b'\x00', b'')
                                            if (
                                                _matches_bytes(chunk)
                                                or _matches_bytes(chunk_no_null)
                                            ):
                                                matched = True
                                                match_type = "📄"
                                                break
                                            if len(chunk) == chunk_size:
                                                bf.seek(bf.tell() - overlap_bytes)

                            # 编码顺序：优先使用该扩展名最近成功编码
                            if not matched:  # 跳过编码尝试如果已匹配
                                preferred_encoding = ext_encoding_cache.get(file_ext)
                                if preferred_encoding and preferred_encoding in CONTENT_SEARCH_ENCODINGS:
                                    encodings = [preferred_encoding] + [enc for enc in CONTENT_SEARCH_ENCODINGS if enc != preferred_encoding]
                                else:
                                    encodings = list(CONTENT_SEARCH_ENCODINGS)
                                content_matched = False
                                
                                for encoding in encodings:
                                    try:
                                        read_limit = min(file_size, max_scan_bytes)
                                        if read_limit <= in_memory_threshold:
                                            # 小文件优化：只读一次二进制，再在内存中尝试不同编码，避免重复磁盘I/O
                                            if encoding == encodings[0]:
                                                with open(file_path, 'rb') as bf:
                                                    raw_content = bf.read(read_limit)
                                            content = raw_content.decode(encoding, errors='ignore')
                                            if _matches_text(content):
                                                matched = True
                                                match_type = "📄"
                                                content_matched = True
                                                if file_ext:
                                                    ext_encoding_cache[file_ext] = encoding
                                                break
                                        else:
                                            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                                                # 大文件分块读取（有总扫描上限）
                                                overlap = len(keyword) * 2  # 重叠区域，防止关键词被分割
                                                scanned_bytes = 0
                                                while True:
                                                    if scanned_bytes >= max_scan_bytes:
                                                        break
                                                    chunk = f.read(chunk_size)
                                                    if not chunk:
                                                        break
                                                    scanned_bytes += len(chunk)
                                                    if _matches_text(chunk):
                                                        matched = True
                                                        match_type = "📄"
                                                        content_matched = True
                                                        if file_ext:
                                                            ext_encoding_cache[file_ext] = encoding
                                                        break
                                                    # 回退overlap字节，避免关键词跨块
                                                    if len(chunk) == chunk_size:
                                                        f.seek(f.tell() - overlap)
                                                if content_matched:
                                                    break
                                    except UnicodeError:
                                        continue
                                    except Exception as e:
                                        # 其他错误，记录日志并尝试下一个编码
                                        debug_print(tr("[Search] 读取文件失败 {} (编码 {}): {}").format(file_path, encoding, e))
                                        continue
                        except Exception as e:
                            # 如果无法以文本方式读取，记录日志并跳过该文件
                            debug_print(tr("[Search] 无法读取文件 {}: {}").format(file_path, e))
                            pass
                    
                    if matched:
                        found_count += 1
                        
                        # 获取文件信息
                        sort_date_ts = None
                        sort_size_bytes = None
                        if _should_degrade_metadata():
                            metadata_degrade_count += 1
                            mtime = "-"
                            size_str = "-"
                        else:
                            try:
                                stat_info = os.stat(file_path)
                                mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                                size_bytes = stat_info.st_size
                                sort_date_ts = stat_info.st_mtime
                                sort_size_bytes = size_bytes
                                # 格式化大小
                                size_str = format_file_size(size_bytes)
                            except Exception:
                                mtime = "-"
                                size_str = "-"
                        
                        # 获取不带扩展名的完整路径
                        path_without_ext = os.path.join(root, name_without_ext)
                        file_type = file_ext.upper() if file_ext else tr("无")
                        
                        result_item = {
                            'path': file_path,
                            'name': f"{match_type} {path_without_ext}",
                            'full_path': f"{match_type} {file_path}",
                            'file_type': file_type,
                            'date': mtime,
                            'size': size_str,
                            'sort_date_ts': sort_date_ts,
                            'sort_size_bytes': sort_size_bytes,
                        }
                        results_buffer.append(result_item)
                        if all_results is not None and len(all_results) < max_cached_results:
                            all_results.append(result_item)  # 保存到缓存列表
                        
                        # 批量更新UI（每20个结果更新一次）
                        _flush_results_buffer()
        except Exception as e:
            debug_print(f"[Search] error: {e}")
        
        # 添加剩余的结果（队列满时等待）
        if results_buffer:
            _flush_results_buffer(force=True)
        
        # 调试信息
        debug_print(tr("[Search] 搜索完成，共扫描 {} 个文件，找到 {} 个结果").format(scanned_files, found_count))
        if search_content and skipped_binary_files > 0:
            debug_print(tr("[Search] 跳过 {} 个二进制文件（不搜索内容）").format(skipped_binary_files))
        if self.queue_overflow_count > 0:
            debug_print(tr("[Search] ⚠️ 队列溢出 {} 次（部分结果未显示）").format(self.queue_overflow_count))
        
        # 将结果存入缓存（限制缓存大小，防止内存溢出）
        if cache_key and all_results:
            global _search_cache
            cached_results = all_results
            _search_cache.put(cache_key, cached_results)
            debug_print(f"[Search] 已将 {len(cached_results)} 个结果存入缓存")
        
        # 重置搜索状态（先重置，避免后续更新被跳过）
        self.is_searching = False
        self.queue_overflow_count = 0  # 重置溢出计数
        
        # 搜索完成，更新UI状态（使用带超时的put，避免卡死）
        if results_limited:
            final_status = tr("搜索完成（已限制），显示前 {} 个结果（扫描了 {} 个文件）⚠️").format(found_count, scanned_files)
        else:
            final_status = tr("搜索完成，共找到 {} 个结果（扫描了 {} 个文件）").format(found_count, scanned_files)
        if metadata_degrade_count > 0:
            final_status += tr("，元数据降级 {} 条").format(metadata_degrade_count)
        
        # 使用超时put，防止队列满时卡死
        try:
            self.result_queue.put({'type': 'status', 'text': final_status}, timeout=1)
            self.result_queue.put({'type': 'button', 'button': 'search', 'enabled': True}, timeout=1)
            self.result_queue.put({'type': 'button', 'button': 'stop', 'enabled': False}, timeout=1)
            # 搜索完成后启用排序
            self.result_queue.put({'type': 'enable_sorting'}, timeout=1)
        except Exception:
            debug_print(tr('[Search] ⚠️ 队列满，最终状态更新失败'))
        
        debug_print(tr('[Search] UI更新已调度（使用队列）'))
    
    def on_result_double_clicked(self, index):
        """双击搜索结果，打开文件所在文件夹或文件夹本身，并选中文件"""
        if not index.isValid():
            return
        file_path = self.result_model.path_for_row(index.row())
        if file_path and os.path.exists(file_path):
            # 如果是文件夹，直接打开文件夹；如果是文件，打开文件所在文件夹并选中文件
            if os.path.isdir(file_path):
                folder_path = file_path
                select_file = None
            else:
                folder_path = os.path.dirname(file_path)
                select_file = os.path.basename(file_path)  # 要选中的文件名
            # 不关闭搜索对话框，保持独立
            if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                self.main_window.add_new_tab(folder_path, select_file=select_file)

    def _launch_file_with_program(self, file_path, program_path, display_name):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("该搜索结果不是可打开的文件"), level="warning")
            return False
        if not program_path:
            show_toast(self, tr("提示"), tr("未找到 {}").format(display_name), level="warning")
            return False
        try:
            launch_detached_async([program_path, file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法使用 {} 打开文件: {}").format(display_name, e), level="error")
            return False

    def open_result_with_notepad(self, file_path):
        self._launch_file_with_program(file_path, 'notepad.exe', tr('记事本'))

    def open_result_with_notepad_plus_plus(self, file_path):
        self._launch_file_with_program(file_path, self.notepad_plus_plus_path, 'Notepad++')

    def open_result_with_default_app(self, file_path):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("该搜索结果不是可打开的文件"), level="warning")
            return False
        try:
            if os.name == 'nt':
                os.startfile(file_path)
            else:
                launch_detached([file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法使用系统默认程序打开文件: {}").format(e), level="error")
            return False

    def open_result_with_system_dialog(self, file_path):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("该搜索结果不是可打开的文件"), level="warning")
            return False
        if os.name != 'nt':
            show_toast(self, tr("提示"), tr("当前系统不支持打开“选择其他应用”对话框"), level="warning")
            return False
        try:
            launch_detached_async(['rundll32.exe', 'shell32.dll,OpenAs_RunDLL', file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开“选择其他应用”对话框: {}").format(e), level="error")
            return False

    def show_result_context_menu(self, pos):
        index = self.result_list.indexAt(pos)
        if not index.isValid():
            return

        row = index.row()
        file_path = self.result_model.path_for_row(row)
        if not file_path or not os.path.exists(file_path):
            return

        self.result_list.selectRow(row)

        menu = QMenu(self)

        open_action = menu.addAction(tr("打开"))
        open_action.triggered.connect(lambda: self.on_result_double_clicked(index))

        open_folder_action = menu.addAction(tr("打开所在目录"))
        open_folder_action.triggered.connect(lambda: self._open_result_parent_folder(file_path))

        if os.path.isfile(file_path):
            default_action = menu.addAction(tr("用系统默认程序打开"))
            default_action.triggered.connect(lambda: self.open_result_with_default_app(file_path))

            menu.addSeparator()

            notepad_action = menu.addAction(tr("用记事本打开"))
            notepad_action.triggered.connect(lambda: self.open_result_with_notepad(file_path))

            notepadpp_action = menu.addAction(tr("用 Notepad++ 打开"))
            notepadpp_action.setEnabled(bool(getattr(self, 'notepad_plus_plus_path', None)))
            if not getattr(self, 'notepad_plus_plus_path', None):
                notepadpp_action.setToolTip(tr("未检测到 Notepad++"))
            notepadpp_action.triggered.connect(lambda: self.open_result_with_notepad_plus_plus(file_path))

            menu.addSeparator()

            system_dialog_action = menu.addAction(tr("选择其他应用..."))
            system_dialog_action.triggered.connect(lambda: self.open_result_with_system_dialog(file_path))

        menu.exec_(self.result_list.viewport().mapToGlobal(pos))

    def _open_result_parent_folder(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return
        if os.path.isdir(file_path):
            folder_path = file_path
            select_file = None
        else:
            folder_path = os.path.dirname(file_path)
            select_file = os.path.basename(file_path)
        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
            self.main_window.add_new_tab(folder_path, select_file=select_file)

import sys
import os
import json
import subprocess
import string
import time
import socket
import threading
import queue
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QListWidget, QLabel, QToolBar, QAction, QMenu, QInputDialog, QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView, QSizePolicy, QFileSystemModel, QSplitter, QProgressBar, QCompleter, QFrame, QToolButton, QFileIconProvider)  # 添加QFrame
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import Qt, QDir, QUrl, pyqtSignal, pyqtSlot, Q_ARG, QObject, QSize, QFileSystemWatcher, QTimer, QThread, QMutex, QMimeData, QFileInfo, QEvent, QPoint, QThreadPool, QRunnable
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QMouseEvent, QCursor, QDrag
import ctypes
import ctypes.wintypes

# 全局调试开关
_DEBUG_MODE = False  # 生产环境关闭，避免性能损耗
_EXPLORER_MONITOR_DEBUG = False  # Explorer Monitor 单独的日志开关

def debug_print(*args, **kwargs):
    """根据调试开关决定是否输出调试信息"""
    if _DEBUG_MODE:
        # 检查是否是 Explorer Monitor 日志
        if args and isinstance(args[0], str) and '[Explorer Monitor]' in args[0]:
            if _EXPLORER_MONITOR_DEBUG:
                print(*args, **kwargs)
        else:
            print(*args, **kwargs)


def dbg_exc(where=""):
    """在调试模式下记录当前正在处理的异常（含类型与消息），生产环境为零开销空操作。
    用于替代关键路径上原本静默吞掉异常的 `except ...: pass`，提升可诊断性而不影响行为。"""
    if not _DEBUG_MODE:
        return
    try:
        import sys as _sys
        exc = _sys.exc_info()[1]
        if exc is not None:
            debug_print(f"[swallowed]{(' ' + where) if where else ''}: {type(exc).__name__}: {exc}")
    except Exception:
        pass


def set_explorer_monitor_debug(enabled):
    """设置 Explorer Monitor 日志开关"""
    global _EXPLORER_MONITOR_DEBUG
    _EXPLORER_MONITOR_DEBUG = enabled
    debug_print(f"[Config] Explorer Monitor debug output: {'enabled' if enabled else 'disabled'}")


def get_process_memory_usage_mb():
    """Return current process working set in MB on Windows, else None."""
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
            _fields_ = [
                ('cb', wintypes.DWORD),
                ('PageFaultCount', wintypes.DWORD),
                ('PeakWorkingSetSize', ctypes.c_size_t),
                ('WorkingSetSize', ctypes.c_size_t),
                ('QuotaPeakPagedPoolUsage', ctypes.c_size_t),
                ('QuotaPagedPoolUsage', ctypes.c_size_t),
                ('QuotaPeakNonPagedPoolUsage', ctypes.c_size_t),
                ('QuotaNonPagedPoolUsage', ctypes.c_size_t),
                ('PagefileUsage', ctypes.c_size_t),
                ('PeakPagefileUsage', ctypes.c_size_t),
                ('PrivateUsage', ctypes.c_size_t),
            ]

        counters = PROCESS_MEMORY_COUNTERS_EX()
        counters.cb = ctypes.sizeof(counters)
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        psapi = ctypes.WinDLL('psapi', use_last_error=True)
        get_current_process = kernel32.GetCurrentProcess
        get_current_process.restype = wintypes.HANDLE
        get_process_memory_info = psapi.GetProcessMemoryInfo
        get_process_memory_info.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(PROCESS_MEMORY_COUNTERS_EX),
            wintypes.DWORD,
        ]
        get_process_memory_info.restype = wintypes.BOOL

        if get_process_memory_info(get_current_process(), ctypes.byref(counters), counters.cb):
            return round(float(counters.WorkingSetSize) / (1024 * 1024), 2)
    except Exception:
        return None
    return None


# CPU 占用率采样状态（进程 CPU 时间 / 挂钟时间 增量法，无 psutil 依赖）
_cpu_last_proc_time = None
_cpu_last_wall_time = None
_cpu_logical_count = os.cpu_count() or 1


def get_process_cpu_percent():
    """返回本进程自上次调用以来的平均 CPU 占用率（0~100，按逻辑核归一）。
    Windows 用 GetProcessTimes 取内核+用户时间增量除以挂钟增量；非 Windows 返回 None。"""
    global _cpu_last_proc_time, _cpu_last_wall_time
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD), ('dwHighDateTime', wintypes.DWORD)]

        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        kernel32.GetProcessTimes.argtypes = [wintypes.HANDLE, ctypes.POINTER(FILETIME),
                                             ctypes.POINTER(FILETIME), ctypes.POINTER(FILETIME),
                                             ctypes.POINTER(FILETIME)]
        kernel32.GetProcessTimes.restype = wintypes.BOOL
        creation = FILETIME(); exitt = FILETIME(); kern = FILETIME(); user = FILETIME()
        if not kernel32.GetProcessTimes(kernel32.GetCurrentProcess(),
                                        ctypes.byref(creation), ctypes.byref(exitt),
                                        ctypes.byref(kern), ctypes.byref(user)):
            return None
        proc_100ns = ((kern.dwHighDateTime << 32) | kern.dwLowDateTime) + \
                     ((user.dwHighDateTime << 32) | user.dwLowDateTime)
        now = time.monotonic()
        if _cpu_last_proc_time is None:
            _cpu_last_proc_time = proc_100ns
            _cpu_last_wall_time = now
            return None
        wall_delta = now - _cpu_last_wall_time
        proc_delta = (proc_100ns - _cpu_last_proc_time) / 1e7  # 100ns -> seconds
        _cpu_last_proc_time = proc_100ns
        _cpu_last_wall_time = now
        if wall_delta <= 0:
            return None
        pct = (proc_delta / wall_delta) * 100.0 / _cpu_logical_count
        return max(0.0, min(100.0, pct))
    except Exception:
        return None


# 整机 CPU 采样状态（GetSystemTimes：idle/kernel/user 时间增量，1 - idle/total = 占用率）
_sys_cpu_last_idle = None
_sys_cpu_last_total = None


def get_system_cpu_percent():
    """返回整机 CPU 占用率（0~100）。Windows 用 GetSystemTimes 计算 idle/total 增量；
    非 Windows 返回 None。需间隔调用两次取增量。"""
    global _sys_cpu_last_idle, _sys_cpu_last_total
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD), ('dwHighDateTime', wintypes.DWORD)]

        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        idle = FILETIME(); kern = FILETIME(); user = FILETIME()
        if not kernel32.GetSystemTimes(ctypes.byref(idle), ctypes.byref(kern), ctypes.byref(user)):
            return None
        idle_t = (idle.dwHighDateTime << 32) | idle.dwLowDateTime
        kern_t = (kern.dwHighDateTime << 32) | kern.dwLowDateTime  # 含 idle
        user_t = (user.dwHighDateTime << 32) | user.dwLowDateTime
        total = kern_t + user_t
        if _sys_cpu_last_total is None:
            _sys_cpu_last_idle = idle_t
            _sys_cpu_last_total = total
            return None
        idle_d = idle_t - _sys_cpu_last_idle
        total_d = total - _sys_cpu_last_total
        _sys_cpu_last_idle = idle_t
        _sys_cpu_last_total = total
        if total_d <= 0:
            return None
        return max(0.0, min(100.0, (1.0 - idle_d / total_d) * 100.0))
    except Exception:
        return None


def get_system_memory_status():
    """返回整机内存 (used_mb, total_mb, percent_used)；非 Windows 或失败返回 None。"""
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ('dwLength', wintypes.DWORD),
                ('dwMemoryLoad', wintypes.DWORD),
                ('ullTotalPhys', ctypes.c_ulonglong),
                ('ullAvailPhys', ctypes.c_ulonglong),
                ('ullTotalPageFile', ctypes.c_ulonglong),
                ('ullAvailPageFile', ctypes.c_ulonglong),
                ('ullTotalVirtual', ctypes.c_ulonglong),
                ('ullAvailVirtual', ctypes.c_ulonglong),
                ('ullAvailExtendedVirtual', ctypes.c_ulonglong),
            ]

        stat = MEMORYSTATUSEX()
        stat.dwLength = ctypes.sizeof(stat)
        if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat)):
            return None
        total_mb = stat.ullTotalPhys / (1024 * 1024)
        used_mb = (stat.ullTotalPhys - stat.ullAvailPhys) / (1024 * 1024)
        return (used_mb, total_mb, int(stat.dwMemoryLoad))
    except Exception:
        return None




# 启动外部进程时与当前进程解耦，避免主程序退出时连带关闭子进程
_DETACHED_FLAGS = 0
_NEW_PROCESS_GROUP = 0
_BREAKAWAY_FROM_JOB = 0
if os.name == 'nt':
    _DETACHED_FLAGS = getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
    _NEW_PROCESS_GROUP = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
    # break away 确保即使当前进程被放入 Job 对象，子进程也能独立存活
    _BREAKAWAY_FROM_JOB = 0x01000000


def launch_detached(cmd, cwd=None, extra_creationflags=0):
    """以与主进程解耦的方式启动外部程序，确保主进程退出后子进程仍然存活。

    在 VS Code / CI 等受限 Job Object 环境下，CREATE_BREAKAWAY_FROM_JOB 会被拒绝
    （ERROR_ACCESS_DENIED）。此时自动降级到 ShellExecuteW，Shell 进程本身在 Job
    之外创建新进程，完全独立于父进程的生命周期。
    """
    if os.name == 'nt':
        flags = _DETACHED_FLAGS | _NEW_PROCESS_GROUP | _BREAKAWAY_FROM_JOB | extra_creationflags
        try:
            return subprocess.Popen(cmd, cwd=cwd, creationflags=flags, close_fds=True)
        except OSError:
            # Job Object 不允许 breakaway —— 改用 ShellExecuteW（经由 Shell 创建，天然在 Job 外）
            exe = cmd[0] if isinstance(cmd, list) else cmd
            params = subprocess.list2cmdline(cmd[1:]) if isinstance(cmd, list) and len(cmd) > 1 else None
            try:
                import ctypes
                ctypes.windll.shell32.ShellExecuteW(
                    None, 'open', exe, params, cwd, 1  # 1 = SW_SHOWNORMAL
                )
                return None  # ShellExecuteW 不返回 Popen 对象
            except Exception:
                # 最终兜底：cmd /c start "" 也能从 Job Object 解脱
                start_cmd = ['cmd.exe', '/c', 'start', '""'] + (cmd if isinstance(cmd, list) else [cmd])
                return subprocess.Popen(start_cmd, cwd=cwd, close_fds=True)
    # 非 Windows 环境使用新 session，避免收到父进程信号
    return subprocess.Popen(cmd, cwd=cwd, start_new_session=True, close_fds=True)


def launch_detached_async(cmd, cwd=None, extra_creationflags=0):
    """在后台守护线程执行 launch_detached，避免 CreateProcess（及 Job Object 降级链）
    在 UI 线程阻塞。仅用于 fire-and-forget 场景（不使用返回的 Popen）。

    子进程一旦创建即与父进程解耦（DETACHED/BREAKAWAY），后台线程随即结束不影响其存活。
    调用方应先在 UI 线程完成校验（路径/可执行文件存在性）再调用本函数，以便错误提示仍能弹出。"""
    def _worker():
        try:
            launch_detached(cmd, cwd=cwd, extra_creationflags=extra_creationflags)
        except Exception as e:
            debug_print(f"[launch_detached_async] failed for {cmd!r}: {e}")
    try:
        threading.Thread(target=_worker, daemon=True).start()
    except Exception as e:
        # 线程创建失败极罕见：退回同步启动，保证功能可用
        debug_print(f"[launch_detached_async] thread start failed, fallback sync: {e}")
        try:
            launch_detached(cmd, cwd=cwd, extra_creationflags=extra_creationflags)
        except Exception as e2:
            debug_print(f"[launch_detached_async] sync fallback failed: {e2}")



def normalize_external_launch_dir(path):
    if not path:
        return None
    return os.path.normpath(path) if os.name == 'nt' else path


def find_git_install_root():
    git_root_candidates = [
        r"C:\Program Files\Git",
        r"C:\Program Files (x86)\Git",
        os.path.expandvars(r"%PROGRAMFILES%\Git"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Git"),
    ]
    return next((path for path in git_root_candidates if os.path.isdir(path)), None)


def launch_shell_tool(tool_name, cwd=None):
    cwd = normalize_external_launch_dir(cwd)

    if tool_name in ('cmd', 'powershell', 'git-bash'):
        if not cwd or not os.path.isdir(cwd):
            raise FileNotFoundError(tr("当前路径无效，无法启动终端"))

    # 使用 CREATE_NEW_CONSOLE | CREATE_NEW_PROCESS_GROUP | CREATE_BREAKAWAY_FROM_JOB
    # 确保启动的终端/程序在 TabEx 退出后仍然存活
    # 注意：CREATE_NEW_CONSOLE 和 DETACHED_PROCESS 互斥，终端需要 console 所以不用 DETACHED
    _DETACH_CONSOLE = (getattr(subprocess, 'CREATE_NEW_CONSOLE', 0x00000010)
                       | _NEW_PROCESS_GROUP | _BREAKAWAY_FROM_JOB)

    def _spawn_console_async(argv):
        """在后台守护线程创建带新控制台的进程，避免 CreateProcess 阻塞 UI 线程。
        路径/工具校验已在上方同步完成，故此处异常仅记录日志。"""
        def _worker():
            try:
                subprocess.Popen(argv, creationflags=_DETACH_CONSOLE, close_fds=True, cwd=cwd)
            except Exception as e:
                debug_print(f"[launch_shell_tool] async spawn failed for {argv!r}: {e}")
        try:
            threading.Thread(target=_worker, daemon=True).start()
        except Exception:
            _worker()  # 线程创建失败：退回同步

    if tool_name == 'cmd':
        _spawn_console_async(['cmd.exe', '/K', 'cd', '/d', cwd])
        return None

    if tool_name == 'powershell':
        escaped_dir = cwd.replace("'", "''")
        _spawn_console_async(['powershell.exe', '-NoExit', '-Command', f"Set-Location -LiteralPath '{escaped_dir}'"])
        return None

    if tool_name == 'git-bash':
        git_root = find_git_install_root()
        if not git_root:
            raise FileNotFoundError(tr("未找到 Git Bash，请确认已安装 Git for Windows"))
        git_bash_exe = os.path.join(git_root, 'git-bash.exe')
        if not os.path.exists(git_bash_exe):
            raise FileNotFoundError(tr("未找到可用的 Git Bash 可执行文件"))
        launch_detached_async([git_bash_exe, f'--cd={cwd}'], cwd=cwd)
        return None

    if tool_name == 'calculator':
        if os.name != 'nt':
            raise OSError(tr("当前系统不支持打开计算器"))
        launch_detached_async(['calc.exe'])
        return None

    raise ValueError(f"Unsupported tool: {tool_name}")


def is_supported_title_shortcut_path(path):
    return isinstance(path, str) and bool(path) and os.path.isfile(path) and path.lower().endswith(TITLE_SHORTCUT_EXTENSIONS)


def normalize_terminal_tool_name(tool_name, default='cmd'):
    if not isinstance(tool_name, str):
        return default
    value = tool_name.strip().lower()
    alias_map = {
        'cmd': 'cmd',
        'command prompt': 'cmd',
        'powershell': 'powershell',
        'ps': 'powershell',
        'pwsh': 'powershell',
        'git bash': 'git-bash',
        'git-bash': 'git-bash',
        'bash': 'git-bash',
    }
    normalized = alias_map.get(value, value)
    return normalized if normalized in SUPPORTED_TERMINAL_TOOLS else default


# 全局轻量提示气泡，用于替换阻塞式消息框
_active_toasts = []


class ToastMessage(QWidget):
    """右下角弹出的轻量提示，5s 自动消失"""

    def __init__(self, parent, title, message, level="info", duration=5000):
        super().__init__(parent)
        self.duration = duration
        self.level = level
        self.remaining_seconds = duration // 1000  # 剩余秒数
        self.setWindowFlags(
            Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.NoDropShadowWindowHint
        )
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        bg_map = {
            "info": "#2d8cf0",
            "warning": "#f0ad4e",
            "error": "#d9534f",
            "critical": "#d9534f",
            "success": "#5cb85c",
        }
        bg_color = bg_map.get(level, "#2d8cf0")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        # 标题行（标题 + 倒计时）
        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(8)
        
        title_label = QLabel(title)
        title_label.setStyleSheet("font-weight: bold; color: white;")
        title_layout.addWidget(title_label)
        
        title_layout.addStretch()
        
        self.countdown_label = QLabel(f"{self.remaining_seconds}s")
        self.countdown_label.setStyleSheet("color: rgba(255, 255, 255, 0.8); font-size: 11px;")
        title_layout.addWidget(self.countdown_label)
        
        layout.addLayout(title_layout)
        
        msg_label = QLabel(message)
        msg_label.setWordWrap(True)
        msg_label.setStyleSheet("color: white;")

        layout.addWidget(msg_label)

        self.setStyleSheet(
            f"background-color: {bg_color}; border-radius: 8px;"
        )

        # 倒计时定时器（每秒更新一次）
        self._countdown_timer = QTimer(self)
        self._countdown_timer.timeout.connect(self._update_countdown)
        self._countdown_timer.start(1000)  # 每1秒触发一次

        # 关闭定时器
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.close)
        self._timer.start(self.duration)

    def _update_countdown(self):
        """更新倒计时显示"""
        self.remaining_seconds -= 1
        if self.remaining_seconds > 0:
            self.countdown_label.setText(f"{self.remaining_seconds}s")
        else:
            self.countdown_label.setText("0s")
            self._countdown_timer.stop()

    def showEvent(self, event):
        super().showEvent(event)
        self.adjustSize()
        anchor = self.parent() if isinstance(self.parent(), QWidget) else None
        
        # 使用软件窗口的几何信息，而不是屏幕的几何信息
        if anchor and anchor.window():
            window_geo = anchor.window().geometry()
        else:
            # 如果没有父窗口，使用屏幕几何作为后备
            window_geo = QApplication.primaryScreen().availableGeometry()
        
        margin = 20
        existing = len(_active_toasts) - 1 if self in _active_toasts else len(_active_toasts)
        x = window_geo.right() - self.width() - margin
        y = window_geo.bottom() - self.height() - margin - existing * (self.height() + 10)
        self.move(x, y)

    def closeEvent(self, event):
        if self in _active_toasts:
            _active_toasts.remove(self)
        # 停止倒计时定时器
        if hasattr(self, '_countdown_timer'):
            self._countdown_timer.stop()
        super().closeEvent(event)


def show_toast(parent, title, message, level="info", duration=5000):
    """在右下角显示非阻塞提示"""
    anchor = parent.window() if isinstance(parent, QWidget) else None
    while len(_active_toasts) >= MAX_ACTIVE_TOASTS:
        old_toast = _active_toasts.pop(0)
        try:
            old_toast.close()
        except Exception:
            pass
    toast = ToastMessage(anchor, title, message, level=level, duration=duration)
    _active_toasts.append(toast)
    toast.show()


class GestureOverlay(QWidget):
    """鼠标手势可视化覆盖层（类似 Mouse Gestures）。

    - 全屏透明、鼠标穿透、始终置顶。
    - 拖拽时绘制跟随光标的轨迹线。
    - 实时显示当前识别到的方向箭头 + 动作名称浮窗。
    坐标统一使用全局屏幕坐标（与 WH_MOUSE_LL 钩子一致）。
    """

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(
            Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
            | Qt.WindowTransparentForInput | Qt.NoDropShadowWindowHint
        )
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self._points = []      # 全局坐标点列表 [QPoint, ...]
        self._arrow = ""       # 当前方向箭头
        self._label = ""       # 当前方向动作名

    def start(self):
        """开始一次手势：铺满虚拟桌面并清空轨迹。"""
        try:
            vg = QApplication.primaryScreen().virtualGeometry()
            self.setGeometry(vg)
        except Exception:
            self.setGeometry(QApplication.primaryScreen().geometry())
        self._points = []
        self._arrow = ""
        self._label = ""
        self.show()
        self.raise_()

    def add_point(self, gx, gy):
        from PyQt5.QtCore import QPoint
        self._points.append(QPoint(int(gx), int(gy)))
        self.update()

    def set_hint(self, arrow, label):
        if arrow != self._arrow or label != self._label:
            self._arrow = arrow
            self._label = label
            self.update()

    def finish(self):
        self.hide()
        self._points = []
        self._arrow = ""
        self._label = ""

    def paintEvent(self, event):
        if not self._points:
            return
        from PyQt5.QtGui import QPainter, QPen, QColor, QFont, QFontMetrics
        from PyQt5.QtCore import QRect
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        try:
            scale = max(1.0, self.logicalDpiX() / 96.0)
        except Exception:
            scale = 1.0
        origin = self.geometry().topLeft()

        # ── 轨迹线 ───────────────────────────────────────────────
        pen = QPen(QColor(0, 200, 120, 230))  # 半透明绿色
        pen.setWidth(max(3, int(4 * scale)))
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        prev = None
        for p in self._points:
            lp = p - origin
            if prev is not None:
                painter.drawLine(prev, lp)
            prev = lp

        # ── 实时方向提示浮窗 ─────────────────────────────────────
        if self._arrow:
            text = (f"{self._arrow}  {self._label}").strip()
            font = QFont()
            font.setPointSizeF(max(12.0, 15.0 * scale))
            font.setBold(True)
            painter.setFont(font)
            fm = QFontMetrics(font)
            tw = fm.boundingRect(text).width()
            th = fm.height()
            pad = int(12 * scale)
            last = self._points[-1] - origin
            bx = last.x() + int(24 * scale)
            by = last.y() + int(24 * scale)
            rect = QRect(bx, by, tw + pad * 2, th + pad * 2)
            if rect.right() > self.width():
                rect.moveRight(self.width() - int(8 * scale))
            if rect.bottom() > self.height():
                rect.moveBottom(self.height() - int(8 * scale))
            if rect.left() < 0:
                rect.moveLeft(int(8 * scale))
            if rect.top() < 0:
                rect.moveTop(int(8 * scale))
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(0, 0, 0, 180))
            painter.drawRoundedRect(rect, int(8 * scale), int(8 * scale))
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(rect, Qt.AlignCenter, text)


# ==================== 异步文件夹大小检查线程 ====================
class GitStatusWorker(QThread):
    """后台线程获取 Git 状态，避免阻塞 UI"""
    finished = pyqtSignal(str, str, object)  # dir_path, repo_root, summary_or_None

    def __init__(self, dir_path, repo_root, git_exe, parent=None):
        super().__init__(parent)
        self.dir_path = dir_path
        self.repo_root = repo_root
        self.git_exe = git_exe

    def run(self):
        try:
            result = subprocess.run(
                [self.git_exe, 'status', '--porcelain=v1', '-b', '--untracked-files=normal'],
                cwd=self.repo_root,
                capture_output=True, text=True, timeout=3,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000),
            )
            if result.returncode != 0:
                self.finished.emit(self.dir_path, self.repo_root, None)
                return
            lines = result.stdout.strip().splitlines()
            branch = ''
            staged = 0
            modified = 0
            untracked = 0
            for line in lines:
                if line.startswith('## '):
                    branch_info = line[3:]
                    branch = branch_info.split('...')[0].split()[0] if branch_info else ''
                    continue
                if len(line) < 2:
                    continue
                x, y = line[0], line[1]
                if x == '?' and y == '?':
                    untracked += 1
                else:
                    if x in ('M', 'A', 'D', 'R', 'C'):
                        staged += 1
                    if y in ('M', 'D'):
                        modified += 1
            is_clean = (staged == 0 and modified == 0 and untracked == 0)
            parts = []
            if branch:
                parts.append(tr("分支: {}").format(branch))
            if is_clean:
                parts.append(tr('<span style="color:#2e7d32;font-weight:bold">✔ 无更改</span>'))
            else:
                def _color(label, value, color_pos, color_zero='#2e7d32'):
                    color = color_pos if value > 0 else color_zero
                    return f'<span style="color:{color}">{label} {value}</span>'
                status_parts = [
                    _color(tr("暂存(Add)"), staged, "#d32f2f"),
                    _color(tr("修改(Commit)"), modified, "#d32f2f"),
                    _color(tr("未跟踪(待Add)"), untracked, "#f9a825"),
                ]
                parts.append('  '.join(status_parts))
            summary = ' | '.join(parts) if parts else None
            self.finished.emit(self.dir_path, self.repo_root, summary)
        except Exception:
            self.finished.emit(self.dir_path, self.repo_root, None)


class FolderSizeChecker(QThread):
    """后台线程检查文件夹大小，避免阻塞UI"""
    finished = pyqtSignal(str, int, bool)  # path, file_count, is_large
    
    def __init__(self, path, parent=None):
        super().__init__(parent)
        self.path = path
        self.should_stop = False
        self._mutex = QMutex()
    
    def run(self):
        """在后台线程中计算文件夹大小"""
        if not os.path.exists(self.path) or not os.path.isdir(self.path):
            self.finished.emit(self.path, 0, False)
            return
        
        try:
            file_count = 0
            start_time = time.time()
            
            # 快速计数，不递归子文件夹
            for entry in os.scandir(self.path):
                if self.should_stop:
                    debug_print(f"[FolderSizeChecker] Stopped checking {self.path}")
                    return
                
                file_count += 1
                
                # 超过阈值或超时则提前返回
                if file_count > LARGE_FOLDER_THRESHOLD:
                    debug_print(f"[FolderSizeChecker] Large folder detected: {self.path} (>{LARGE_FOLDER_THRESHOLD} files)")
                    self.finished.emit(self.path, file_count, True)
                    return
                
                # 超时检查
                if (time.time() - start_time) * 1000 > FOLDER_CHECK_TIMEOUT:
                    debug_print(f"[FolderSizeChecker] Timeout checking {self.path}")
                    self.finished.emit(self.path, file_count, file_count > LARGE_FOLDER_THRESHOLD)
                    return
            
            is_large = file_count > LARGE_FOLDER_THRESHOLD
            debug_print(f"[FolderSizeChecker] Folder {self.path}: {file_count} files (large={is_large})")
            self.finished.emit(self.path, file_count, is_large)
            
        except Exception as e:
            debug_print(f"[FolderSizeChecker] Error checking {self.path}: {e}")
            self.finished.emit(self.path, 0, False)
    
    def stop(self):
        """停止检查"""
        self._mutex.lock()
        self.should_stop = True
        self._mutex.unlock()


class _PidlResolver(QThread):
    """后台解析 shell 绝对 PIDL。

    SHParseDisplayName 对网络/UNC/映射盘是同步阻塞调用，若在 UI 线程执行会冻结
    整个 Qt 事件循环——表现为“一个慢标签把所有标签/窗口都卡住”。本线程在后台完成
    解析，把解析出的绝对 PIDL（进程内有效、非 COM 接口指针，可跨线程传递并用
    CoTaskMemFree 释放）通过信号回传 UI 线程，由 UI 线程调用 BrowseToIDList 导航。
    """
    resolved = pyqtSignal(str, object, int, int)  # path, pidl(int|None), hr, generation

    def __init__(self, path, generation, parent=None):
        super().__init__(parent)
        self._path = path
        self._generation = generation

    def run(self):
        import ctypes
        pidl_val = None
        hr = -1
        co_init = False
        try:
            try:
                # 后台线程需自备 COM 环境；COINIT_APARTMENTTHREADED = 0x2
                ctypes.windll.ole32.CoInitializeEx(None, 0x2)
                co_init = True
            except Exception:
                pass
            _spdn = ctypes.windll.shell32.SHParseDisplayName
            _spdn.restype = ctypes.c_long
            _spdn.argtypes = [
                ctypes.c_wchar_p, ctypes.c_void_p,
                ctypes.POINTER(ctypes.c_void_p),
                ctypes.c_ulong, ctypes.POINTER(ctypes.c_ulong),
            ]
            pidl = ctypes.c_void_p(0)
            sfgao = ctypes.c_ulong(0)
            hr = _spdn(self._path, None, ctypes.byref(pidl), 0, ctypes.byref(sfgao))
            if hr == 0 and pidl.value:
                pidl_val = pidl.value  # 绝对 PIDL：跨线程有效
        except Exception as e:
            debug_print(f"[PidlResolver] error for '{self._path}': {e}")
            hr = -1
        finally:
            if co_init:
                try:
                    ctypes.windll.ole32.CoUninitialize()
                except Exception:
                    pass
        self.resolved.emit(self._path, pidl_val, int(hr), self._generation)


def set_debug_mode(enabled):
    """设置全局调试模式"""
    global _DEBUG_MODE
    _DEBUG_MODE = enabled

def qt_message_handler(mode, context, message):
    """自定义 Qt 消息处理器，过滤 QAxBase 等不需要的警告"""
    # 只在调试模式下输出 Qt 警告
    if _DEBUG_MODE:
        # 如果是调试模式，输出所有消息
        debug_print(f"Qt Message: {message}")
    else:
        # 非调试模式下，只输出严重错误（Critical 和 Fatal）
        from PyQt5.QtCore import QtCriticalMsg, QtFatalMsg
        if mode in (QtCriticalMsg, QtFatalMsg):
            debug_print(f"Qt Error: {message}")
        # 其他消息（Debug, Warning, Info）都被过滤

# Optional native hit-test support (Windows)
try:
    import ctypes
    import win32gui
    import win32con
    HAS_PYWIN = True
except Exception:
    HAS_PYWIN = False

# Windows API for monitoring new Explorer windows
if HAS_PYWIN:
    try:
        import ctypes.wintypes as wintypes
        user32 = ctypes.windll.user32
        ole32 = ctypes.windll.ole32
        
        # Constants for SetWinEventHook
        EVENT_OBJECT_CREATE = 0x8000
        EVENT_SYSTEM_FOREGROUND = 0x0003
        WINEVENT_OUTOFCONTEXT = 0x0000
        
        # Define callback type
        WinEventProcType = ctypes.WINFUNCTYPE(
            None,
            wintypes.HANDLE,
            wintypes.DWORD,
            wintypes.HWND,
            wintypes.LONG,
            wintypes.LONG,
            wintypes.DWORD,
            wintypes.DWORD
        )
    except Exception as e:
        debug_print(f"Failed to setup Windows API monitoring: {e}")


class SimplePathBar(QWidget):
    """Windows Explorer 风格面包屑地址栏（单控件自绘实现）。
    - 普通模式：paintEvent 一次性绘制可点击的路径分段 (C: › project › EOL)，
      通过命中测试判断点击了哪一段；不再为每一段创建/销毁子控件。
    - 编辑模式：覆盖显示常驻 QLineEdit，按 Enter 导航，按 Escape 取消。
    - 点击路径段之间的空白区域或调用 enter_edit_mode() 进入编辑模式。
    接口与旧实现兼容：set_path / pathChanged /
    enter_edit_mode / exit_edit_mode / get_path_for_copy。

    重要：旧实现每次导航都销毁并重建 N 个 QPushButton/QToolButton，叠加多轮
    强制重排与自愈循环，导致最小化恢复后偶发“地址栏不刷新/卡顿”（数据正确、
    像素滞后）。单控件自绘从根本上消除该问题——一个 paintEvent 一定会被 Qt
    正确绘制，且没有子控件析构/创建开销。
    """
    pathChanged = pyqtSignal(str)

    _FONT   = "font-family: 'Segoe UI', 'Microsoft YaHei UI', sans-serif; font-size: 11pt;"
    _S_BAR  = ("SimplePathBar { background: #ffffff; border: none; }")
    _S_EDIT = ("QLineEdit { background: white; border: 1px solid #ccc; border-radius: 3px;"
               " font-family: 'Segoe UI', 'Microsoft YaHei UI', sans-serif; font-size: 10pt;"
               " color: #202020; padding: 2px 6px; selection-background-color: #0078d4; }")

    # 绘制配色（与旧样式表一致）
    _COL_SEG   = '#003d7a'   # 路径段文字
    _COL_SEP   = '#888888'   # 分隔符 ›
    _COL_HOVER = '#cce5ff'   # 悬停背景
    _CHEVRON   = '\u203a'    # ›
    _PAD       = 6           # 段内文字左右内边距
    _SEP_W     = 16          # 分隔符宽度

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_path = ''
        self._segments = []            # [(显示名, 完整路径), …] 全部层级
        self._display_regions = []     # 绘制时记录的命中区: dict(kind,x0,x1,payload,label,is_current)
        self._hover_idx = -1
        self._in_edit = False
        self._completer = None         # 编辑模式路径自动补全（首次进入编辑时惰性创建）
        self._press_pos = None         # 左键按下位置（用于拖拽判定）
        self._press_idx = -1           # 按下时命中的区索引
        self._dragging = False
        self.setFixedHeight(30)
        self.setMouseTracking(True)    # hover 高亮需要
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(self._S_BAR)

        from PyQt5.QtGui import QFont
        self._font = QFont("Segoe UI", 11)
        try:
            self._font.setWeight(QFont.Medium)  # 500
        except Exception:
            pass
        # 让 self.fontMetrics()（_choose_display_parts / _elide_crumb_text 使用）
        # 与 paintEvent 绘制字体一致，避免宽度测算与实际绘制不符导致的折叠/裁剪偏差。
        self.setFont(self._font)
        # 编辑框：常驻子控件，编辑模式下覆盖显示，不参与常态绘制
        self._edit = QLineEdit(self)
        self._edit.setStyleSheet(self._S_EDIT)
        self._edit.setFrame(False)
        self._edit.returnPressed.connect(self._commit_edit)
        self._edit.installEventFilter(self)
        self._edit.hide()

    # ── 公共 API ─────────────────────────────────────────────────────────────

    def set_path(self, path):
        new_path = path or ''
        path_changed = (new_path != self._current_path)
        if _DEBUG_MODE:
            debug_print(f"[SimplePathBar] set_path: new='{new_path}', old='{self._current_path}', changed={path_changed}, in_edit={self._in_edit}")
        self._current_path = new_path
        if self._in_edit:
            if path_changed:
                # 路径已变化（用户通过 Explorer 导航了）→ 退出编辑模式并显示新路径
                self.exit_edit_mode()
            # 路径未变化时保持编辑模式不中断（用户可能在复制路径）
            return
        # 单控件自绘：只更新分段数据并请求重绘，绝不创建/销毁子控件。
        if path_changed:
            # 空路径时保留现有分段，避免瞬时空白。
            if new_path:
                self._segments = self._split_path(new_path)
            self._hover_idx = -1
            self.update()
        else:
            # 路径未变化（来自轮询/保活的重复回写）：轻量异步刷新即可，无需强制同步。
            self.update()

    def force_refresh(self):
        """同步重绘：供窗口恢复/手动兜底调用，强制像素立即跟上当前路径。
        单控件 repaint() 一定被 Qt 立即执行，规避恢复后异步 update() 被合成器丢弃。"""
        if self._in_edit:
            return
        if self._current_path:
            self._segments = self._split_path(self._current_path)
        self.repaint()

    def enter_edit_mode(self):
        if self._in_edit:
            return
        debug_print(f"[SimplePathBar] enter_edit_mode: path='{self._current_path}'")
        self._ensure_completer()
        self._in_edit = True
        self._edit.setGeometry(self.rect())
        self._edit.setText(self._current_path)
        self._edit.show()
        self._edit.raise_()
        self._edit.setFocus()
        self._edit.selectAll()
        # 安装应用级事件过滤器，监听点击外部区域时退出编辑模式
        QApplication.instance().installEventFilter(self)
        self.update()

    def exit_edit_mode(self):
        if not self._in_edit:
            return  # 已经不在编辑模式
        self._in_edit = False
        self._edit.hide()
        # 移除应用级事件过滤器
        try:
            QApplication.instance().removeEventFilter(self)
        except Exception:
            pass
        self.update()

    def changeEvent(self, event):
        """窗口激活时请求重绘（单控件自绘，无子控件丢失问题）。"""
        if event.type() == QEvent.WindowActivate and not self._in_edit:
            self.update()
        super().changeEvent(event)

    # ── 自绘 + 命中测试 ────────────────────────────────────────────────────────

    def _region_at(self, x):
        """返回横坐标 x 命中的显示区索引；未命中返回 -1。"""
        for i, r in enumerate(self._display_regions):
            if r['x0'] <= x < r['x1']:
                return i
        return -1

    def paintEvent(self, event):
        if self._in_edit:
            return
        from PyQt5.QtGui import QPainter, QColor, QFont
        from PyQt5.QtCore import QRectF
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor('#ffffff'))
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        painter.setFont(self._font)
        fm = painter.fontMetrics()
        h = self.height()
        y = int((h + fm.ascent() - fm.descent()) / 2)
        regions = []
        parts = self._segments
        if not parts:
            self._display_regions = regions
            painter.end()
            return
        display_parts = self._choose_display_parts(parts)
        left_collapsed = len(display_parts) < len(parts)
        col_seg = QColor(self._COL_SEG)
        col_sep = QColor(self._COL_SEP)
        col_hover = QColor(self._COL_HOVER)
        pad = self._PAD
        sep_w = self._SEP_W
        chevron = self._CHEVRON
        x = 2

        def hover_bg(x0, x1):
            painter.save()
            painter.setPen(Qt.NoPen)
            painter.setBrush(col_hover)
            painter.drawRoundedRect(QRectF(float(x0), 4.0, float(x1 - x0), float(h - 8)), 3.0, 3.0)
            painter.restore()

        idx = 0
        if left_collapsed:
            # 折叠省略号（点击进入编辑模式，与旧行为一致）
            ell = '...'
            w = fm.horizontalAdvance(ell) + pad * 2
            if self._hover_idx == idx:
                hover_bg(x, x + w)
            painter.setPen(col_sep)
            painter.drawText(x + pad, y, ell)
            regions.append({'kind': 'ellipsis', 'x0': x, 'x1': x + w,
                            'payload': None, 'label': self._current_path, 'is_current': False})
            x += w
            idx += 1
            # 折叠后的分隔符：列出首个显示段的同级文件夹
            first_parent = os.path.dirname(display_parts[0][1]) if display_parts else None
            if self._hover_idx == idx and first_parent:
                hover_bg(x, x + sep_w)
            painter.setPen(col_sep)
            painter.drawText(x + (sep_w - fm.horizontalAdvance(chevron)) // 2, y, chevron)
            regions.append({'kind': 'separator', 'x0': x, 'x1': x + sep_w,
                            'payload': first_parent, 'label': None, 'is_current': False})
            x += sep_w
            idx += 1

        for i, (label, full) in enumerate(display_parts):
            if i > 0:
                # 分隔符下拉列出左侧段的子文件夹（即当前段的同级）
                parent = display_parts[i - 1][1]
                if self._hover_idx == idx and parent:
                    hover_bg(x, x + sep_w)
                painter.setPen(col_sep)
                painter.drawText(x + (sep_w - fm.horizontalAdvance(chevron)) // 2, y, chevron)
                regions.append({'kind': 'separator', 'x0': x, 'x1': x + sep_w,
                                'payload': parent, 'label': None, 'is_current': False})
                x += sep_w
                idx += 1
            is_current = (full == self._current_path)
            shown = self._elide_crumb_text(label, is_current=is_current)
            w = fm.horizontalAdvance(shown) + pad * 2
            hovered = (self._hover_idx == idx)
            if hovered:
                hover_bg(x, x + w)
            f = QFont(self._font)
            f.setUnderline(hovered)
            painter.setFont(f)
            painter.setPen(col_seg)
            painter.drawText(x + pad, y, shown)
            painter.setFont(self._font)
            regions.append({'kind': 'segment', 'x0': x, 'x1': x + w,
                            'payload': full, 'label': label, 'is_current': is_current})
            x += w
            idx += 1

        self._display_regions = regions
        painter.end()

    def mouseMoveEvent(self, event):
        if self._in_edit:
            return super().mouseMoveEvent(event)
        # 拖拽判定：在某个路径段上按下并拖动 → 拖到标签栏打开新 tab
        if self._press_pos is not None and (event.buttons() & Qt.LeftButton):
            if (0 <= self._press_idx < len(self._display_regions) and
                    (event.pos() - self._press_pos).manhattanLength() >= QApplication.startDragDistance()):
                r = self._display_regions[self._press_idx]
                if r['kind'] == 'segment' and r['payload']:
                    self._dragging = True
                    drag = QDrag(self)
                    mime = QMimeData()
                    mime.setUrls([QUrl.fromLocalFile(r['payload'])])
                    mime.setText(r['payload'])
                    drag.setMimeData(mime)
                    self._press_pos = None
                    self._press_idx = -1
                    drag.exec_(Qt.CopyAction | Qt.MoveAction)
                    return
        idx = self._region_at(int(event.pos().x()))
        if idx != self._hover_idx:
            self._hover_idx = idx
            if idx >= 0:
                r = self._display_regions[idx]
                clickable = (r['kind'] == 'segment' or
                             (r['kind'] == 'separator' and r['payload']) or
                             r['kind'] == 'ellipsis')
                self.setCursor(Qt.PointingHandCursor if clickable else Qt.IBeamCursor)
                self.setToolTip(r.get('label') or '')
            else:
                self.setCursor(Qt.IBeamCursor)   # 空白处点击可进入编辑模式
                self.setToolTip('')
            self.update()
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and not self._in_edit:
            self._press_pos = event.pos()
            self._press_idx = self._region_at(int(event.pos().x()))
            self._dragging = False
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self._in_edit or event.button() != Qt.LeftButton:
            return super().mouseReleaseEvent(event)
        was_dragging = self._dragging
        press_idx = self._press_idx
        self._press_pos = None
        self._press_idx = -1
        self._dragging = False
        if was_dragging:
            return super().mouseReleaseEvent(event)
        idx = self._region_at(int(event.pos().x()))
        if idx < 0:
            # 空白区域 → 进入编辑模式（与旧行为一致）
            if press_idx < 0:
                self.enter_edit_mode()
            return super().mouseReleaseEvent(event)
        if idx != press_idx:
            # 按下与释放不在同一区，视为取消
            return super().mouseReleaseEvent(event)
        r = self._display_regions[idx]
        kind = r['kind']
        if kind == 'segment':
            if r.get('is_current'):
                self.enter_edit_mode()        # 点击当前(末级)目录 → 编辑
            elif r['payload']:
                self.pathChanged.emit(r['payload'])
        elif kind == 'separator':
            parent = r['payload']
            if parent and os.path.isdir(parent):
                self._show_sibling_menu(parent)
        elif kind == 'ellipsis':
            self.enter_edit_mode()
        super().mouseReleaseEvent(event)

    def leaveEvent(self, event):
        if self._hover_idx != -1:
            self._hover_idx = -1
            self.setToolTip('')
            self.update()
        super().leaveEvent(event)

    def get_path_for_copy(self, separator='\\'):
        p = self._current_path
        if separator != '\\':
            p = p.replace('\\', separator)
        return p

    def _ensure_completer(self):
        """惰性创建路径自动补全器（仅补全目录，类似 Explorer 地址栏）。
        QCompleter 对 QFileSystemModel 有内置的路径拆分支持，能逐级补全 D:\\a\\b。"""
        if self._completer is not None:
            return  # 已创建或已尝试失败
        try:
            from PyQt5.QtWidgets import QCompleter, QFileSystemModel
            model = QFileSystemModel(self)
            model.setRootPath('')
            model.setFilter(QDir.Dirs | QDir.NoDotAndDotDot | QDir.Drives)
            comp = QCompleter(model, self)
            comp.setCaseSensitivity(Qt.CaseInsensitive)
            comp.setCompletionMode(QCompleter.PopupCompletion)
            self._edit.setCompleter(comp)
            self._completer = comp
        except Exception as e:
            debug_print(f"[SimplePathBar] completer init failed: {e}")
            self._completer = False  # 标记已尝试，不重试

    def _show_sibling_menu(self, parent_path):
        """弹出 parent_path 下的子文件夹菜单，选择后导航到该文件夹。"""
        try:
            from PyQt5.QtGui import QCursor
            entries = []
            try:
                with os.scandir(parent_path) as it:
                    for e in it:
                        try:
                            if e.is_dir():
                                entries.append(e.name)
                        except Exception:
                            pass
            except Exception as e:
                debug_print(f"[SimplePathBar] sibling scandir error: {e}")
                return
            entries.sort(key=lambda s: s.lower())
            menu = QMenu(self)
            if not entries:
                act = menu.addAction(tr("无子文件夹"))
                act.setEnabled(False)
            else:
                cur_norm = os.path.normcase(os.path.normpath(self._current_path or ''))
                for name in entries[:300]:
                    full = os.path.join(parent_path, name)
                    act = menu.addAction(name)
                    try:
                        full_norm = os.path.normcase(os.path.normpath(full))
                        if cur_norm == full_norm or cur_norm.startswith(full_norm + os.sep):
                            f = act.font(); f.setBold(True); act.setFont(f)
                    except Exception:
                        pass
                    act.triggered.connect(lambda _=False, p=full: self.pathChanged.emit(p))
            menu.exec_(QCursor.pos())
        except Exception as e:
            debug_print(f"[SimplePathBar] sibling menu error: {e}")

    def _elide_crumb_text(self, text, is_current=False):
        try:
            fm = self.fontMetrics()
            # 当前目录尽量完整保留；中间层更积极省略。
            limit = 240 if is_current else 140
            return fm.elidedText(str(text), Qt.ElideMiddle, limit)
        except Exception:
            return text

    def _choose_display_parts(self, parts):
        """根据可用宽度选择要显示的路径尾部，优先显示当前目录。"""
        try:
            if not parts:
                return []
            if len(parts) <= 2:
                return parts

            fm = self.fontMetrics()
            avail = max(120, int(self.width() or 0) - 12)
            sep_w = 18
            ellipsis_w = max(24, fm.horizontalAdvance('...') + 8)

            def part_width(label, is_current=False):
                shown = self._elide_crumb_text(label, is_current=is_current)
                return max(24, fm.horizontalAdvance(shown) + 16)

            def total_parts_width(items):
                total = 0
                for idx, (label, _full_path) in enumerate(items):
                    total += part_width(label, is_current=(idx == len(items) - 1))
                    if idx > 0:
                        total += sep_w
                return total

            # 先判断完整路径是否本来就放得下。之前的增量算法在尝试从右向左保留时，
            # 会过早为左侧 "..." 预留宽度，导致“其实足够宽却被折叠”。
            if total_parts_width(parts) <= avail:
                return parts

            # 从最后一级开始向左保留，确保当前目录优先可见。
            selected_rev = []
            used = 0
            for rev_idx, (label, full_path) in enumerate(reversed(parts)):
                is_current = (rev_idx == 0)
                w = part_width(label, is_current=is_current)
                extra = w if not selected_rev else (sep_w + w)
                if selected_rev and used + extra + ellipsis_w > avail:
                    break
                if not selected_rev and w > avail:
                    selected_rev.append((label, full_path))
                    used = min(w, avail)
                    break
                selected_rev.append((label, full_path))
                used += extra

            selected = list(reversed(selected_rev))
            if len(selected) == len(parts):
                return parts

            # 若仍有空间，尽量把根目录也保留在 ... 后面，提升定位感。
            root = parts[0]
            if selected and selected[0][1] != root[1]:
                root_w = part_width(root[0], is_current=False)
                needed = used + ellipsis_w + sep_w + root_w + sep_w
                if needed <= avail:
                    return [root] + selected

            return selected
        except Exception:
            return parts

    def _split_path(self, path):
        """将路径拆分为 [(显示名, 完整路径), …]。"""
        if not path:
            return []
        if path.startswith('shell:') or '::' in path:
            return [(path, path)]
        norm = os.path.normpath(path)
        # splitdrive 可正确识别本地盘符（'C:'）与 UNC 共享根（'\\\\server\\share'），
        # 直接按 os.sep 拆分会丢失 UNC 的 '\\\\' 前缀，导致面包屑各段路径失效。
        drive, tail = os.path.splitdrive(norm)
        result = []
        if drive:
            root = drive + os.sep
            result.append((drive, root))
            current = root
            parts = [p for p in tail.split(os.sep) if p]
        else:
            parts = [p for p in norm.split(os.sep) if p]
            if not parts:
                return []
            current = parts[0] + os.sep
            result.append((parts[0], current))
            parts = parts[1:]
        for part in parts:
            current = os.path.join(current, part)
            result.append((part, current))
        return result

    def _commit_edit(self):
        path = self._edit.text().strip()
        self.exit_edit_mode()
        if path:
            self.pathChanged.emit(path)

    def resizeEvent(self, event):
        # 单控件自绘：尺寸变化时重定位编辑框覆盖层并请求重绘（无子控件重建）
        if self._in_edit:
            self._edit.setGeometry(self.rect())
        self.update()
        super().resizeEvent(event)

    def eventFilter(self, obj, event):
        if obj is self._edit and event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Escape:
                debug_print(f"[SimplePathBar] eventFilter: Escape pressed, exit edit mode")
                self.exit_edit_mode()
                return True
        # 应用级事件过滤：点击编辑框外部时退出编辑模式
        if self._in_edit and event.type() == QEvent.MouseButtonPress:
            try:
                # 自动补全下拉框可见时，点击下拉项不应退出编辑模式
                if self._completer and self._completer.popup() and self._completer.popup().isVisible():
                    return super().eventFilter(obj, event)
                global_pos = event.globalPos()
                edit_rect = self._edit.rect()
                local_pos = self._edit.mapFromGlobal(global_pos)
                if not edit_rect.contains(local_pos):
                    # 额外检查：点击是否在整个 SimplePathBar 区域内（如果在路径栏区域内不退出）
                    bar_local = self.mapFromGlobal(global_pos)
                    if self.rect().contains(bar_local):
                        debug_print(f"[SimplePathBar] eventFilter: click inside path bar area but outside edit, ignoring exit")
                    else:
                        debug_print(f"[SimplePathBar] eventFilter: click outside path bar, exit edit mode. global_pos={global_pos.x()},{global_pos.y()}")
                        self.exit_edit_mode()
            except Exception:
                pass
        return super().eventFilter(obj, event)


class ClickableLabel(QLabel):
    """可点击的标签，用于面包屑导航，支持拖拽到标签栏"""
    clicked = pyqtSignal(str)
    
    def __init__(self, text, path, parent=None):
        super().__init__(text, parent)
        self.path = path
        # 保存完整文本以便在缩小时进行省略显示
        self.full_text = text
        self.drag_start_position = None
        self.is_dragging = False
        # 设置最小宽度，允许横向压缩但保持可见
        try:
            from PyQt5.QtWidgets import QApplication
            dpi = QApplication.primaryScreen().logicalDotsPerInch()
            scale = dpi / 96.0
        except Exception:
            scale = 1.0
        min_label_w = int(40 * scale)
        self.setMinimumWidth(min_label_w)
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self.setStyleSheet("""
            QLabel {
                color: #003d7a;
                font-family: 'Segoe UI', 'Microsoft YaHei UI', sans-serif;
                font-size: 11pt;
                font-weight: 500;
                padding: 1px;
                margin: 0 2px;
                border-radius: 2px;
            }
            QLabel:hover {
                background-color: #cce5ff;
                text-decoration: underline;
            }
        """)
        self.setCursor(Qt.PointingHandCursor)

    def resizeEvent(self, event):
        """根据当前宽度使用中间省略显示文本"""
        try:
            from PyQt5.QtGui import QFontMetrics
            fm = QFontMetrics(self.font())
            # 使用中间省略，保留路径前后信息
            elided = fm.elidedText(self.full_text, Qt.ElideMiddle, max(10, self.width()))
            super().setText(elided)
        except Exception:
            pass
        super().resizeEvent(event)
    
    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            # 记录拖拽起始位置
            self.drag_start_position = event.pos()
            self.is_dragging = False
        super().mousePressEvent(event)
    
    def mouseMoveEvent(self, event: QMouseEvent):
        # 检查是否应该开始拖拽
        if not (event.buttons() & Qt.LeftButton):
            return
        if self.drag_start_position is None:
            return
        
        # 计算移动距离
        distance = (event.pos() - self.drag_start_position).manhattanLength()
        
        # 如果移动距离超过阈值，开始拖拽
        if distance >= QApplication.startDragDistance():
            self.is_dragging = True
            self.start_drag()
    
    def mouseReleaseEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            # 只有在没有拖拽的情况下才发出点击信号
            if not self.is_dragging and self.drag_start_position is not None:
                self.clicked.emit(self.path)
            # 重置状态
            self.drag_start_position = None
            self.is_dragging = False
        super().mouseReleaseEvent(event)
    
    def start_drag(self):
        """开始拖拽操作"""
        drag = QDrag(self)
        mime_data = QMimeData()
        
        # 设置拖拽的数据（文件夹路径）
        from PyQt5.QtCore import QUrl
        url = QUrl.fromLocalFile(self.path)
        mime_data.setUrls([url])
        mime_data.setText(self.path)
        
        drag.setMimeData(mime_data)
        
        # 执行拖拽
        debug_print(f"[Breadcrumb Drag] Starting drag for path: {self.path}")
        drag.exec_(Qt.CopyAction | Qt.MoveAction)


class BookmarkManager:
    def __init__(self, config_file="bookmarks.json"):
        if os.path.isabs(config_file):
            self.config_file = config_file
        else:
            self.config_file = get_app_data_path(config_file)
        self.bookmark_tree = self.load_bookmarks()
        # 优化：延迟保存机制，避免频繁写入磁盘
        self._save_timer = None
        self._pending_save = False

    def load_bookmarks(self):
        # 只加载主书签文件，不做备份和恢复
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if 'roots' in data:
                    return data['roots']
                return data
            except Exception as e:
                debug_print(f"Failed to load bookmarks: {e}")
                return {}
        else:
            debug_print("No bookmark file found, starting with empty bookmarks")
            return {}

    def save_bookmarks(self, immediate=False):
        # 优化：延迟保存，避免频繁操作时多次写入
        if immediate:
            tmp_path = self.config_file + ".tmp"
            try:
                with open(tmp_path, 'w', encoding='utf-8') as f:
                    json.dump({"roots": self.bookmark_tree}, f, ensure_ascii=False, indent=2)
                os.replace(tmp_path, self.config_file)  # 原子替换，防止断电损坏
                self._pending_save = False
            except Exception as e:
                debug_print(f"Failed to save bookmarks: {e}")
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
        else:
            self._pending_save = True
            if self._save_timer is None:
                from PyQt5.QtCore import QTimer
                self._save_timer = QTimer()
                self._save_timer.setSingleShot(True)
                self._save_timer.timeout.connect(lambda: self.save_bookmarks(immediate=True))
            self._save_timer.start(500)

    def get_all_bookmarks(self):
        # 返回所有书签（递归）
        bookmarks = []
        def collect(node):
            if isinstance(node, dict):
                if node.get('type') == 'url':
                    bookmarks.append(node)
                elif node.get('type') == 'folder' and 'children' in node:
                    for child in node['children']:
                        collect(child)
            elif isinstance(node, list):
                for item in node:
                    collect(item)
        for root in self.bookmark_tree.values():
            collect(root)
        return bookmarks

    def get_tree(self):
        # 返回完整树结构
        return self.bookmark_tree

    def add_bookmark(self, parent_folder_id, name, url):
        # 在指定文件夹下添加书签
        def find_folder(node, folder_id):
            if isinstance(node, dict):
                if node.get('type') == 'folder' and node.get('id') == folder_id:
                    return node
                if 'children' in node:
                    for child in node['children']:
                        found = find_folder(child, folder_id)
                        if found:
                            return found
            elif isinstance(node, list):
                for item in node:
                    found = find_folder(item, folder_id)
                    if found:
                        return found
            return None
        folder = None
        for root in self.bookmark_tree.values():
            folder = find_folder(root, parent_folder_id)
            if folder:
                break
        if folder is not None:
            # 生成唯一id
            import time
            new_id = str(int(time.time() * 1000000))
            bookmark = {
                "date_added": new_id,
                "id": new_id,
                "name": name,
                "type": "url",
                "url": url
            }
            folder.setdefault('children', []).append(bookmark)
            self.save_bookmarks()
            return True
        return False


# ─────────────────────────────────────────────────────────────────────────────
# IExplorerBrowser-based file view
# Hosts the real Windows Explorer shell component (IExplorerBrowser COM) which
# supports TortoiseGit overlay icons – unlike Shell.Explorer (IE/WebBrowser
# ActiveX) which does not load shell icon-overlay extensions.
# ─────────────────────────────────────────────────────────────────────────────
_COMTYPES_AVAILABLE = False
try:
    import comtypes
    import comtypes.client
    from comtypes import GUID, HRESULT, IUnknown, COMMETHOD
    _COMTYPES_AVAILABLE = True
except ImportError:
    pass

if _COMTYPES_AVAILABLE:
    class _FOLDERSETTINGS(ctypes.Structure):
        _fields_ = [("ViewMode", ctypes.c_uint), ("fFlags", ctypes.c_uint)]

    class _IExplorerBrowserEvents(IUnknown):
        _iid_ = GUID("{361BBDC7-E6EE-4E13-BE58-58E2240C810F}")
        _methods_ = [
            COMMETHOD([], HRESULT, 'OnNavigationPending',
                      (['in'], ctypes.c_void_p, 'pidlFolder')),
            COMMETHOD([], HRESULT, 'OnViewCreated',
                      (['in'], ctypes.POINTER(IUnknown), 'psv')),
            COMMETHOD([], HRESULT, 'OnNavigationComplete',
                      (['in'], ctypes.c_void_p, 'pidlFolder')),
            COMMETHOD([], HRESULT, 'OnNavigationFailed',
                      (['in'], ctypes.c_void_p, 'pidlFolder')),
        ]

    class _IExplorerBrowser(IUnknown):
        _iid_ = GUID("{DFD3B6B5-C10C-4BE9-85F6-A66969F402F6}")
        _methods_ = [
            COMMETHOD([], HRESULT, 'Initialize',
                      (['in'], ctypes.wintypes.HWND, 'hwndParent'),
                      (['in'], ctypes.POINTER(ctypes.wintypes.RECT), 'prc'),
                      (['in'], ctypes.POINTER(_FOLDERSETTINGS), 'pfs')),
            COMMETHOD([], HRESULT, 'Destroy'),
            COMMETHOD([], HRESULT, 'SetRect',
                      (['in'], ctypes.c_void_p, 'phdwp'),
                      (['in'], ctypes.wintypes.RECT, 'rcBrowser')),
            COMMETHOD([], HRESULT, 'SetPropertyBag',
                      (['in'], ctypes.c_wchar_p, 'pszPropertyBag')),
            COMMETHOD([], HRESULT, 'SetEmptyText',
                      (['in'], ctypes.c_wchar_p, 'pszEmptyText')),
            COMMETHOD([], HRESULT, 'SetFolderSettings',
                      (['in'], ctypes.POINTER(_FOLDERSETTINGS), 'pfs')),
            COMMETHOD([], HRESULT, 'Advise',
                      (['in'], ctypes.POINTER(_IExplorerBrowserEvents), 'psbe'),
                      (['out'], ctypes.POINTER(ctypes.c_ulong), 'pdwCookie')),
            COMMETHOD([], HRESULT, 'Unadvise',
                      (['in'], ctypes.c_ulong, 'dwCookie')),
            COMMETHOD([], HRESULT, 'SetOptions',
                      (['in'], ctypes.c_uint, 'dwFlag')),
            COMMETHOD([], HRESULT, 'GetOptions',
                      (['out'], ctypes.POINTER(ctypes.c_uint), 'pdwFlag')),
            COMMETHOD([], HRESULT, 'BrowseToIDList',
                      (['in'], ctypes.c_void_p, 'pidl'),
                      (['in'], ctypes.c_uint, 'uFlags')),
            COMMETHOD([], HRESULT, 'BrowseToObject',
                      (['in'], ctypes.POINTER(IUnknown), 'punk'),
                      (['in'], ctypes.c_uint, 'uFlags')),
            COMMETHOD([], HRESULT, 'FillFromObject',
                      (['in'], ctypes.POINTER(IUnknown), 'punk'),
                      (['in'], ctypes.c_uint, 'dwFlags')),
            COMMETHOD([], HRESULT, 'RemoveAll'),
            COMMETHOD([], HRESULT, 'GetCurrentView',
                      (['in'], ctypes.POINTER(GUID), 'riid'),
                      (['out'], ctypes.POINTER(ctypes.c_void_p), 'ppv')),
        ]

    _CLSID_ExplorerBrowser = GUID("{71F96385-DDD6-48D3-A0C1-AE06E8B055FB}")

    class _NavEventSink(comtypes.COMObject):
        """IExplorerBrowserEvents sink – receives navigation-complete callbacks."""
        _com_interfaces_ = [_IExplorerBrowserEvents]

        def __init__(self, on_complete):
            super().__init__()
            self._on_complete = on_complete

        def OnNavigationPending(self, pidlFolder):
            return 0  # S_OK

        def OnViewCreated(self, psv):
            return 0  # S_OK

        def OnNavigationComplete(self, pidlFolder):
            # DIAGNOSTIC: this is the raw COM navigation event. If this line
            # stops printing after a minimize/restore while in-shell navigation
            # still visibly happens, the COM event sink has gone deaf (which
            # would leave current_path / _location_url permanently stale).
            debug_print("[IEB NavSink] OnNavigationComplete fired (COM event alive)")
            try:
                if pidlFolder and self._on_complete:
                    _fn = ctypes.windll.shell32.SHGetPathFromIDListW
                    _fn.restype  = ctypes.c_bool
                    _fn.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p]
                    buf = ctypes.create_unicode_buffer(32768)
                    if _fn(pidlFolder, buf):
                        path = buf.value
                        if path:
                            debug_print(f"[IEB NavSink] resolved path: {path}")
                            self._on_complete(path)
            except Exception:
                pass
            return 0  # S_OK

        def OnNavigationFailed(self, pidlFolder):
            return 0  # S_OK


# ── TortoiseGit overlay fix ───────────────────────────────────────────────────
# TortoiseOverlays.dll (registered in HKLM ShellIconOverlayIdentifiers) provides
# GetOverlayInfo correctly (icon paths) but its IsMemberOf returns S_FALSE in
# non-Explorer processes (process name check).
#
# Fix: Patch TortoiseOverlays' IsMemberOf vtable[3] in our process to delegate
# to TortoiseGitStub's IsMemberOf (which works in any process).
# TortoiseOverlays type stored at this+0x08 (0=Normal..8=Unversioned).
# TortoiseGitStub type stored at this+0x28 (1=Normal..9=Unversioned).
# ─────────────────────────────────────────────────────────────────────────────

# TortoiseOverlays CLSIDs (from HKLM, shared vtable in TortoiseOverlays64.dll)
_TORTOISE_OVERLAYS_CLSIDS = [
    '{C5994560-53D9-4125-87C9-F193FC689CB2}',  # Normal (type=0)
    '{C5994561-53D9-4125-87C9-F193FC689CB2}',  # Modified (type=1)
    '{C5994562-53D9-4125-87C9-F193FC689CB2}',  # Conflict (type=2)
    '{C5994563-53D9-4125-87C9-F193FC689CB2}',  # Locked (type=3)
    '{C5994564-53D9-4125-87C9-F193FC689CB2}',  # ReadOnly (type=4)
    '{C5994565-53D9-4125-87C9-F193FC689CB2}',  # Deleted (type=5)
    '{C5994566-53D9-4125-87C9-F193FC689CB2}',  # Added (type=6)
    '{C5994567-53D9-4125-87C9-F193FC689CB2}',  # Ignored (type=7)
    '{C5994568-53D9-4125-87C9-F193FC689CB2}',  # Unversioned (type=8)
]

# TortoiseGitStub CLSIDs (IsMemberOf works in any process)
_TORTOISE_OVERLAY_PATCHED = False


def _patch_tortoise_overlays():
    """Pre-load TortoiseGit overlay icons into the process system image list.

    In non-explorer.exe processes, TortoiseOverlays' IsMemberOf returns S_FALSE,
    but the shell's per-user overlay CACHE (populated by explorer.exe) still
    provides correct overlay indices via IShellIconOverlay::GetOverlayIndex.

    The issue is that the overlay ICONS are not loaded into our process's
    system image list until something triggers them. We call SHGetFileInfo
    with SHGFI_OVERLAYINDEX which forces the shell to lazily load the
    overlay handler's icon (via GetOverlayInfo, which works in any process).

    Once loaded, the shell view (DefView/ItemsView) can render overlays.
    """
    global _TORTOISE_OVERLAY_PATCHED
    if _TORTOISE_OVERLAY_PATCHED:
        return

    try:
        import ctypes
        import ctypes.wintypes
        import winreg

        # Check TortoiseOverlays is installed
        try:
            winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                           rf'SOFTWARE\Classes\CLSID\{_TORTOISE_OVERLAYS_CLSIDS[0]}',
                           0, winreg.KEY_READ).Close()
        except OSError:
            return  # TortoiseOverlays not installed

        _TORTOISE_OVERLAY_PATCHED = True
        print("[TortoiseGit] Overlay icon pre-load enabled "
              "(will trigger on first navigation)")

    except Exception as e:
        print(f"[TortoiseGit] Overlay init error: {e}")

# Cache of directories already preloaded (overlay icons persist in system image list)
_OVERLAY_PRELOADED_DIRS = set()


def _path_is_slow_for_shell(path):
    """模块级慢盘判断：UNC / OneDrive / 映射网络盘上同步 Shell 调用可能阻塞 UI 线程。
    供 _preload_overlay_icons 早退使用（与 FileExplorerTab._is_slow_path 逻辑一致）。"""
    if not path:
        return False
    if path.startswith('\\\\') or path.startswith('//'):
        return True
    if 'onedrive' in path.replace('\\', '/').lower():
        return True
    try:
        import ctypes
        drive = os.path.splitdrive(path)[0]
        if drive and ctypes.windll.kernel32.GetDriveTypeW(drive + '\\') == 4:  # DRIVE_REMOTE
            return True
    except Exception:
        pass
    return False


def _preload_overlay_icons(directory_path):
    """Force-load overlay icons into the system image list by calling
    SHGetFileInfo(SHGFI_OVERLAYINDEX) on files in the given directory.

    This triggers the shell to lazily load overlay icons from handlers
    whose GetOverlayInfo works (TortoiseOverlays provides correct icons).
    The overlay INDEX comes from the cross-process shell cache (populated
    by explorer.exe), so TortoiseOverlays' IsMemberOf doesn't need to work.
    """
    # 慢盘（网络/OneDrive）上 SHGetFileInfo+scandir 是同步阻塞调用，导航前预加载会卡 UI；
    # 跳过即可，overlay 仍由导航完成后 600ms 的 IShellView::Refresh() 兜底渲染。
    if _path_is_slow_for_shell(directory_path):
        return
    try:
        import ctypes
        import ctypes.wintypes

        class _SHFILEINFOW(ctypes.Structure):
            _fields_ = [
                ('hIcon', ctypes.wintypes.HICON),
                ('iIcon', ctypes.c_int),
                ('dwAttributes', ctypes.wintypes.DWORD),
                ('szDisplayName', ctypes.c_wchar * 260),
                ('szTypeName', ctypes.c_wchar * 80),
            ]

        _SHGetFileInfoW = ctypes.windll.shell32.SHGetFileInfoW
        _SHGetFileInfoW.argtypes = [ctypes.c_wchar_p, ctypes.wintypes.DWORD,
                                    ctypes.POINTER(_SHFILEINFOW), ctypes.c_uint,
                                    ctypes.c_uint]
        _SHGetFileInfoW.restype = ctypes.c_void_p

        _DestroyIcon = ctypes.windll.user32.DestroyIcon

        # SHGFI_ICON=0x100, SHGFI_SMALLICON=0x1, SHGFI_OVERLAYINDEX=0x40
        FLAGS = 0x100 | 0x1 | 0x40

        loaded_overlays = set()

        # Query the directory itself first
        sfi = _SHFILEINFOW()
        _SHGetFileInfoW(directory_path, 0, ctypes.byref(sfi),
                        ctypes.sizeof(sfi), FLAGS)
        if sfi.hIcon:
            _DestroyIcon(sfi.hIcon)
        ovl = (sfi.iIcon >> 24) & 0xFF
        if ovl:
            loaded_overlays.add(ovl)

        # Query a few files to trigger overlay icon loading.
        # Once a TortoiseGit overlay (index > 4) is found, we can stop -
        # the overlay slot is process-global and applies to all files.
        MAX_FILES = 3
        count = 0
        found_git_overlay = any(o > 4 for o in loaded_overlays)
        try:
            for entry in os.scandir(directory_path):
                if found_git_overlay or count >= MAX_FILES:
                    break
                # Skip .exe/.msi/.dll - SHGetFileInfo can be very slow for these
                if entry.name.lower().endswith(('.exe', '.msi', '.dll')):
                    continue
                sfi2 = _SHFILEINFOW()
                _SHGetFileInfoW(entry.path, 0, ctypes.byref(sfi2),
                                ctypes.sizeof(sfi2), FLAGS)
                if sfi2.hIcon:
                    _DestroyIcon(sfi2.hIcon)
                ovl2 = (sfi2.iIcon >> 24) & 0xFF
                if ovl2:
                    loaded_overlays.add(ovl2)
                    if ovl2 > 4:
                        found_git_overlay = True
                count += 1
        except OSError:
            pass

        if loaded_overlays:
            debug_print(f"[TortoiseGit] Pre-loaded overlay icons: "
                        f"{sorted(loaded_overlays)} ({count} files scanned)")
            # Mark directory as preloaded (overlay icons are process-global)
            _OVERLAY_PRELOADED_DIRS.add(directory_path)

    except Exception:
        pass


class _OverlayPreloadSignals(QObject):
    """后台 overlay 图标预加载完成信号：done(path)。"""
    done = pyqtSignal(str)


class _OverlayPreloadRunnable(QRunnable):
    """在 QThreadPool 后台线程执行 _preload_overlay_icons（含 SHGetFileInfo 扫描），
    避免在 UI 线程做同步 Shell 调用造成切标签卡顿。完成后经信号回 UI 线程调用
    IShellView::Refresh()（COM 必须在 UI/STA 线程）。工作线程自行 CoInitialize。"""
    def __init__(self, path, signals):
        super().__init__()
        self._path = path
        self._signals = signals

    def run(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            _preload_overlay_icons(self._path)
        except Exception as e:
            debug_print(f"[TortoiseGit] async overlay preload error: {e}")
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            try:
                self._signals.done.emit(self._path)
            except RuntimeError:
                pass  # 信号对象已随标签销毁


class _DirSnapshotSignals(QObject):
    """后台目录快照计算完成信号：done(path, snapshot_or_None)。"""
    done = pyqtSignal(str, object)


class _DirSnapshotRunnable(QRunnable):
    """在 QThreadPool 后台线程计算目录快照，完成后经信号回到 UI 线程比较。

    用于 FileExplorerTab 的 8 秒兜底轮询，避免在 UI 线程逐项 stat 造成周期性卡顿。"""
    def __init__(self, path, ignore_check, signals):
        super().__init__()
        self._path = path
        self._ignore_check = ignore_check
        self._signals = signals

    def run(self):
        snap = _compute_dir_snapshot(self._path, self._ignore_check)
        try:
            self._signals.done.emit(self._path, snap)
        except RuntimeError:
            # 信号对象已随标签销毁：忽略
            pass




# ── Early overlay initialization ─────────────────────────────────────────────
# Must run BEFORE IExplorerBrowser creates its shell view (which triggers SIOM).
if _COMTYPES_AVAILABLE and HAS_PYWIN:
    try:
        _patch_tortoise_overlays()
    except Exception as _e:
        print(f"[TortoiseGit] Early init error: {_e}")


# ── IExplorerBrowser 键盘消息过滤器 ──────────────────────────────────────────
# Qt 的事件循环会把所有 WM_KEYDOWN/WM_KEYUP/WM_CHAR 消息转换成 QKeyEvent。
# 但 IExplorerBrowser 的子窗口（SHELLDLL_DefView / SysListView32）是纯 Win32 控件，
# 它们依赖直接收到 WM_KEYDOWN 来实现 Delete/Ctrl+C/V/X/F2 等功能。
# QAbstractNativeEventFilter 安装在 QApplication 层，能拦截线程中所有 HWND 的消息。
# 我们在此检测键盘消息是否发往 IExplorerBrowser 子窗口，如果是则手动 Dispatch，
# 返回 True 告诉 Qt 跳过后续处理。

from PyQt5.QtCore import QAbstractNativeEventFilter

class _IEBKeyboardFilter(QAbstractNativeEventFilter):
    """Application-level native event filter that forwards keyboard messages
    to IExplorerBrowser child windows so that Delete, Ctrl+C/V/X, F2, etc. work."""

    # TabEx 自有快捷键的 VK 码（Ctrl 组合）
    _TABEX_CTRL_VKS = frozenset([
        0x54,  # T (Ctrl+T new tab, Ctrl+Shift+T reopen)
        0x57,  # W (Ctrl+W close tab)
        0x46,  # F (Ctrl+F search)
        0x47,  # G (Ctrl+G quick find)
        0x44,  # D (Ctrl+D bookmark)
        0x4C,  # L (Ctrl+L focus path bar)
        0x09,  # Tab (Ctrl+Tab switch)
    ])
    # TabEx 自有快捷键的 VK 码（Alt 组合）
    _TABEX_ALT_VKS = frozenset([
        0x5A,  # Z (Alt+Z copy filename)
        0x58,  # X (Alt+X copy path)
        0x25,  # Left (Alt+Left back)
        0x27,  # Right (Alt+Right forward)
        0x26,  # Up (Alt+Up go up)
    ])

    def __init__(self):
        super().__init__()
        self._main_window = None
        self._ieb_hwnds = set()  # 缓存 IExplorerBrowserWidget 的顶层 HWND
        self._ieb_widgets = {}   # HWND → IExplorerBrowserWidget 实例映射
        self._debug_first_key = True  # 首次键盘消息诊断
        self._debug_first_call = True  # 首次调用诊断
        # IShellView COM IID
        self._IID_IShellView = None  # 延迟初始化（需要 comtypes）

    def set_main_window(self, mw):
        self._main_window = mw

    def register_ieb_hwnd(self, hwnd, widget=None):
        """注册 IExplorerBrowserWidget 的 winId，用于快速判断子窗口归属"""
        if hwnd:
            self._ieb_hwnds.add(hwnd)
            if widget is not None:
                self._ieb_widgets[hwnd] = widget
            debug_print(f"[IEB KeyFilter] Registered HWND: 0x{hwnd:X}, total: {len(self._ieb_hwnds)}")

    def unregister_ieb_hwnd(self, hwnd):
        self._ieb_hwnds.discard(hwnd)
        self._ieb_widgets.pop(hwnd, None)

    def _find_owner_widget(self, hwnd):
        """从目标 HWND 向上遍历找到拥有它的 IExplorerBrowserWidget"""
        import ctypes
        _GetParent = ctypes.windll.user32.GetParent
        _GetParent.restype = ctypes.wintypes.HWND
        h = hwnd
        for _ in range(30):
            if h in self._ieb_widgets:
                return self._ieb_widgets[h]
            h = _GetParent(h)
            if not h:
                break
        return None

    def _is_ieb_descendant(self, hwnd):
        """判断 hwnd 是否是已注册的某个 IExplorerBrowserWidget 的后代窗口"""
        import ctypes
        _GetParent = ctypes.windll.user32.GetParent
        _GetParent.restype = ctypes.wintypes.HWND  # 确保 64 位 HWND 不被截断
        h = hwnd
        for _ in range(30):
            if h in self._ieb_hwnds:
                return True
            h = _GetParent(h)
            if not h:
                break
        return False

    # 自定义 MSG 结构体，避免依赖 wintypes.MSG 的字段名
    class _MSG(ctypes.Structure):
        _fields_ = [
            ("hwnd", ctypes.c_void_p),
            ("message", ctypes.c_uint),
            ("wParam", ctypes.c_size_t),   # WPARAM = UINT_PTR
            ("lParam", ctypes.c_ssize_t),  # LPARAM = LONG_PTR
            ("time", ctypes.c_uint),
            ("pt_x", ctypes.c_long),
            ("pt_y", ctypes.c_long),
        ]

    def nativeEventFilter(self, eventType, message):
        try:
            # 首次调用诊断 — 确认过滤器被 Qt 调用
            if self._debug_first_call:
                self._debug_first_call = False
                debug_print(f"[IEB KeyFilter] Filter active! eventType={eventType} "
                            f"ieb_hwnds={[hex(h) for h in self._ieb_hwnds]}")

            if eventType not in (b"windows_generic_MSG", b"windows_dispatcher_MSG"):
                return False, 0
            if not self._ieb_hwnds:
                return False, 0

            import ctypes
            from ctypes import cast, POINTER
            msg = cast(int(message), POINTER(self._MSG)).contents

            # 只处理键盘类消息
            if msg.message not in (0x0100, 0x0101, 0x0102, 0x0104, 0x0105, 0x0106):
                return False, 0

            # 首次键盘消息诊断
            if self._debug_first_key and msg.message == 0x0100:
                self._debug_first_key = False
                debug_print(f"[IEB KeyFilter] First WM_KEYDOWN: VK=0x{msg.wParam & 0xFF:02X} "
                            f"target_hwnd=0x{msg.hwnd or 0:X} "
                            f"registered={[hex(h) for h in self._ieb_hwnds]} "
                            f"is_descendant={self._is_ieb_descendant(msg.hwnd) if msg.hwnd else False}")

            # msg.hwnd 是消息的目标窗口；检查它是否属于 IExplorerBrowser
            target_hwnd = msg.hwnd
            if not target_hwnd or not self._is_ieb_descendant(target_hwnd):
                return False, 0

            # 对 WM_KEYDOWN / WM_SYSKEYDOWN，排除 TabEx 自有快捷键
            if msg.message in (0x0100, 0x0104):
                vk = msg.wParam & 0xFF
                ctrl = (ctypes.windll.user32.GetKeyState(0x11) & 0x8000) != 0
                alt  = (ctypes.windll.user32.GetKeyState(0x12) & 0x8000) != 0
                if ctrl and not alt and vk in self._TABEX_CTRL_VKS:
                    return False, 0  # 留给 TabEx _check_shortcuts
                if alt and not ctrl and vk in self._TABEX_ALT_VKS:
                    return False, 0
                if not ctrl and not alt and vk == 0x74:  # F5
                    return False, 0
                if not ctrl and not alt and vk == 0x72:  # F3 分屏
                    return False, 0

            # 通过 IShellView::TranslateAccelerator 转发键盘消息
            # 这是 Shell 控件处理 Ctrl+C/V/X, Delete, F2 等的正确 COM 方式
            widget = self._find_owner_widget(target_hwnd)
            if widget and getattr(widget, '_browser', None):
                try:
                    if self._IID_IShellView is None:
                        from comtypes import GUID as _GUID
                        self._IID_IShellView = _GUID("{000214E3-0000-0000-C000-000000000046}")

                    # comtypes 的 ['out'] 参数自动返回，不需要手动传
                    ppv_result = widget._browser.GetCurrentView(
                        ctypes.byref(self._IID_IShellView))
                    # ppv_result 是 c_void_p 或整数
                    iface_ptr = int(ppv_result) if ppv_result else 0
                    if iface_ptr:
                        try:
                            # COM 对象布局: 对象地址 → vtable 指针 → 函数指针数组
                            # IShellView vtable: [QI(0), AddRef(1), Release(2),
                            #   GetWindow(3), ContextSensitiveHelp(4), TranslateAccelerator(5)]
                            _vp_size = ctypes.sizeof(ctypes.c_void_p)
                            vtable_ptr = ctypes.c_void_p.from_address(iface_ptr).value
                            fn_addr = ctypes.c_void_p.from_address(vtable_ptr + 5 * _vp_size).value
                            # TranslateAccelerator(IShellView* this, MSG* pmsg) → HRESULT
                            _TA = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_void_p)
                            translate_accel = _TA(fn_addr)
                            # 构造 wintypes.MSG（保证内存布局与 Windows 一致）
                            native_msg = ctypes.wintypes.MSG()
                            native_msg.hWnd = target_hwnd
                            native_msg.message = msg.message
                            native_msg.wParam = msg.wParam
                            native_msg.lParam = msg.lParam
                            native_msg.time = msg.time
                            native_msg.pt.x = msg.pt_x
                            native_msg.pt.y = msg.pt_y
                            ta_hr = translate_accel(iface_ptr, ctypes.byref(native_msg))
                            if ta_hr == 0:  # S_OK — Shell 已处理
                                return True, 0
                        finally:
                            # Release IShellView
                            release_addr = ctypes.c_void_p.from_address(
                                vtable_ptr + 2 * _vp_size).value
                            _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                            _REL(release_addr)(iface_ptr)
                except Exception as e:
                    debug_print(f"[IEB KeyFilter] TranslateAccelerator failed: {e}")

            # 回退：直接 Dispatch（处理 TranslateAccelerator 不支持的按键）
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))
            return True, 0

        except Exception as e:
            debug_print(f"[IEB KeyFilter] Exception: {e}")
            return False, 0


# 全局单例（在 main() 中安装）
_ieb_keyboard_filter = _IEBKeyboardFilter()


class IExplorerBrowserWidget(QWidget):
    """
    Qt widget that hosts IExplorerBrowser – the real Windows Explorer shell
    component. Supports TortoiseGit overlay icons, unlike the legacy
    Shell.Explorer (IE/WebBrowser ActiveX) control.

    Provides a drop-in subset of the QAxWidget API used by FileExplorerTab:
      dynamicCall('Navigate[2](...)', url)  – navigate to a URL or path
      dynamicCall('Refresh()')              – refresh current view
      property('LocationURL')               – current location as file:/// URL
      NavigateComplete2 signal              – emitted on navigation complete
      querySubObject(name)                  – returns None (graceful degradation)
      clear()                               – destroy COM resources
    """
    # (pDisp, url) – matches QAxWidget's NavigateComplete2 signature
    NavigateComplete2 = pyqtSignal(object, object)

    # 异步导航（慢盘）状态信号：宿主标签据此显示 loading / 解除导航锁
    navigationStarted  = pyqtSignal(str)        # path — 后台 PIDL 解析已开始
    navigationFinished = pyqtSignal(str, bool)  # path, ok — 解析失败/无法访问时触发

    # IExplorerBrowser option flags  (SDK shobjidl_core.h values)
    _EBO_SHOWFRAMES  = 0x00000002  # 显示左侧导航窗格
    _EBO_NOTRAVELLOG = 0x00000008  # 禁用前进/后退历史（避免与 TabEx 自身历史冲突）
    _EBO_NOBORDER    = 0x00000040  # 隐藏导航工具栏（地址栏）

    def __init__(self, parent=None):
        super().__init__(parent)
        self._browser      = None
        self._cookie       = ctypes.c_ulong(0)
        self._nav_sink     = None
        self._location_url = ''
        self._pending_path = None
        self._init_ok      = False
        # 异步导航（慢盘）：导航代号用于作废过期的后台解析结果；线程引用防 GC
        self._nav_generation = 0
        self._nav_resolver   = None
        # 后台 overlay 图标预加载：懒创建信号对象 + in-flight 去重标志
        self._overlay_signals = None
        self._overlay_preload_inflight = False
        # Force native HWND creation; IExplorerBrowser::Initialize needs a real HWND
        self.setAttribute(Qt.WA_NativeWindow, True)

    # ── QAxWidget compatibility ───────────────────────────────────────────────

    def property(self, name):  # noqa: A003
        """Override QObject.property to expose LocationURL."""
        if name == 'LocationURL':
            return self._location_url
        return super().property(name)

    def dynamicCall(self, sig, *args):
        """QAxWidget-compatible dispatch for Navigate / Refresh calls."""
        s = sig.lower()
        # Match only actual Navigate/Navigate2 calls, skip event signatures
        # (NavigateComplete2, BeforeNavigate2, etc.)
        if (s.startswith('navigate') or s.startswith('navigate2')) and \
           'complete' not in s and 'before' not in s:
            url = str(args[0]) if args else ''
            if url and url != 'None':
                self._navigate_url(url)
        elif 'refresh()' in s:
            self._do_refresh()
        # All other calls (Visible, ToolBar, Silent, NavigateComplete2 events…) are silently ignored

    def querySubObject(self, _name, *_args):
        """Returns None; selection info is not yet available via IExplorerBrowser."""
        return None

    def clear(self):
        """QAxWidget compat: release COM resources."""
        self.cleanup()

    # ── Navigation helpers ────────────────────────────────────────────────────

    def _navigate_url(self, url):
        path = self._url_to_path(url)
        if path:
            self._navigate(path)

    @staticmethod
    def _url_to_path(url):
        if not url:
            return ''
        url = str(url)
        if url.startswith('file:///'):
            from urllib.parse import unquote
            p = unquote(url[8:]).replace('/', os.sep)
            return p[1:] if p.startswith(os.sep) else p
        if url.startswith('file://') and not url.startswith('file:///'):
            from urllib.parse import unquote
            return '\\\\' + unquote(url[7:]).replace('/', '\\')
        return url  # local path or shell: path – pass through

    def _navigate(self, path):
        if not _COMTYPES_AVAILABLE:
            return
        if not self._ensure_browser():
            self._pending_path = path
            return
        # 慢盘（网络/UNC/OneDrive/映射远程盘）：把阻塞的 PIDL 解析放到后台线程，避免冻结
        # UI 事件循环——一个慢标签不再拖垮其它标签/窗口。本地快盘保持同步（无线程开销）。
        if _path_is_slow_for_shell(path):
            self._navigate_async(path)
        else:
            self._navigate_sync(path)

    def _navigate_sync(self, path):
        """同步导航（本地快盘）：UI 线程内解析 PIDL 并浏览。"""
        try:
            # Pre-load overlay icons BEFORE navigation so DefView has them
            # when it first renders (skip if already cached for this dir)
            if _TORTOISE_OVERLAY_PATCHED and path not in _OVERLAY_PRELOADED_DIRS:
                _preload_overlay_icons(path)

            _spdn = ctypes.windll.shell32.SHParseDisplayName
            _spdn.restype  = ctypes.c_long   # HRESULT
            _spdn.argtypes = [
                ctypes.c_wchar_p,
                ctypes.c_void_p,
                ctypes.POINTER(ctypes.c_void_p),
                ctypes.c_ulong,
                ctypes.POINTER(ctypes.c_ulong),
            ]
            pidl  = ctypes.c_void_p(0)
            sfgao = ctypes.c_ulong(0)
            hr    = _spdn(path, None, ctypes.byref(pidl), 0, ctypes.byref(sfgao))
            if hr == 0 and pidl.value:
                try:
                    self._browser.BrowseToIDList(pidl, 0)  # 0 = SBSP_ABSOLUTE
                finally:
                    ctypes.windll.ole32.CoTaskMemFree(pidl)
            else:
                debug_print(f"[IExplorerBrowser] SHParseDisplayName failed "
                            f"hr=0x{hr & 0xFFFFFFFF:08x} for '{path}'")
        except Exception as e:
            debug_print(f"[IExplorerBrowser] _navigate error: {e}")

    def _navigate_async(self, path):
        """异步导航（慢盘）：后台线程解析 PIDL，解析完再由 UI 线程浏览。"""
        # 递增导航代号：解析期间若再次发起导航，旧线程的结果会被作废，避免错误落地
        self._nav_generation += 1
        gen = self._nav_generation
        try:
            self.navigationStarted.emit(path)
        except Exception:
            pass
        resolver = _PidlResolver(path, gen, self)
        resolver.resolved.connect(self._on_pidl_resolved)
        resolver.finished.connect(resolver.deleteLater)
        self._nav_resolver = resolver
        resolver.start()

    def _on_pidl_resolved(self, path, pidl_val, hr, generation):
        """后台 PIDL 解析完成（UI 线程）：作废过期结果，否则浏览到该 PIDL。"""
        # 期间又发起了新导航 → 丢弃过期结果并释放其 PIDL
        if generation != self._nav_generation:
            if pidl_val:
                try:
                    ctypes.windll.ole32.CoTaskMemFree(ctypes.c_void_p(pidl_val))
                except Exception:
                    pass
            return
        try:
            if pidl_val and self._browser is not None:
                pidl = ctypes.c_void_p(pidl_val)
                try:
                    self._browser.BrowseToIDList(pidl, 0)  # 0 = SBSP_ABSOLUTE
                finally:
                    ctypes.windll.ole32.CoTaskMemFree(pidl)
                # 浏览已发起；内容加载完成由 NavigateComplete2 通知宿主标签隐藏 loading
            else:
                debug_print(f"[IExplorerBrowser] async SHParseDisplayName failed "
                            f"hr=0x{hr & 0xFFFFFFFF:08x} for '{path}'")
                # 解析失败（网络不可达/路径无效）：NavigateComplete2 不会触发，主动通知结束
                if pidl_val:
                    try:
                        ctypes.windll.ole32.CoTaskMemFree(ctypes.c_void_p(pidl_val))
                    except Exception:
                        pass
                try:
                    self.navigationFinished.emit(path, False)
                except Exception:
                    pass
        except Exception as e:
            debug_print(f"[IExplorerBrowser] async _navigate error: {e}")
            try:
                self.navigationFinished.emit(path, False)
            except Exception:
                pass

    def _do_refresh(self):
        if self._location_url:
            self._navigate_url(self._location_url)

    # ── IExplorerBrowser lifecycle ────────────────────────────────────────────

    def _ensure_browser(self):
        """Create/initialize IExplorerBrowser on first use."""
        if self._init_ok:
            return True
        if not _COMTYPES_AVAILABLE:
            return False
        try:
            hwnd = int(self.winId())
            if not hwnd:
                return False
            # TortoiseGit overlay icons are pre-loaded on navigation via
            # _preload_overlay_icons(). No patching needed here.
            browser = comtypes.client.CreateObject(
                _CLSID_ExplorerBrowser,
                interface=_IExplorerBrowser,
            )
            w = max(self.width(),  100)
            h = max(self.height(), 100)
            rc = ctypes.wintypes.RECT(0, 0, w, h)
            fs = _FOLDERSETTINGS(4, 0)  # ViewMode=FVM_DETAILS, fFlags=0
            browser.Initialize(hwnd, ctypes.byref(rc), ctypes.byref(fs))
            # EBO_SHOWFRAMES: 显示左侧导航窗格（目录树）
            browser.SetOptions(self._EBO_SHOWFRAMES | self._EBO_NOTRAVELLOG)
            sink   = _NavEventSink(self._on_nav_complete)
            cookie = browser.Advise(sink)  # out-param returned by comtypes
            self._browser  = browser
            self._nav_sink = sink
            self._cookie   = ctypes.c_ulong(cookie)
            self._init_ok  = True
            _ieb_keyboard_filter.register_ieb_hwnd(hwnd, widget=self)
            debug_print("[IExplorerBrowser] Initialized successfully")
            return True
        except Exception as e:
            debug_print(f"[IExplorerBrowser] Init failed: {e}")
            return False

    def _on_nav_complete(self, path):
        """Called by _NavEventSink when navigation completes."""
        norm = os.path.normpath(path)
        url  = 'file:///' + norm.replace('\\', '/')
        self._location_url = url
        self.NavigateComplete2.emit(None, url)
        # Force shell to re-evaluate icon overlays for this directory
        # Skip if this navigation was triggered by our own Refresh()
        if not getattr(self, '_overlay_refreshing', False):
            self._notify_overlay_refresh(norm)

    def _notify_overlay_refresh(self, path):
        """Refresh overlay rendering after navigation if needed."""
        if not _TORTOISE_OVERLAY_PATCHED:
            return
        # Pre-nav preload already loaded icons; just refresh the view once
        if not getattr(self, '_overlay_refresh_done', False):
            self._overlay_refresh_done = True
            QTimer.singleShot(600, self._do_first_overlay_refresh)

    def _do_first_overlay_refresh(self):
        """One-time refresh to ensure overlays render after first navigation."""
        # Only refresh if this widget is currently visible on screen.
        # Background tabs get refreshed in showEvent when they become active.
        if not self.isVisible():
            self._overlay_refresh_done = False  # allow showEvent to trigger it
            return
        if not getattr(self, '_overlay_refreshing', False):
            self._overlay_refreshing = True
            self._overlay_visible_refreshed = True
            try:
                self._refresh_shell_view()
            finally:
                QTimer.singleShot(500, self._clear_overlay_refreshing)

    def _do_overlay_notify(self, path):
        """后台预加载 overlay 图标，完成后回到 UI 线程刷新视图渲染图标。"""
        if path in _OVERLAY_PRELOADED_DIRS:
            # 已预加载过：无需再扫描，直接在 UI 线程刷新一次即可
            self._overlay_refreshing = True
            self._overlay_visible_refreshed = True
            try:
                self._refresh_shell_view()
            finally:
                QTimer.singleShot(500, self._clear_overlay_refreshing)
            return
        # 预加载（SHGetFileInfo 扫描）放到后台线程，避免在 UI 线程同步 Shell 调用卡顿；
        # in-flight 去重，防止同一控件重复提交。
        if getattr(self, '_overlay_preload_inflight', False):
            return
        try:
            if not getattr(self, '_overlay_signals', None):
                self._overlay_signals = _OverlayPreloadSignals(self)
                self._overlay_signals.done.connect(self._on_overlay_preload_done)
            self._overlay_preload_inflight = True
            QThreadPool.globalInstance().start(_OverlayPreloadRunnable(path, self._overlay_signals))
        except Exception as e:
            self._overlay_preload_inflight = False
            debug_print(f"[TortoiseGit] overlay notify error: {e}")

    def _on_overlay_preload_done(self, _path):
        """后台预加载完成（回到 UI 线程）：调用 IShellView::Refresh() 渲染 overlay。"""
        self._overlay_preload_inflight = False
        try:
            self._overlay_refreshing = True
            self._overlay_visible_refreshed = True
            try:
                self._refresh_shell_view()
            finally:
                QTimer.singleShot(500, self._clear_overlay_refreshing)
        except Exception as e:
            debug_print(f"[TortoiseGit] overlay refresh error: {e}")

    def _clear_overlay_refreshing(self):
        self._overlay_refreshing = False

    def _refresh_shell_view(self):
        """Call IShellView::Refresh() to force the DefView to re-render items."""
        if not self._browser:
            return
        try:
            from comtypes import GUID as _GUID
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_result = self._browser.GetCurrentView(ctypes.byref(iid_sv))
            iface_ptr = int(ppv_result) if ppv_result else 0
            if not iface_ptr:
                return
            try:
                # IShellView vtable layout (inherits IOleWindow):
                # [QI(0), AddRef(1), Release(2), GetWindow(3),
                #  ContextSensitiveHelp(4), TranslateAccelerator(5),
                #  EnableModeless(6), UIActivate(7), Refresh(8), ...]
                _vp_size = ctypes.sizeof(ctypes.c_void_p)
                vtable_ptr = ctypes.c_void_p.from_address(iface_ptr).value
                # Refresh is at vtable index 8
                fn_addr = ctypes.c_void_p.from_address(
                    vtable_ptr + 8 * _vp_size).value
                # IShellView::Refresh(this) → HRESULT
                _REFRESH = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p)
                refresh_fn = _REFRESH(fn_addr)
                hr = refresh_fn(iface_ptr)
                debug_print(f"[TortoiseGit] IShellView::Refresh() → hr=0x{hr & 0xFFFFFFFF:08X}")
            finally:
                # Release the IShellView
                _vp_size = ctypes.sizeof(ctypes.c_void_p)
                vtable_ptr = ctypes.c_void_p.from_address(iface_ptr).value
                release_addr = ctypes.c_void_p.from_address(
                    vtable_ptr + 2 * _vp_size).value
                _RELEASE = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _RELEASE(release_addr)(iface_ptr)
        except Exception as e:
            debug_print(f"[TortoiseGit] _refresh_shell_view error: {e}")

    # ── Qt event overrides ────────────────────────────────────────────────────

    def showEvent(self, event):
        super().showEvent(event)
        if not self._init_ok:
            if self._ensure_browser() and self._pending_path:
                path, self._pending_path = self._pending_path, None
                self._navigate(path)
        elif self._browser and _TORTOISE_OVERLAY_PATCHED:
            # Tab became visible: ensure overlays are rendered
            # Refresh is needed because IShellView::Refresh on hidden views is a no-op
            url = getattr(self, '_location_url', '')
            if url and url.startswith('file:'):
                norm = self._url_to_path(url)
                if norm:
                    if norm not in _OVERLAY_PRELOADED_DIRS:
                        # Not yet preloaded - do full preload + refresh
                        QTimer.singleShot(300, lambda p=norm: self._do_overlay_notify(p))
                    elif not getattr(self, '_overlay_visible_refreshed', False):
                        # Already preloaded but never refreshed while visible
                        self._overlay_visible_refreshed = True
                        QTimer.singleShot(200, self._refresh_shell_view)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._browser:
            try:
                rc = ctypes.wintypes.RECT(0, 0, self.width(), self.height())
                self._browser.SetRect(None, rc)
            except Exception as e:
                debug_print(f"[IExplorerBrowser] SetRect failed: {e}")

    # ── Cleanup ───────────────────────────────────────────────────────────────

    def cleanup(self):
        # 作废任何在途的后台 PIDL 解析，并断开信号，避免结果回调已销毁的 COM 浏览器
        self._nav_generation += 1
        resolver = getattr(self, '_nav_resolver', None)
        if resolver is not None:
            try:
                resolver.resolved.disconnect()
            except Exception:
                pass
            try:
                if resolver.isRunning():
                    resolver.wait(200)  # SHParseDisplayName 无法中断，短等即可，结果会被作废
            except Exception:
                pass
            self._nav_resolver = None
        # 断开后台 overlay 预加载信号，避免延迟到达的 QThreadPool 结果回调已销毁的控件
        sig = getattr(self, '_overlay_signals', None)
        if sig is not None:
            try:
                sig.done.disconnect()
            except Exception:
                pass
            try:
                sig.deleteLater()
            except Exception:
                pass
            self._overlay_signals = None
        self._overlay_preload_inflight = False
        if self._browser is None:
            return
        try:
            if self._cookie.value:
                self._browser.Unadvise(self._cookie)
        except Exception:
            dbg_exc("IExplorerBrowser.Unadvise")
        try:
            self._browser.Destroy()
        except Exception:
            dbg_exc("IExplorerBrowser.Destroy")
        try:
            hwnd = int(self.winId())
            _ieb_keyboard_filter.unregister_ieb_hwnd(hwnd)
        except Exception:
            dbg_exc("IExplorerBrowser.unregister_hwnd")
        self._browser  = None
        self._nav_sink = None
        self._init_ok  = False

    def closeEvent(self, event):
        self.cleanup()
        super().closeEvent(event)


class FileExplorerTab(QWidget):
    # Singleton WH_MOUSE_LL hook for double-click detection (shared across all IEB tabs)
    _global_mouse_hook_handle = None
    _global_mouse_hook_cb = None

    def _force_remove_watcher(self, path):
        """强制移除watcher路径，防止事件风暴"""
        try:
            if path in self.file_watcher.directories():
                self.file_watcher.removePath(path)
                debug_print(f"[FileWatcher] Force removed: {path}")
            if path in self.file_watcher.files():
                self.file_watcher.removePath(path)
                debug_print(f"[FileWatcher] Force removed file: {path}")
        except Exception as e:
            debug_print(f"[FileWatcher] Exception in force_remove_watcher: {e}")

    def _build_dir_snapshot(self, path):
        """构建当前目录的轻量元数据快照，用于检测文件修改时间/大小变化。

        纯计算逻辑委托给模块级 _compute_dir_snapshot()，以便定时兜底轮询可在后台线程
        复用同一算法（见 _poll_directory_changes 的 QRunnable 异步路径），避免 UI 线程逐项 stat。"""
        try:
            if not path or not os.path.isdir(path):
                return None
            # OneDrive/网络路径的os.scandir()会阻塞UI线程（云同步），直接跳过
            if self._is_slow_path(path):
                return None
            return _compute_dir_snapshot(path, self._should_ignore_internal_dir_entry)
        except Exception:
            return None

    def _should_ignore_internal_dir_entry(self, dir_path, entry_name):
        try:
            if not dir_path or not entry_name:
                return False
            app_base_dir = os.path.normcase(os.path.normpath(get_app_base_dir()))
            current_dir = os.path.normcase(os.path.normpath(dir_path))
            if current_dir != app_base_dir:
                return False
            entry_name_lower = str(entry_name).lower()
            if entry_name_lower in APP_INTERNAL_CHANGE_FILENAMES:
                return True
            if entry_name_lower.startswith('config.json.'):
                return True
        except Exception:
            return False
        return False

    def _refresh_file_watch_paths(self, path):
        """同步当前目录下的文件 watcher，提升对文件内容修改的检测能力。"""
        # 文件级 Watcher 已关闭（节省系统句柄资源），目录 Watcher + DirPoll 兜底仍生效
        return
        if not hasattr(self, 'file_watcher'):
            return
        try:
            # 清理旧文件监听
            old_files = getattr(self, '_watched_files', set())
            for file_path in list(old_files):
                self._force_remove_watcher(file_path)
            self._watched_files = set()

            if not path or not os.path.isdir(path):
                return
            # OneDrive/网络路径的os.scandir()会阻塞UI线程，跳过文件监视
            if self._is_slow_path(path):
                debug_print(f"[FileWatcher] Skipping file watch for slow path: {path}")
                return

            max_watch_files = 200
            watched = set()
            with os.scandir(path) as entries:
                for entry in entries:
                    if len(watched) >= max_watch_files:
                        break
                    try:
                        if entry.is_file(follow_symlinks=False):
                            file_path = os.path.normpath(entry.path)
                            if self.file_watcher.addPath(file_path):
                                watched.add(file_path)
                    except Exception:
                        continue
            self._watched_files = watched
            debug_print(f"[FileWatcher] Watching {len(watched)} files in current dir")
        except Exception as e:
            debug_print(f"[FileWatcher] Failed to refresh file watch paths: {e}")

    def set_refresh_active(self, active: bool):
        """设置当前标签的刷新活跃态：仅当前可见标签执行高频刷新。"""
        self._refresh_active = bool(active)
        current_path = getattr(self, 'current_path', '')

        if self._refresh_active:
            if current_path and os.path.isdir(current_path):
                is_slow = self._is_slow_path(current_path)
                # OneDrive/网络路径：跳过os.scandir操作，避免阻塞UI线程和COM消息泵
                if not is_slow:
                    self._refresh_file_watch_paths(current_path)
                    self._last_dir_snapshot = self._build_dir_snapshot(current_path)
                if hasattr(self, 'dir_poll_timer') and self.dir_poll_timer and not self.dir_poll_timer.isActive():
                    if not is_slow:
                        self.dir_poll_timer.start()
                        debug_print(f"[DirPoll] Activated polling for visible tab: {current_path}")
                    else:
                        debug_print(f"[DirPoll] Skipped polling for slow path: {current_path}")
                self._consume_pending_refresh(fallback_reason="activate_tab")
            # 标签激活时，若主同步未运行，启动保活轮询以捕获导航变化
            if not (hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive()):
                self._start_keepalive_sync()
        else:
            self._refresh_file_watch_paths(None)
            if hasattr(self, 'dir_poll_timer') and self.dir_poll_timer and self.dir_poll_timer.isActive():
                self.dir_poll_timer.stop()
                debug_print(f"[DirPoll] Suspended polling for background tab: {current_path}")
            # 标签停用时，停止保活轮询
            if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer:
                self._keepalive_sync_timer.stop()
            # 若标签在导航过程中被切到后台，立即停止瞬时路径同步轮询，
            # 避免后台标签继续以 120~360ms 高频跨进程读取 COM LocationURL
            # （否则需等到下一次 tick 才自停，确保“仅活动标签轮询”）。
            if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive():
                self._path_sync_timer.stop()
            if hasattr(self, '_path_sync_stop_timer') and self._path_sync_stop_timer and self._path_sync_stop_timer.isActive():
                self._path_sync_stop_timer.stop()

    def is_auto_refresh_frozen(self):
        return bool(getattr(self, '_manual_refresh_frozen', False))

    def set_auto_refresh_frozen(self, frozen):
        self._manual_refresh_frozen = bool(frozen)
        if self._manual_refresh_frozen and hasattr(self, 'refresh_timer') and self.refresh_timer.isActive():
            self.refresh_timer.stop()
        if not self._manual_refresh_frozen:
            self._consume_pending_refresh(fallback_reason="manual_unfreeze")
        return self._manual_refresh_frozen

    def _arm_selection_guard(self, seconds=6.0):
        """Ctrl+G 选中后，开启一段时间的选中保护窗口，期间抑制自动刷新。"""
        try:
            self._selection_guard_until = time.monotonic() + float(seconds)
        except Exception:
            self._selection_guard_until = 0.0

    def _selection_guard_active(self):
        try:
            return time.monotonic() < float(getattr(self, '_selection_guard_until', 0) or 0)
        except Exception:
            return False

    def _request_refresh(self, reason="manual"):
        """统一记录刷新请求，当前标签可见时立即调度，不可见时延后到激活后消费。"""
        self._refresh_pending = True
        self._refresh_pending_reason = reason

        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[AutoRefresh] Suppressed during navigation (reason={reason})")
            return False
        if self._selection_guard_active():
            debug_print(f"[AutoRefresh] Suppressed during selection guard (reason={reason})")
            return False
        if getattr(self, '_manual_refresh_frozen', False):
            debug_print(f"[AutoRefresh] Manually frozen (reason={reason})")
            return False
        if not getattr(self, '_refresh_active', True):
            debug_print(f"[AutoRefresh] Deferred while tab inactive (reason={reason})")
            return False

        self._schedule_refresh(reason=reason)
        return True

    def _consume_pending_refresh(self, fallback_reason="manual"):
        if not getattr(self, '_refresh_pending', False):
            return False
        if getattr(self, '_manual_refresh_frozen', False):
            return False
        reason = getattr(self, '_refresh_pending_reason', None) or fallback_reason
        self._schedule_refresh(reason=reason)
        return True

    def on_file_changed(self, path):
        """文件内容或元数据变化时，主动刷新当前视图。"""
        try:
            import time
            now_ms = time.time() * 1000
            norm_file = os.path.normcase(os.path.normpath(path))
            # 文件级去抖：同一路径短时间内重复事件只处理一次
            last_ms = self._last_file_event.get(norm_file, 0)
            if now_ms - last_ms < getattr(self, '_file_event_debounce_ms', 400):
                return
            self._last_file_event[norm_file] = now_ms
            if len(self._last_file_event) > 200:
                # 仅保留最近触发的100条，避免字典无限增长
                items = sorted(self._last_file_event.items(), key=lambda x: x[1], reverse=True)
                self._last_file_event = dict(items[:100])

            debug_print(f"[FileWatcher] File changed: {path}")
            current_dir = getattr(self, 'current_path', '')
            if current_dir and os.path.normcase(os.path.dirname(path)) == os.path.normcase(current_dir):
                if not getattr(self, '_refresh_active', True):
                    self._request_refresh(reason="file_changed")
                    debug_print(f"[FileWatcher] Background tab file changed, marked dirty: {path}")
                    return
                # 某些编辑器会以重命名替换文件，变化后需重新添加 watcher
                if os.path.isfile(path):
                    try:
                        if path not in self.file_watcher.files():
                            self.file_watcher.addPath(path)
                    except Exception:
                        pass
                self._request_refresh(reason="file_changed")
        except Exception as e:
            debug_print(f"[FileWatcher] on_file_changed error: {e}")

    def update_tab_title(self):
        if hasattr(self, 'current_path'):
            # 兜底同步路径栏：有些导航路径变化来自 Explorer 内部事件，
            # 可能只触发标题更新，不走 navigate_to()。
            try:
                if hasattr(self, 'path_bar') and self.path_bar:
                    self.path_bar.set_path(self.current_path)
            except Exception:
                pass

            # shell: 路径中文映射（字典 O(1) 查找）
            _SHELL_PATH_MAP = {
                'shell:RecycleBinFolder': tr('回收站'),
                'shell:MyComputerFolder': tr('此电脑'),
                'shell:Desktop': '桌面',
                'shell:NetworkPlacesFolder': tr('网络'),
            }
            path = self.current_path
            display = _SHELL_PATH_MAP.get(path, None)
            
            if not display and path.startswith('shell:'):
                display = path  # 兜底显示原始shell:路径
            
            if not display:
                # 普通路径：显示最后一层文件夹名称
                # 标准化路径分隔符
                normalized_path = path.replace('/', '\\') if os.name == 'nt' else path
                # 移除末尾的分隔符
                normalized_path = normalized_path.rstrip('\\/')
                
                if normalized_path:
                    # 获取最后一层文件夹名称
                    folder_name = os.path.basename(normalized_path)
                    
                    # 如果是驱动器根目录（如 C:），直接显示
                    if not folder_name and ':' in normalized_path:
                        folder_name = normalized_path
                    # 如果是UNC路径根目录（如 \\server\share），显示share
                    elif not folder_name and normalized_path.startswith('\\\\'):
                        parts = normalized_path.split('\\')
                        folder_name = parts[-1] if parts[-1] else parts[-2] if len(parts) > 2 else normalized_path
                    
                    display = folder_name if folder_name else path
                else:
                    display = path
            
            # 处理固定标签和普通标签的显示
            is_pinned = getattr(self, 'is_pinned', False)
            pin_prefix = "📌 " if is_pinned else ""
            
            # 如果是固定标签，限制display长度以确保📌始终显示
            # 标签宽度120px，📌+空格约占15px，剩余约105px可显示文本
            # 一个中文字符约12px，英文约7px，预估最多显示约15个字符
            if is_pinned:
                max_display_len = 15  # 为📌预留空间
                if len(display) > max_display_len:
                    display = "..." + display[-(max_display_len-3):]
            
            title = pin_prefix + display
            debug_print(f"DEBUG update_tab_title: path={path}, is_pinned={is_pinned}, pin_prefix='{pin_prefix}', title='{title}'")
            mw = self.main_window
            if mw and hasattr(mw, 'tab_widget'):
                # 定位该标签所属的组（左侧主组或右侧分屏组），只更新其所在标签栏
                target_tw = None
                idx = -1
                if hasattr(mw, '_all_groups'):
                    for cand_tw, cand_cs in mw._all_groups():
                        if cand_cs is not None:
                            j = cand_cs.indexOf(self)
                            if j != -1:
                                target_tw, idx = cand_tw, j
                                break
                if target_tw is None:
                    # 回退：默认左侧主组（因 tab_widget 只含占位符，需在 content_stack 中查找索引）
                    target_tw = mw.tab_widget
                    if hasattr(mw, 'content_stack'):
                        idx = mw.content_stack.indexOf(self)
                    else:
                        idx = mw.tab_widget.indexOf(self)

                if idx != -1:
                    target_tw.setTabText(idx, title)
                    debug_print(f"DEBUG: Set tab {idx} text to '{title}'")
                    if hasattr(mw, '_schedule_session_snapshot'):
                        mw._schedule_session_snapshot()

    def start_path_sync_timer(self, duration_ms=2000):
        """启动路径同步定时器，duration_ms 后自动停止（按需触发，减少持续COM调用）"""
        from PyQt5.QtCore import QTimer
        # 主同步启动时暂停保活轮询，避免双重COM查询
        if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer and self._keepalive_sync_timer.isActive():
            self._keepalive_sync_timer.stop()
        self._path_sync_stable_hits = 0
        self._path_sync_interval_ms = 120
        if not hasattr(self, '_path_sync_timer') or self._path_sync_timer is None:
            self._path_sync_timer = QTimer(self)
            self._path_sync_timer.timeout.connect(self.sync_path_bar_with_explorer)
        # 设置自动停止定时器：导航期间轮询，稳定后停止
        if not hasattr(self, '_path_sync_stop_timer') or self._path_sync_stop_timer is None:
            self._path_sync_stop_timer = QTimer(self)
            self._path_sync_stop_timer.setSingleShot(True)
            self._path_sync_stop_timer.timeout.connect(self._stop_path_sync_timer)
        self._path_sync_stop_timer.start(duration_ms)
        if not self._path_sync_timer.isActive():
            self._path_sync_timer.start(self._path_sync_interval_ms)
        elif self._path_sync_timer.interval() != self._path_sync_interval_ms:
            self._path_sync_timer.setInterval(self._path_sync_interval_ms)

    def _stop_path_sync_timer(self):
        """停止路径同步轮询（稳定后调用，避免持续COM跨进程调用）"""
        if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive():
            self._path_sync_timer.stop()
        self._path_sync_stable_hits = 0
        self._path_sync_interval_ms = 120
        # 主同步结束后，启动低频保活轮询以兜底 NavigateComplete2 遗漏的用户导航
        if getattr(self, '_refresh_active', False):
            self._start_keepalive_sync()

    def _read_location_url_timed(self):
        """读取 LocationURL（同步跨进程 COM），并测量耗时以检测系统高负载。

        返回读取到的 URL；若本次调用耗时超过 COM_POLL_SLOW_MS，则设置高负载退避截止
        时间戳 self._com_stress_until，供轮询函数据此拉长间隔，打断 CPU 高时的死亡螺旋。

        对真正跨进程的旧 QAx Shell.Explorer 控件，改用带硬超时的
        看门狗读取：调用挂死时 UI 线程最多阻塞 COM_POLL_HARD_DEADLINE_MS 即放弃并返回
        上次已知 URL，彻底根治“无超时同步 COM 卡死 UI”。IExplorerBrowser 的 LocationURL
        是进程内缓存值（快），直接在 UI 线程读取。"""
        # IEB 缓存值：进程内、零阻塞，直接读
        if isinstance(self.explorer, IExplorerBrowserWidget):
            return self.explorer.property('LocationURL')
        # 旧 QAx：跨进程 COM，加看门狗硬超时
        start = time.perf_counter()
        url = self._read_qax_location_with_watchdog()
        elapsed_ms = (time.perf_counter() - start) * 1000.0
        if elapsed_ms >= COM_POLL_SLOW_MS:
            self._com_stress_until = time.monotonic() + (COM_POLL_STRESS_BACKOFF_MS / 1000.0)
            debug_print(f"[PathSync] Slow LocationURL read {elapsed_ms:.0f}ms -> backoff "
                        f"{COM_POLL_STRESS_BACKOFF_MS}ms")
        if url:
            self._last_location_url = url
        return url

    def _read_qax_location_with_watchdog(self):
        """带硬超时的 QAx LocationURL 读取：超时则返回上次缓存值，避免阻塞 UI。

        看门狗线程读取跨进程 COM 属性前先 CoInitialize（STA），避免在未初始化
        COM 的工作线程上调用导致失败；同时用 _com_inflight 抑制重叠提交，避免超时后旧任务
        仍占用单 worker 造成 future 堆积。"""
        from concurrent.futures import ThreadPoolExecutor, TimeoutError as _FTimeout
        # 上一次读取尚未返回：不重复提交，直接用缓存值（worker 阻塞时避免任务堆积）
        if getattr(self, '_com_inflight', False):
            return getattr(self, '_last_location_url', '')
        ex = getattr(self, '_com_executor', None)
        if ex is None:
            ex = ThreadPoolExecutor(max_workers=1, thread_name_prefix='com-loc')
            self._com_executor = ex
        self._com_inflight = True
        try:
            fut = ex.submit(self._qax_read_location_worker)
            return fut.result(timeout=COM_POLL_HARD_DEADLINE_MS / 1000.0)
        except _FTimeout:
            self._com_stress_until = time.monotonic() + (COM_POLL_STRESS_BACKOFF_MS / 1000.0)
            debug_print(f"[PathSync] LocationURL watchdog timeout -> using cached, backoff")
            return getattr(self, '_last_location_url', '')
        except Exception as e:
            debug_print(f"[PathSync] watchdog read error: {e}")
            return getattr(self, '_last_location_url', '')

    def _qax_read_location_worker(self):
        """worker 线程：初始化 STA 公寓后读取 LocationURL，结束后清除 in-flight 标志。"""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            return self.explorer.property('LocationURL')
        finally:
            self._com_inflight = False
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass



    def _com_under_stress(self):
        """当前是否处于 COM 高负载退避窗口内。"""
        return time.monotonic() < getattr(self, '_com_stress_until', 0.0)

    def _poll_burst_guard(self):
        """挂钟防抖：吸收 UI 线程卡顿后堆积的爆发式定时器触发。

        若距上次实际轮询的真实时间间隔小于 COM_POLL_MIN_GAP_MS，返回 True 表示应跳过本次，
        避免连续多次慢 COM 调用把线程进一步压死。"""
        now_ms = time.monotonic() * 1000.0
        last_ms = getattr(self, '_last_poll_wall_ms', 0.0)
        if now_ms - last_ms < COM_POLL_MIN_GAP_MS:
            return True
        self._last_poll_wall_ms = now_ms
        return False

    def _start_keepalive_sync(self):
        """启动低频保活轮询（每1500ms检测LocationURL变化，主同步未覆盖时兜底）"""
        # 如果主同步正在运行，不启动保活（避免双重轮询）
        if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive():
            return
        if not hasattr(self, '_keepalive_sync_timer') or self._keepalive_sync_timer is None:
            self._keepalive_sync_timer = QTimer(self)
            self._keepalive_sync_timer.timeout.connect(self._keepalive_sync_check)
        if not self._keepalive_sync_timer.isActive():
            self._keepalive_sync_timer.start(1500)

    def _keepalive_sync_check(self):
        """保活检查：若检测到LocationURL与current_path不一致，重启主同步更新路径栏。"""
        if not getattr(self, '_refresh_active', False):
            if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer:
                self._keepalive_sync_timer.stop()
            return
        # 窗口最小化或失去前台焦点时跳过COM轮询：用户无法在内嵌shell视图中导航，
        # 窗口最小化时跳过COM轮询：用户无法在内嵌shell视图中导航，
        # 持续读取 LocationURL 只会触发后台 shell worker 线程 churn（见 runtime_health.log）。
        # 注意：不要用 isActiveWindow() 判断——内嵌 IExplorerBrowser/QAx 控件经常抢占焦点，
        # 主窗口虽在前台却报告非激活，会导致路径同步永久休眠、地址栏冻结需手动 resize 才恢复。
        mw = getattr(self, 'main_window', None)
        if mw is not None:
            try:
                if mw.windowState() & Qt.WindowMinimized:
                    return
            except Exception:
                pass
        # 抗高负载：爆发触发跳过 + 退避窗口内跳过 COM 读取（与主同步一致），避免 CPU 高时压死 UI 线程
        if self._poll_burst_guard():
            return
        if self._com_under_stress():
            return
        try:
            url = self._read_location_url_timed()
            if url:
                url_str = str(url)
                local_path = None
                if url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = unquote(url_str[8:])
                    if os.name == 'nt' and local_path.startswith('/'):
                        local_path = local_path[1:]
                    local_path = self._normalize_local_path(local_path)
                elif url_str.startswith('file://') and not url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = self._normalize_local_path('\\\\' + unquote(url_str[7:]).replace('/', '\\'))
                elif '::' in url_str:
                    # 尝试从 CLSID URL 中提取盘符路径（如映射网络驱动器 ::{...}\X:\）
                    import re
                    drive_match = re.search(r'([A-Za-z]:\\[^"]*)', url_str.replace('/', '\\'))
                    if not drive_match:
                        drive_match = re.search(r'([A-Za-z]:/[^"]*)', url_str)
                    if drive_match:
                        local_path = self._normalize_local_path(drive_match.group(1))
                current_path = self._normalize_local_path(getattr(self, 'current_path', ''))
                if local_path and local_path != current_path:
                    debug_print(f"[Keepalive] Navigation detected: {current_path!r} -> {local_path!r}, updating path bar")
                    # 保活轮询仅在无程序化导航（主同步定时器停止）时运行，
                    # 因此此处检测到的差异必为真实用户导航（如双击进入子目录）。
                    # 直接更新路径栏，绕过 _suppress_auto_refresh 门控，
                    # 避免 NavigateComplete2 偶发遗漏时路径栏长时间停留在旧路径。
                    self.current_path = local_path
                    if hasattr(self, 'path_bar') and self.path_bar:
                        self.path_bar.set_path(local_path)
                    self.update_tab_title()
                    self._schedule_status_update(track_selection=True)
                    if not getattr(self, '_navigating_programmatically', False) and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer:
                        self._keepalive_sync_timer.stop()
                    self.start_path_sync_timer(duration_ms=3000)
        except Exception as e:
            debug_print(f"[Keepalive] ERROR: {e}")

    def sync_path_bar_with_explorer(self):
        # 通过QAxWidget的LocationURL属性获取当前路径
        # 若当前标签页不是活跃标签（_refresh_active=False），直接停止同步（避免背景标签持续COM轮询）
        if not getattr(self, '_refresh_active', True):
            self._stop_path_sync_timer()
            return
        # 抗高负载①：卡顿后堆积的爆发式定时器触发直接跳过，避免连续慢 COM 调用进一步压死 UI 线程
        if self._poll_burst_guard():
            return
        # 抗高负载②：高负载退避窗口内拉长轮询间隔并跳过本次 COM 读取，给 UI 线程喘息；
        # 路径仍由 NavigateComplete2 信号与保活轮询兜底更新，窗口到期后自动恢复正常频率。
        if self._com_under_stress():
            t = getattr(self, '_path_sync_timer', None)
            if t and t.interval() < COM_POLL_STRESS_BACKOFF_MS:
                t.setInterval(COM_POLL_STRESS_BACKOFF_MS)
            return
        try:
            url = self._read_location_url_timed()
            if url:
                url_str = str(url)
                local_path = None
                
                # 处理 file:/// 本地路径
                if url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = unquote(url_str[8:])
                    if os.name == 'nt' and local_path.startswith('/'):
                        local_path = local_path[1:]
                    local_path = self._normalize_local_path(local_path)
                # 处理 file://server/share 网络(UNC)路径
                elif url_str.startswith('file://') and not url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = '\\\\' + unquote(url_str[7:]).replace('/', '\\')
                    local_path = self._normalize_local_path(local_path)
                # 处理 shell: 特殊路径
                elif url_str.startswith('shell:') or '::' in url_str:
                    # 尝试从 CLSID URL 中提取盘符路径（如映射网络驱动器 ::{...}\X:\）
                    import re
                    drive_match = re.search(r'([A-Za-z]:\\[^"]*)', url_str.replace('/', '\\'))
                    if not drive_match:
                        drive_match = re.search(r'([A-Za-z]:/[^"]*)', url_str)
                    if drive_match:
                        local_path = self._normalize_local_path(drive_match.group(1))
                        debug_print(f"[PathSync] Extracted drive path from CLSID URL: {local_path}")
                    else:
                        # 纯 shell: 路径，无法提取盘符，不更新
                        return
                
                current_path = self._normalize_local_path(self.current_path)

                if local_path and local_path != current_path:
                    # 程序化导航期间（Navigate2 已发出但 Shell.Explorer 尚未完成），
                    # LocationURL 可能仍返回旧路径，此时不回写，避免路径栏倒退。
                    # NavigateComplete2 信号会在导航真正完成后更新路径。
                    if getattr(self, '_suppress_auto_refresh', False):
                        return
                    # 窗口恢复后树面板自动展开的虚假导航，抑制
                    restore_guard = getattr(self, '_restore_guard_until', 0)
                    if restore_guard and time.monotonic() < restore_guard:
                        return
                    self._path_sync_stable_hits = 0
                    self._path_sync_interval_ms = 120
                    if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.interval() != self._path_sync_interval_ms:
                        self._path_sync_timer.setInterval(self._path_sync_interval_ms)
                    self.current_path = local_path
                    if hasattr(self, 'path_bar'):
                        self.path_bar.set_path(local_path)
                    self.update_tab_title()
                    self._schedule_status_update(track_selection=True)
                    # 只在非程序化导航时添加到历史记录
                    if not self._navigating_programmatically and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    # 导航完成后清理标志并恢复同步
                    if getattr(self, '_navigating_folder', False):
                        self._navigating_folder = False
                    self._resume_path_sync_after_navigation()
                elif local_path and local_path == current_path:
                    self._path_sync_stable_hits = int(getattr(self, '_path_sync_stable_hits', 0)) + 1
                    # 即使路径未变化，也做一次轻量UI回写，修复偶发地址栏未重绘。
                    # 优化：若面包屑已存在，只做 repaint 而非全量重建，减少控件析构/创建开销。
                    if hasattr(self, 'path_bar'):
                        pb = self.path_bar
                        if not getattr(pb, '_in_edit', False):
                            # 单控件自绘面包屑：轻量更新分段并同步重绘，无子控件重建开销
                            try:
                                pb.set_path(local_path)
                                pb.repaint()
                            except Exception:
                                pass
                    if hasattr(self, '_path_sync_timer') and self._path_sync_timer:
                        target_interval = 220 if self._path_sync_stable_hits == 1 else 360
                        if self._path_sync_timer.interval() != target_interval:
                            self._path_sync_timer.setInterval(target_interval)
                    if self._path_sync_stable_hits >= 2:
                        # 若停止定时器剩余时间较长（说明调用方设置了长窗口，例如 SelectFile 后的 8s），
                        # 不提前退出——继续以 360ms 低频轮询，等待用户按 Enter 进入子目录。
                        # 若剩余时间较短（正常 2s 窗口快到期），才做提前停止优化。
                        _remaining_ms = 0
                        try:
                            if (hasattr(self, '_path_sync_stop_timer') and self._path_sync_stop_timer
                                    and self._path_sync_stop_timer.isActive()):
                                _remaining_ms = self._path_sync_stop_timer.remainingTime()
                        except Exception:
                            pass
                        if _remaining_ms <= 2000:
                            # 导航已确认完成（Shell.Explorer 稳定在 current_path），
                            # 立即解除抑制标志，避免用户随后双击子目录时路径栏无法更新。
                            self._suppress_auto_refresh = False
                            self._stop_path_sync_timer()
        except Exception as e:
            debug_print(f"[PathSync] ERROR: {e}")

    def _resume_path_sync_after_navigation(self):
        """导航后重启路径同步定时器（短窗口轮询，自动停止）"""
        try:
            self.start_path_sync_timer(duration_ms=2000)
        except Exception as e:
            debug_print(f"[PathSync] ERROR resuming timer: {e}")

    def _schedule_status_update(self, delay_ms=STATUS_UPDATE_DEFER_MS, track_selection=False):
        if track_selection:
            self._start_status_tracking()
        if delay_ms <= 0:
            if self.status_update_timer.isActive():
                self.status_update_timer.stop()
            self.update_explorer_status()
            return
        if self.status_update_timer.isActive() and self.status_update_timer.remainingTime() <= delay_ms:
            return
        self.status_update_timer.start(delay_ms)

    def _start_status_tracking(self, duration_ms=STATUS_TRACKING_WINDOW_MS):
        self._status_tracking_deadline_ms = int(time.time() * 1000) + int(duration_ms)
        if not self.status_tracking_timer.isActive():
            self.status_tracking_timer.start()

    def _stop_status_tracking(self):
        self._status_tracking_deadline_ms = 0
        if self.status_tracking_timer.isActive():
            self.status_tracking_timer.stop()

    def _poll_status_during_interaction(self):
        self.update_explorer_status()
        now_ms = int(time.time() * 1000)
        if now_ms >= int(getattr(self, '_status_tracking_deadline_ms', 0) or 0):
            self._stop_status_tracking()

    def setup_ui(self):
        from PyQt5.QtWidgets import QLabel
        # 设置FileExplorerTab背景为白色
        self.setStyleSheet("background: white;")
        self.setAutoFillBackground(True)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 路径栏（极简单行输入框）
        self.path_bar = SimplePathBar(self)
        self.path_bar.pathChanged.connect(self.on_path_bar_changed)
        layout.addWidget(self.path_bar)
        
        # 加载指示器（初始隐藏）
        # 注意：改为悬浮覆盖层，不加入垂直布局，避免显示/隐藏时顶部区域高度跳动
        # （以前进度条占用布局高度，show 时路径栏下方多出 20px，看起来像路径栏“变两倍高又恢复”）
        self.loading_bar = QProgressBar(self)
        self.loading_bar.setMaximum(0)  # 不确定进度模式
        self.loading_bar.setTextVisible(True)
        self.loading_bar.setFormat(tr("正在加载大文件夹..."))
        self.loading_bar.setFixedHeight(20)
        self.loading_bar.hide()

        # 优先使用 IExplorerBrowser（真实 Windows 资源管理器外壳，支持 TortoiseGit 图标覆盖）
        # 回退到 Shell.Explorer（IE/WebBrowser ActiveX，不支持 TortoiseGit 图标覆盖）
        _ieb_ok = False
        if _COMTYPES_AVAILABLE:
            try:
                self.explorer = IExplorerBrowserWidget(self)
                _ieb_ok = True
                debug_print("[WindowsShellExplorer] Using IExplorerBrowser (TortoiseGit overlay supported)")
            except Exception as _ieb_err:
                debug_print(f"[WindowsShellExplorer] IExplorerBrowserWidget failed: {_ieb_err}")
        if not _ieb_ok:
            self.explorer = QAxWidget(self)
            if not self.explorer.setControl("Shell.Explorer"):
                raise RuntimeError("Shell.Explorer control initialization failed")
            debug_print("[WindowsShellExplorer] Using Shell.Explorer (no TortoiseGit overlay support)")
        
        # 设置为NoFocus，防止QAxWidget拦截键盘事件
        self.explorer.setFocusPolicy(Qt.NoFocus)
        # 允许Explorer控件横向压缩，减小右侧面板最小宽度
        try:
            self.explorer.setMinimumWidth(0)
        except Exception:
            pass
        layout.addWidget(self.explorer)

        # 状态栏（参考系统 Explorer 样式：细高、浅底色、顶部分割线）
        self.status_bar = QLabel(tr("就绪"))
        self.status_bar.setFixedHeight(20)
        self.status_bar.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.status_bar.setStyleSheet(
            "QLabel { padding: 2px 8px; background: white; border-top: 1px solid #e0e0e0; font-size: 12px; color: #444; }"
        )
        # 支持在状态栏上按住拖动整个软件窗口
        self._status_drag_pos = None
        self.status_bar.setCursor(Qt.SizeAllCursor)
        self.status_bar.mousePressEvent = self._status_bar_mouse_press
        self.status_bar.mouseMoveEvent = self._status_bar_mouse_move
        self.status_bar.mouseReleaseEvent = self._status_bar_mouse_release
        self.status_bar.mouseDoubleClickEvent = self._status_bar_mouse_double_click
        # 右侧资源占用标签（CPU/内存），默认隐藏，可在设置中开启
        self.resource_label = QLabel("")
        self.resource_label.setFixedHeight(20)
        self.resource_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.resource_label.setStyleSheet(
            "QLabel { padding: 2px 10px; background: white; border-top: 1px solid #e0e0e0; font-size: 12px; }"
        )
        self.resource_label.mouseDoubleClickEvent = self._status_bar_mouse_double_click
        self.resource_label.hide()
        status_row = QHBoxLayout()
        status_row.setContentsMargins(0, 0, 0, 0)
        status_row.setSpacing(0)
        status_row.addWidget(self.status_bar, 1)
        status_row.addWidget(self.resource_label, 0)
        layout.addLayout(status_row)

        
        # 异步加载相关
        self.folder_checker = None  # 文件夹大小检查线程
        self.pending_navigation = None  # 待处理的导航请求
        # 慢盘异步导航状态：_nav_in_progress 表示后台 PIDL 解析未完成（用于显示 loading
        # 与拦截重复点击）；_nav_in_progress_path 记录当前正在解析的目标路径。
        self._nav_in_progress = False
        self._nav_in_progress_path = None
        
        # Explorer 基础配置：保留必要项，避免重复 COM 调用拖慢初始化
        self.explorer.dynamicCall('Visible', True)
        self.explorer.dynamicCall('RegisterAsBrowser', True)
        self.explorer.dynamicCall('RegisterAsDropTarget', True)
        self.explorer.dynamicCall('TheaterMode', False)
        self.explorer.dynamicCall('ToolBar', False)
        self.explorer.dynamicCall('StatusBar', False)
        self.explorer.dynamicCall('MenuBar', False)
        self.explorer.dynamicCall('AddressBar', False)
        self.explorer.dynamicCall('Resizable', True)
        self.explorer.dynamicCall('FullScreen', False)
        self.explorer.dynamicCall('Offline', False)
        self.explorer.dynamicCall('Silent', True)
        # 预绑定关键导航事件签名，避免首次导航时事件分发延迟
        self.explorer.dynamicCall('NavigateComplete2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('DocumentComplete(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('BeforeNavigate2(QVariant,QVariant,QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None, None, None)


        # 兼容原有空白双击（保留控件但不占用空间，避免底部留白）
        self.blank = QLabel()
        self.blank.setFixedHeight(0)
        self.blank.setStyleSheet("background: transparent;")
        self.blank.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.blank.mouseDoubleClickEvent = self.blank_double_click
        # 不再额外增加可见高度
        layout.addWidget(self.blank)

        # 安装事件过滤器以捕获 Explorer 的鼠标按下与双击事件
        try:
            self.explorer.installEventFilter(self)
        except Exception:
            pass

        # 直接连接 NavigateComplete2 信号，实现路径栏即时更新（无 polling 延迟）
        try:
            self.explorer.NavigateComplete2.connect(self._on_shell_navigate_complete)
            debug_print("[Explorer] NavigateComplete2 signal connected")
        except (AttributeError, TypeError) as e:
            debug_print(f"[Explorer] NavigateComplete2 direct signal unavailable: {e}")

        # 连接异步导航（慢盘）状态信号：显示 loading + 进入/解除“导航中”锁定
        try:
            self.explorer.navigationStarted.connect(self._on_async_nav_started)
            self.explorer.navigationFinished.connect(self._on_async_nav_finished)
        except (AttributeError, TypeError):
            pass  # Shell.Explorer 回退控件无此信号，忽略

        # 初始设置路径栏（确保路径栏显示初始路径）
        if hasattr(self, 'path_bar'):
            self.path_bar.set_path(self.current_path)
        
        # 启动路径同步定时器
        self.start_path_sync_timer()

        self.update_explorer_status()
        
        # 初始导航到当前路径（在setup_ui最后调用，确保所有设置已应用）
        self.explorer.dynamicCall('Navigate(const QString&)', QDir.toNativeSeparators(self.current_path))

    def _on_async_nav_started(self, path):
        """慢盘异步导航开始：显示 loading 并进入导航中锁定状态。"""
        self._nav_in_progress = True
        self._nav_in_progress_path = self._normalize_local_path(path)
        self._show_loading_indicator()
        # 安全兜底：网络中断时 NavigateComplete2 可能永不触发，超时后强制解锁并隐藏 loading
        QTimer.singleShot(
            ASYNC_NAV_TIMEOUT_MS,
            lambda p=self._nav_in_progress_path: self._async_nav_safety_timeout(p),
        )

    def _on_async_nav_finished(self, path, ok):
        """慢盘异步导航结束（仅解析失败/无法访问时触发）：解锁并隐藏 loading。"""
        self._nav_in_progress = False
        self._hide_loading_indicator()
        if not ok:
            show_toast(self, tr("路径错误"), tr("无法访问: {}").format(path), level="warning")

    def _async_nav_safety_timeout(self, path):
        """导航中锁定的兜底超时：仍在等待同一目标时强制解锁，避免永久卡在加载态。"""
        if (getattr(self, '_nav_in_progress', False) and
                getattr(self, '_nav_in_progress_path', None) == path):
            debug_print(f"[AsyncNav] Safety timeout, clearing nav lock: {path}")
            self._nav_in_progress = False
            self._hide_loading_indicator()

    def _clear_async_nav_lock(self):
        """导航真正完成（NavigateComplete2）时解除导航中锁定并隐藏 loading。"""
        if getattr(self, '_nav_in_progress', False):
            self._nav_in_progress = False
            self._hide_loading_indicator()

    def _on_shell_navigate_complete(self, *args):
        """Shell.Explorer NavigateComplete2 直接信号处理：路径栏即时更新"""
        try:
            # NavigateComplete2 签名: (IDispatch* pDisp, VARIANT* URL)
            # PyQt5 传入参数可能是 (dispatch, url) 或仅 (url,)
            url = None
            for arg in args:
                s = str(arg)
                if s.startswith('file:///') or s.startswith('file://') or s.startswith('shell:') or '::' in s:
                    url = s
                    break
            if url is None and args:
                url = str(args[-1])
            if not url:
                return
            local_path = None
            if url.startswith('file:'):
                # 统一解析 file: URL，正确还原本地盘符路径与 UNC 网络路径。
                # UNC 路径经 _on_nav_complete 编码后形如 file://///server/share，
                # 直接剥离固定前缀会丢失一个反斜杠，破坏 UNC 前缀，故按前导斜杠数判定。
                from urllib.parse import unquote
                rest = unquote(url[5:]).replace('/', '\\')
                stripped = rest.lstrip('\\')
                if len(stripped) >= 2 and stripped[1] == ':':
                    # 本地盘符路径，如 C:\...
                    local_path = self._normalize_local_path(stripped)
                elif stripped:
                    # UNC 网络路径，如 \\server\share\...
                    local_path = self._normalize_local_path('\\\\' + stripped)
            # 处理 CLSID 路径（如 ::{20D04FE0-...}\X:\）：尝试提取盘符路径
            if local_path is None and '::' in url:
                import re
                drive_match = re.search(r'([A-Za-z]:\\[^"]*)', url.replace('/', '\\'))
                if not drive_match:
                    drive_match = re.search(r'([A-Za-z]:/[^"]*)', url)
                if drive_match:
                    local_path = self._normalize_local_path(drive_match.group(1))
                    debug_print(f"[NavigateComplete2] Extracted drive path from CLSID URL: {local_path}")
            if local_path is not None:
                current = self._normalize_local_path(getattr(self, 'current_path', ''))
                # 抑制窗口恢复后 IEB 树面板自动展开导致的虚假导航
                # （树面板可能自动展开到之前记忆的子目录，导致 NavigateComplete2 连续触发错误路径）
                restore_guard = getattr(self, '_restore_guard_until', 0)
                if (restore_guard and time.monotonic() < restore_guard and
                        local_path != current and current and
                        not current.startswith('shell:')):
                    # 虚假导航：忽略并强制 IEB 回到正确路径
                    debug_print(f"[NavigateComplete2] Suppressed spurious post-restore nav: {local_path}")
                    if hasattr(self, 'explorer') and hasattr(self.explorer, '_navigate'):
                        self._restore_guard_until = 0  # 防止无限循环
                        self.explorer._navigate(current)
                    return
                # 无条件更新路径栏（即使路径未变也刷新显示）
                if hasattr(self, 'path_bar'):
                    self.path_bar.set_path(local_path)
                # 导航真正完成：解除慢盘异步导航的“导航中”锁定并隐藏 loading
                self._clear_async_nav_lock()
                # 延迟安装/更新 IExplorerBrowser 双击钩子（SysListView32 在首次导航后才创建）
                QTimer.singleShot(200, self._install_listview_dblclick_hook)
                if local_path and local_path != current:
                    self.current_path = local_path
                    self.update_tab_title()
                    self._schedule_status_update(track_selection=True)
                    if not getattr(self, '_navigating_programmatically', False) and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    if getattr(self, '_navigating_folder', False):
                        self._navigating_folder = False
                    # 记录导航时间戳，用于抑制 WH_MOUSE_LL 双击误判
                    self._last_nav_complete_time = time.monotonic()
                    if self.main_window and hasattr(self.main_window, 'update_chat_context'):
                        try:
                            if self.main_window.get_current_tab_widget() is self:
                                self.main_window.update_chat_context()
                        except Exception:
                            pass
                    debug_print(f"[NavigateComplete2] Path updated: {local_path}")
        except Exception as ex:
            debug_print(f"[NavigateComplete2] Error: {ex}")

    def event(self, e):
        # 捕获QAxWidget的NavigateComplete2事件（MetaCall备用通道，type=43）
        if e.type() == 43:  # QEvent.MetaCall
            if hasattr(e, 'arguments') and hasattr(e, 'signal'):
                if 'NavigateComplete2' in str(e.signal):
                    url = str(e.arguments[1])
                    # 检查是否为控制面板及其子目录，若是则用原生窗口打开
                    if self._is_control_panel_path(url):
                        try:
                            import subprocess
                            launch_detached(['explorer.exe', url])
                            show_toast(self, tr("已打开"), tr("控制面板已在新窗口打开"), level="info", duration=2000)
                        except Exception as ex:
                            show_toast(self, tr("错误"), tr("无法打开控制面板: {}").format(ex), level="error")
                        # 关闭当前标签页（定位其所属标签组，左右分屏均可）
                        mw = self.main_window
                        if mw and hasattr(mw, '_all_groups'):
                            for cand_tw, cand_cs in mw._all_groups():
                                if cand_cs is not None:
                                    j = cand_cs.indexOf(self)
                                    if j != -1:
                                        mw.close_tab(j, target_tabwidget=cand_tw)
                                        break
                        elif mw and hasattr(mw, 'tab_widget'):
                            idx = mw.tab_widget.indexOf(self)
                            if idx != -1:
                                mw.close_tab(idx)
                        return True
                    if url.startswith('file:///'):
                        from urllib.parse import unquote
                        local_path = unquote(url[8:])
                        if os.name == 'nt' and local_path.startswith('/'):
                            local_path = local_path[1:]
                        local_path = self._normalize_local_path(local_path)
                    elif url.startswith('file://') and not url.startswith('file:///'):
                        from urllib.parse import unquote
                        local_path = self._normalize_local_path('\\\\' + unquote(url[7:]).replace('/', '\\'))
                    else:
                        local_path = None
                    if local_path is not None:
                        current = self._normalize_local_path(getattr(self, 'current_path', ''))
                        if local_path != current:
                            self.current_path = local_path
                        if hasattr(self, 'path_bar'):
                            self.path_bar.set_path(local_path)
                        self._schedule_status_update(track_selection=True)
                        # 导航完成，清除导航标志（hit-test在双击时已完成判断，100ms后清除即可）
                        if hasattr(self, '_navigating_folder') and self._navigating_folder:
                            from PyQt5.QtCore import QTimer
                            QTimer.singleShot(100, lambda: setattr(self, '_navigating_folder', False))
                        if self.main_window and hasattr(self.main_window, 'update_chat_context'):
                            try:
                                if self.main_window.get_current_tab_widget() is self:
                                    self.main_window.update_chat_context()
                            except Exception:
                                pass
        return super().event(e)

    def open_tortoisegit_log(self):
        """打开 TortoiseGit 日志查看器"""
        try:
            current_path = self.current_path
            if not current_path or not os.path.exists(current_path):
                show_toast(self, tr("提示"), tr("当前路径无效"), level="warning")
                return

            repo_root = self._find_git_root(current_path)
            if not repo_root:
                show_toast(self, tr("提示"), tr("当前目录不是 Git 仓库，未找到 .git"), level="warning")
                return
            
            # TortoiseGit 命令行：TortoiseGitProc.exe /command:log /path:"路径"
            # 尝试找到 TortoiseGitProc.exe
            tortoisegit_paths = [
                r"C:\Program Files\TortoiseGit\bin\TortoiseGitProc.exe",
                r"C:\Program Files (x86)\TortoiseGit\bin\TortoiseGitProc.exe",
            ]
            
            tortoisegit_exe = None
            for path in tortoisegit_paths:
                if os.path.exists(path):
                    tortoisegit_exe = path
                    break
            
            if not tortoisegit_exe:
                show_toast(
                    self,
                    tr("提示"),
                    tr("未找到 TortoiseGit，请确认已安装 TortoiseGit\n下载地址: https://tortoisegit.org/download/"),
                    level="warning",
                )
                return
            
            # 启动 TortoiseGit Log
            launch_detached_async([tortoisegit_exe, '/command:log', f'/path:{repo_root}'])
            debug_print(f"[TortoiseGit] Opened log for: {repo_root}")
            
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开 TortoiseGit Log: {}").format(e), level="error")
            debug_print(f"[TortoiseGit] Failed to open log: {e}")
    
    def open_tortoisegit_commit(self):
        """打开 TortoiseGit 提交窗口"""
        try:
            current_path = self.current_path
            if not current_path or not os.path.exists(current_path):
                show_toast(self, tr("提示"), tr("当前路径无效"), level="warning")
                return

            repo_root = self._find_git_root(current_path)
            if not repo_root:
                show_toast(self, tr("提示"), tr("当前目录不是 Git 仓库，未找到 .git"), level="warning")
                return
            
            # TortoiseGit 命令行：TortoiseGitProc.exe /command:commit /path:"路径"
            tortoisegit_paths = [
                r"C:\Program Files\TortoiseGit\bin\TortoiseGitProc.exe",
                r"C:\Program Files (x86)\TortoiseGit\bin\TortoiseGitProc.exe",
            ]
            
            tortoisegit_exe = None
            for path in tortoisegit_paths:
                if os.path.exists(path):
                    tortoisegit_exe = path
                    break
            
            if not tortoisegit_exe:
                show_toast(
                    self,
                    tr("提示"),
                    tr("未找到 TortoiseGit，请确认已安装 TortoiseGit\n下载地址: https://tortoisegit.org/download/"),
                    level="warning",
                )
                return
            
            # 启动 TortoiseGit Commit
            launch_detached_async([tortoisegit_exe, '/command:commit', f'/path:{repo_root}'])
            debug_print(f"[TortoiseGit] Opened commit for: {repo_root}")
            
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开 TortoiseGit Commit: {}").format(e), level="error")
            debug_print(f"[TortoiseGit] Failed to open commit: {e}")

    def _find_git_root(self, start_path):
        """向上查找包含 .git 的目录，找到则返回仓库根路径，否则返回 None"""
        if not start_path:
            return None
        path = os.path.abspath(start_path)
        while True:
            git_marker = os.path.join(path, '.git')
            if os.path.isdir(git_marker):
                return path
            if os.path.isfile(git_marker):
                try:
                    with open(git_marker, 'r', encoding='utf-8', errors='ignore') as f:
                        line = f.readline().strip()
                    if line.lower().startswith('gitdir:'):
                        gitdir_path = line[7:].strip()
                        if not os.path.isabs(gitdir_path):
                            gitdir_path = os.path.abspath(os.path.join(path, gitdir_path))
                        if os.path.exists(gitdir_path):
                            return path
                except Exception:
                    pass
            parent = os.path.dirname(path)
            if parent == path:
                break
            path = parent
        return None

    def _request_git_status_async(self, dir_path):
        """异步请求 Git 状态（不阻塞 UI 线程）"""
        if not dir_path or dir_path.startswith('shell:') or '::' in dir_path:
            return
        # 慢盘（网络/UNC/映射盘）：_find_git_root 向上逐级 os.path.isdir/isfile 探测，
        # 在挂起的网络路径上会同步阻塞 UI 线程导致卡死。网络盘一般非 Git 仓库，直接跳过。
        if self._is_slow_path(dir_path):
            self._git_status_cache = {'path': dir_path, 'result': None,
                                      'ts_ms': int(time.time() * 1000)}
            return
        # 防止重复请求同一路径
        if getattr(self, '_git_status_pending_path', None) == dir_path:
            return
        cache = getattr(self, '_git_status_cache', None)
        now_ms = int(time.time() * 1000)
        if cache and cache.get('path') == dir_path and (now_ms - cache.get('ts_ms', 0)) < 5000:
            return  # 缓存仍有效，无需重新查询
        repo_root = self._find_git_root(dir_path)
        if not repo_root:
            self._git_status_cache = {'path': dir_path, 'result': None, 'ts_ms': now_ms}
            return
        git_exe = 'git.exe'
        git_root = find_git_install_root()
        if git_root:
            candidate = os.path.join(git_root, 'cmd', 'git.exe')
            if os.path.isfile(candidate):
                git_exe = candidate
        self._git_status_pending_path = dir_path
        worker = GitStatusWorker(dir_path, repo_root, git_exe, parent=self)
        worker.finished.connect(self._on_git_status_finished)
        worker.finished.connect(worker.deleteLater)
        self._git_status_worker = worker  # prevent GC
        worker.start()

    def _on_git_status_finished(self, dir_path, repo_root, summary):
        """Git 状态查询完成回调（主线程）"""
        now_ms = int(time.time() * 1000)
        self._git_status_cache = {'path': dir_path, 'result': summary, 'ts_ms': now_ms}
        if getattr(self, '_git_status_pending_path', None) == dir_path:
            self._git_status_pending_path = None
        # 刷新状态栏（仅当当前路径匹配时）
        if getattr(self, 'current_path', None) == dir_path:
            self.update_explorer_status()

    def on_path_bar_changed(self, path):
        """处理面包屑路径栏的路径变化，支持特殊shell路径自动跳转"""
        import os
        path = path.strip()

        # 处理 file: URL（从邮件/浏览器/SharePoint等复制的链接）
        # 支持 file:///C:/path、file://server/share、file:\\server\share、file:\server\share 等格式
        if path.lower().startswith('file:'):
            from urllib.parse import unquote
            raw = path[5:]  # 去掉 'file:'
            # 去掉所有前导斜杠和反斜杠
            stripped = raw.lstrip('/\\')
            stripped = unquote(stripped).replace('/', '\\')
            # 判断是本地路径(有盘符)还是UNC网络路径
            if len(stripped) >= 2 and stripped[1] == ':':
                path = stripped  # 本地路径 C:\path
            else:
                path = '\\\\' + stripped  # UNC路径 \\server\share
            debug_print(f"[PathBar] file: URL converted to: {path}")

        lower_path = path.lower()
        debug_print(f"[PathBar] on_path_bar_changed: '{path}'")

        if lower_path in ('terminal', 'term'):
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    preferred_tool = 'cmd'
                    if self.main_window and hasattr(self.main_window, 'get_preferred_terminal_tool'):
                        preferred_tool = self.main_window.get_preferred_terminal_tool()
                    launch_shell_tool(preferred_tool, current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开默认终端"), level="error")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开默认终端: {}").format(e), level="error")
            return

        # 处理cmd命令
        if lower_path == 'cmd':
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    launch_shell_tool('cmd', current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开命令行"), level="error")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开命令行: {}").format(e), level="error")
            return

        if lower_path in ('powershell', 'ps', 'pwsh'):
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    launch_shell_tool('powershell', current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开 PowerShell"), level="error")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开 PowerShell: {}").format(e), level="error")
            return

        if lower_path in ('gitbash', 'git-bash', 'bash'):
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    launch_shell_tool('git-bash', current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开 Git Bash"), level="error")
            except FileNotFoundError as e:
                show_toast(self, tr("提示"), str(e), level="warning")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开 Git Bash: {}").format(e), level="error")
            return

        # 支持特殊shell路径映射
        special_map = {
            tr('回收站'): 'shell:RecycleBinFolder',
            tr('此电脑'): 'shell:MyComputerFolder',
            tr('我的电脑'): 'shell:MyComputerFolder',
            '桌面': 'shell:Desktop',
            tr('网络'): 'shell:NetworkPlacesFolder',
            tr('启动项'): 'shell:Startup',
            tr('开机启动项'): 'shell:Startup',
            tr('启动文件夹'): 'shell:Startup',
            'Startup': 'shell:Startup',
            'OneDrive': 'shell:OneDrive',
            'onedrive': 'shell:OneDrive',
        }
        # shell:Startup 等特殊路径，自动解析为真实系统路径
        shell_path_map = {
            'shell:Startup': lambda: os.path.join(os.environ.get('APPDATA', ''), r'Microsoft\Windows\Start Menu\Programs\Startup'),
            'shell:OneDrive': lambda: os.environ.get('OneDrive', ''),
        }

        if path in special_map:
            shell_path = special_map[path]
            # 如果是 shell:Startup，自动跳转到真实路径
            if shell_path in shell_path_map:
                real_path = shell_path_map[shell_path]()
                if os.path.exists(real_path):
                    self.navigate_to(real_path)
                else:
                    show_toast(self, tr("路径错误"), tr("启动文件夹不存在: {}").format(real_path), level="warning")
                    if hasattr(self, 'current_path') and self.current_path:
                        self.path_bar.set_path(self.current_path)
                return
            else:
                self.navigate_to(shell_path, is_shell=True)
                return

        # 允许直接输入 shell:XXX 跳转
        if path.lower().startswith('shell:'):
            # shell:Startup 也做特殊处理
            if path.lower() == 'shell:startup' and 'shell:Startup' in shell_path_map:
                real_path = shell_path_map['shell:Startup']()
                if os.path.exists(real_path):
                    self.navigate_to(real_path)
                else:
                    show_toast(self, tr("路径错误"), tr("启动文件夹不存在: {}").format(real_path), level="warning")
                    if hasattr(self, 'current_path') and self.current_path:
                        self.path_bar.set_path(self.current_path)
                return
            # shell:OneDrive 解析为真实路径
            if path.lower() == 'shell:onedrive' and 'shell:OneDrive' in shell_path_map:
                real_path = shell_path_map['shell:OneDrive']()
                if real_path and os.path.exists(real_path):
                    self.navigate_to(real_path)
                else:
                    show_toast(self, tr("路径错误"), tr("未找到 OneDrive 文件夹"), level="warning")
                    if hasattr(self, 'current_path') and self.current_path:
                        self.path_bar.set_path(self.current_path)
                return
            self.navigate_to(path, is_shell=True)
            return

        # 路径与当前路径相同时（如 resizeEvent 误触发 pathChanged），跳过导航，防止循环刷新
        current_norm = getattr(self, 'current_path', '').replace('/', '\\').lower().rstrip('\\')
        path_norm = path.replace('/', '\\').lower().rstrip('\\')
        if path_norm == current_norm:
            debug_print(f"[PathBar] on_path_bar_changed: same as current_path, skip navigate")
            return

        # 尝试中英文目录互转
        path2 = translate_common_path(path)
        # 慢盘（网络/UNC/映射盘/OneDrive）：os.path.exists 在挂起路径上会阻塞 UI 线程，
        # 直接交给 navigate_to（其内部对慢盘走后台异步解析，不做同步文件系统探测）。
        if self._is_slow_path(path2) or path2.startswith('\\\\') or path2.startswith('//'):
            self.navigate_to(path2)
        elif os.path.exists(path2):
            self.navigate_to(path2)
        else:
            show_toast(self, tr("路径错误"), tr("路径不存在: {}").format(path2), level="warning")
            if hasattr(self, 'current_path') and self.current_path:
                self.path_bar.set_path(self.current_path)

    def explorer_mouse_press(self, event):
        # 在鼠标按下时记录当时的选中项数量和鼠标位置点击测试结果
        try:
            cnt = self._get_selected_count_safe()
            self._selected_before_click = int(cnt) if cnt is not None else None
            
            # 使用原生点击测试来判断是否点击在项目上（更准确）
            if HAS_PYWIN:
                try:
                    from PyQt5.QtGui import QCursor
                    gx = QCursor.pos().x()
                    gy = QCursor.pos().y()
                    self._clicked_on_item = self._native_listview_hit_test(gx, gy)
                except Exception:
                    self._clicked_on_item = False
            else:
                self._clicked_on_item = False
                
        except Exception:
            self._selected_before_click = None
            self._clicked_on_item = False
        # 继续默认处理（不阻止控件行为）
        # 直接返回 None — 不尝试调用 ActiveX 的原始处理（事件仍会被控件处理）
        return None

    def _get_selected_count_safe(self):
        """安全地获取当前选中项数量，避免触发 ActiveX 属性不存在的警告"""
        try:
            # IExplorerBrowser 模式：使用 IFolderView COM 接口
            if isinstance(self.explorer, IExplorerBrowserWidget) and getattr(self.explorer, '_browser', None):
                return self._get_ieb_selection_count()

            # 优先通过 Document 接口获取 SelectedItems（避免 WebBrowser 直接调用警告）
            doc = None
            try:
                doc = self.explorer.querySubObject('Document') if hasattr(self, 'explorer') else None
            except Exception:
                doc = None

            sel = None
            if doc:
                try:
                    sel = doc.querySubObject('SelectedItems()')
                except Exception:
                    sel = None

            if sel is None:
                return None

            count = None
            try:
                if hasattr(sel, 'property'):
                    count = sel.property('Count')
            except Exception:
                count = None

            if count is None:
                try:
                    count = len(sel)
                except Exception:
                    count = None

            return int(count) if count is not None else None
        except Exception:
            return None

    def _get_ieb_selection_count(self):
        """通过 IFolderView COM 接口获取 IExplorerBrowser 的选中项数量。
        先获取 IShellView（已验证可用），再 QueryInterface 获取 IFolderView。"""
        try:
            from comtypes import GUID as _GUID
            # 先获取 IShellView（已验证此路径可靠工作）
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_sv = self.explorer._browser.GetCurrentView(ctypes.byref(iid_sv))
            sv_ptr = int(ppv_sv) if ppv_sv else 0
            if not sv_ptr:
                return None
            try:
                _vp_size = ctypes.sizeof(ctypes.c_void_p)
                # 通过 IShellView 的 QueryInterface 获取 IFolderView
                iid_fv = _GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                qi_addr = ctypes.c_void_p.from_address(vtable_ptr).value  # QI at index 0
                _QI = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                         ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
                qi_fn = _QI(qi_addr)
                fv_ptr = ctypes.c_void_p(0)
                hr_qi = qi_fn(sv_ptr, ctypes.byref(iid_fv), ctypes.byref(fv_ptr))
                if hr_qi != 0 or not fv_ptr.value:
                    return None
                try:
                    # IFolderView vtable (inherits IUnknown directly):
                    # QI(0), AddRef(1), Release(2), GetCurrentViewMode(3),
                    # SetCurrentViewMode(4), GetFolder(5), Item(6), ItemCount(7)
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value
                    fn_addr = ctypes.c_void_p.from_address(fv_vtable + 7 * _vp_size).value
                    _IC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                             ctypes.c_uint, ctypes.POINTER(ctypes.c_int))
                    item_count_fn = _IC(fn_addr)
                    count = ctypes.c_int(0)
                    hr = item_count_fn(fv_ptr.value, 0x1, ctypes.byref(count))  # SVGIO_SELECTION=1
                    if hr == 0:
                        return count.value
                finally:
                    # Release IFolderView
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value
                    rel_addr = ctypes.c_void_p.from_address(fv_vtable + 2 * _vp_size).value
                    _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                    _REL(rel_addr)(fv_ptr.value)
            finally:
                # Release IShellView
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                rel_addr = ctypes.c_void_p.from_address(vtable_ptr + 2 * _vp_size).value
                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _REL(rel_addr)(sv_ptr)
        except Exception as e:
            debug_print(f"[IEB] _get_ieb_selection_count failed: {e}")
        return None

    def _get_ieb_selected_paths(self):
        """通过 IFolderView::Items(SVGIO_SELECTION) 获取 IExplorerBrowser 选中文件路径列表。"""
        paths = []
        try:
            from comtypes import GUID as _GUID
            _vp_size = ctypes.sizeof(ctypes.c_void_p)
            # 获取 IShellView
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_sv = self.explorer._browser.GetCurrentView(ctypes.byref(iid_sv))
            sv_ptr = int(ppv_sv) if ppv_sv else 0
            if not sv_ptr:
                return paths
            try:
                # QI for IFolderView
                iid_fv = _GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                qi_addr = ctypes.c_void_p.from_address(vtable_ptr).value
                _QI = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                         ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
                qi_fn = _QI(qi_addr)
                fv_ptr = ctypes.c_void_p(0)
                hr_qi = qi_fn(sv_ptr, ctypes.byref(iid_fv), ctypes.byref(fv_ptr))
                if hr_qi != 0 or not fv_ptr.value:
                    return paths
                try:
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value
                    # IFolderView::Items at vtable index 8
                    # HRESULT Items(UINT uFlags, REFIID riid, void** ppv)
                    iid_sia = _GUID("{B63EA76D-1F85-456F-A19C-48159EFA858B}")  # IShellItemArray
                    fn_addr = ctypes.c_void_p.from_address(fv_vtable + 8 * _vp_size).value
                    _ITEMS = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                ctypes.c_uint, ctypes.POINTER(_GUID),
                                                ctypes.POINTER(ctypes.c_void_p))
                    items_fn = _ITEMS(fn_addr)
                    sia_ptr = ctypes.c_void_p(0)
                    hr = items_fn(fv_ptr.value, 0x1, ctypes.byref(iid_sia), ctypes.byref(sia_ptr))
                    if hr != 0 or not sia_ptr.value:
                        return paths
                    try:
                        # IShellItemArray::GetCount at vtable index 7
                        sia_vtable = ctypes.c_void_p.from_address(sia_ptr.value).value
                        gc_addr = ctypes.c_void_p.from_address(sia_vtable + 7 * _vp_size).value
                        _GC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                  ctypes.POINTER(ctypes.c_uint))
                        gc_fn = _GC(gc_addr)
                        count = ctypes.c_uint(0)
                        hr = gc_fn(sia_ptr.value, ctypes.byref(count))
                        if hr != 0 or count.value == 0:
                            return paths
                        # IShellItemArray::GetItemAt at vtable index 8
                        gia_addr = ctypes.c_void_p.from_address(sia_vtable + 8 * _vp_size).value
                        _GIA = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                   ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p))
                        gia_fn = _GIA(gia_addr)
                        for i in range(count.value):
                            si_ptr = ctypes.c_void_p(0)
                            hr = gia_fn(sia_ptr.value, i, ctypes.byref(si_ptr))
                            if hr != 0 or not si_ptr.value:
                                continue
                            try:
                                # IShellItem::GetDisplayName at vtable index 5
                                # SIGDN_FILESYSPATH = 0x80058000
                                si_vtable = ctypes.c_void_p.from_address(si_ptr.value).value
                                gdn_addr = ctypes.c_void_p.from_address(si_vtable + 5 * _vp_size).value
                                _GDN = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                          ctypes.c_uint, ctypes.POINTER(ctypes.c_wchar_p))
                                gdn_fn = _GDN(gdn_addr)
                                name_ptr = ctypes.c_wchar_p()
                                hr = gdn_fn(si_ptr.value, 0x80058000, ctypes.byref(name_ptr))
                                if hr == 0 and name_ptr.value:
                                    paths.append(name_ptr.value)
                                    ctypes.windll.ole32.CoTaskMemFree(name_ptr)
                            finally:
                                # Release IShellItem
                                si_vt = ctypes.c_void_p.from_address(si_ptr.value).value
                                rel_addr = ctypes.c_void_p.from_address(si_vt + 2 * _vp_size).value
                                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                                _REL(rel_addr)(si_ptr.value)
                    finally:
                        # Release IShellItemArray
                        sia_vt = ctypes.c_void_p.from_address(sia_ptr.value).value
                        rel_addr = ctypes.c_void_p.from_address(sia_vt + 2 * _vp_size).value
                        _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                        _REL(rel_addr)(sia_ptr.value)
                finally:
                    # Release IFolderView
                    fv_vt = ctypes.c_void_p.from_address(fv_ptr.value).value
                    rel_addr = ctypes.c_void_p.from_address(fv_vt + 2 * _vp_size).value
                    _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                    _REL(rel_addr)(fv_ptr.value)
            finally:
                # Release IShellView
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                rel_addr = ctypes.c_void_p.from_address(vtable_ptr + 2 * _vp_size).value
                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _REL(rel_addr)(sv_ptr)
        except Exception as e:
            debug_print(f"[IEB] _get_ieb_selected_paths failed: {e}")
        return paths

    def _select_ieb_item_by_name(self, filename):
        """通过 IFolderView 直接选中 IExplorerBrowser 当前视图中的项目。"""
        try:
            if not filename:
                return False
            if not isinstance(self.explorer, IExplorerBrowserWidget) or not getattr(self.explorer, '_browser', None):
                return False

            from comtypes import GUID as _GUID

            _vp_size = ctypes.sizeof(ctypes.c_void_p)
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_sv = self.explorer._browser.GetCurrentView(ctypes.byref(iid_sv))
            sv_ptr = int(ppv_sv) if ppv_sv else 0
            if not sv_ptr:
                debug_print("[IEB Select] IShellView unavailable")
                return False

            def _release_interface(ptr_value):
                if not ptr_value:
                    return
                vt = ctypes.c_void_p.from_address(ptr_value).value
                rel_addr = ctypes.c_void_p.from_address(vt + 2 * _vp_size).value
                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _REL(rel_addr)(ptr_value)

            def _get_display_name(si_ptr_value, sigdn):
                si_vtable = ctypes.c_void_p.from_address(si_ptr_value).value
                gdn_addr = ctypes.c_void_p.from_address(si_vtable + 5 * _vp_size).value
                _GDN = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                          ctypes.c_uint, ctypes.POINTER(ctypes.c_wchar_p))
                gdn_fn = _GDN(gdn_addr)
                name_ptr = ctypes.c_wchar_p()
                hr = gdn_fn(si_ptr_value, sigdn, ctypes.byref(name_ptr))
                if hr == 0 and name_ptr.value:
                    value = name_ptr.value
                    ctypes.windll.ole32.CoTaskMemFree(name_ptr)
                    return value
                return None

            try:
                iid_fv = _GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                qi_addr = ctypes.c_void_p.from_address(vtable_ptr).value
                _QI = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                         ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
                qi_fn = _QI(qi_addr)
                fv_ptr = ctypes.c_void_p(0)
                hr_qi = qi_fn(sv_ptr, ctypes.byref(iid_fv), ctypes.byref(fv_ptr))
                if hr_qi != 0 or not fv_ptr.value:
                    debug_print(f"[IEB Select] QI IFolderView failed: hr=0x{hr_qi & 0xFFFFFFFF:08X}")
                    return False

                try:
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value

                    iid_sia = _GUID("{B63EA76D-1F85-456F-A19C-48159EFA858B}")
                    items_addr = ctypes.c_void_p.from_address(fv_vtable + 8 * _vp_size).value
                    _ITEMS = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                ctypes.c_uint, ctypes.POINTER(_GUID),
                                                ctypes.POINTER(ctypes.c_void_p))
                    items_fn = _ITEMS(items_addr)
                    sia_ptr = ctypes.c_void_p(0)
                    # SVGIO_ALLVIEW(0x2) | SVGIO_FLAG_VIEWORDER(0x80000000)：
                    # 必须带 VIEWORDER，使枚举顺序与 IFolderView::SelectItem(iItem) 的
                    # 视图索引一致；否则匹配到的索引会指向显示顺序中的另一个文件，
                    # 表现为“定位到了但选中了错误的文件”。
                    SVGIO_ALLVIEW_VIEWORDER = 0x2 | 0x80000000
                    hr_items = items_fn(fv_ptr.value, SVGIO_ALLVIEW_VIEWORDER, ctypes.byref(iid_sia), ctypes.byref(sia_ptr))
                    if hr_items != 0 or not sia_ptr.value:
                        debug_print(f"[IEB Select] Items(SVGIO_ALLVIEW) failed: hr=0x{hr_items & 0xFFFFFFFF:08X}")
                        return False

                    try:
                        sia_vtable = ctypes.c_void_p.from_address(sia_ptr.value).value
                        gc_addr = ctypes.c_void_p.from_address(sia_vtable + 7 * _vp_size).value
                        _GC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                  ctypes.POINTER(ctypes.c_uint))
                        gc_fn = _GC(gc_addr)
                        count = ctypes.c_uint(0)
                        hr_count = gc_fn(sia_ptr.value, ctypes.byref(count))
                        if hr_count != 0 or count.value == 0:
                            debug_print(f"[IEB Select] View items unavailable: hr=0x{hr_count & 0xFFFFFFFF:08X}, count={count.value}")
                            return False

                        gia_addr = ctypes.c_void_p.from_address(sia_vtable + 8 * _vp_size).value
                        _GIA = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                   ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p))
                        gia_fn = _GIA(gia_addr)

                        target_index = -1
                        filename_cf = filename.casefold()
                        for i in range(count.value):
                            si_ptr = ctypes.c_void_p(0)
                            hr_item = gia_fn(sia_ptr.value, i, ctypes.byref(si_ptr))
                            if hr_item != 0 or not si_ptr.value:
                                continue
                            try:
                                item_name = _get_display_name(si_ptr.value, 0)
                                if not item_name:
                                    item_path = _get_display_name(si_ptr.value, 0x80058000)
                                    item_name = os.path.basename(item_path) if item_path else None
                                if item_name and item_name.casefold() == filename_cf:
                                    target_index = i
                                    break
                            finally:
                                _release_interface(si_ptr.value)

                        if target_index < 0:
                            debug_print(f"[IEB Select] Target not found in current view: {filename}")
                            return False

                        select_addr = ctypes.c_void_p.from_address(fv_vtable + 15 * _vp_size).value
                        _SELECT = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                     ctypes.c_int, ctypes.c_uint)
                        select_fn = _SELECT(select_addr)
                        svsi_flags = 0x1 | 0x4 | 0x8 | 0x10
                        hr_select = select_fn(fv_ptr.value, target_index, svsi_flags)
                        if hr_select != 0:
                            debug_print(f"[IEB Select] SelectItem failed: hr=0x{hr_select & 0xFFFFFFFF:08X}, index={target_index}")
                            return False

                        self.activateWindow()
                        self.raise_()
                        if hasattr(self, 'explorer') and self.explorer:
                            self.explorer.setFocus(Qt.OtherFocusReason)

                        debug_print(f"[IEB Select] Successfully selected item via IFolderView: {filename}")
                        self._suppress_auto_refresh = False
                        self._arm_selection_guard()
                        self.start_path_sync_timer(duration_ms=8000)
                        return True
                    finally:
                        _release_interface(sia_ptr.value)
                finally:
                    _release_interface(fv_ptr.value)
            finally:
                _release_interface(sv_ptr)
        except Exception as e:
            debug_print(f"[IEB Select] Failed to select item '{filename}': {e}")
        return False

    # --- Windows native helpers for listview hit-testing ---
    def _find_syslistview_hwnd(self):
        if not HAS_PYWIN:
            return None
        try:
            parent = int(self.explorer.winId())
        except Exception:
            return None

        def find_in_tree(hwnd):
            try:
                cls = win32gui.GetClassName(hwnd)
            except Exception:
                return None
            if cls == 'SysListView32':
                return hwnd
            result = None
            try:
                def cb(h, lparam):
                    nonlocal result
                    if result:
                        return False
                    found = find_in_tree(h)
                    if found:
                        result = found
                        return False
                    return True
                win32gui.EnumChildWindows(hwnd, cb, None)
            except Exception:
                return None
            return result

        return find_in_tree(parent)

    def _install_listview_dblclick_hook(self):
        """Install a single process-wide WH_MOUSE_LL hook (singleton) for
        double-click detection across all IExplorerBrowser tabs.

        Only one hook exists; it dispatches to the current active tab.
        """
        if not HAS_PYWIN:
            return
        if not isinstance(getattr(self, 'explorer', None), IExplorerBrowserWidget):
            return
        if not getattr(self.explorer, '_init_ok', False):
            return

        # Singleton: if a global hook already exists, skip
        if getattr(FileExplorerTab, '_global_mouse_hook_handle', None):
            return

        _self_ref = self  # only used to locate main_window

        class _POINT(ctypes.Structure):
            _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]

        class _MSLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [
                ('pt',          _POINT),
                ('mouseData',   ctypes.c_ulong),
                ('flags',       ctypes.c_ulong),
                ('time',        ctypes.c_ulong),
                ('dwExtraInfo', ctypes.c_size_t),
            ]

        _HOOKPROC = ctypes.WINFUNCTYPE(
            ctypes.c_ssize_t,
            ctypes.c_int,
            ctypes.wintypes.WPARAM,
            ctypes.wintypes.LPARAM,
        )

        WM_LBUTTONDOWN = 0x0201
        WM_MOUSEMOVE   = 0x0200
        WM_RBUTTONDOWN = 0x0204
        WM_RBUTTONUP   = 0x0205
        _GESTURE_MARKER    = 0x47455354  # 'GEST'：标记本程序合成的右键事件，避免被本钩子再次拦截
        _GESTURE_THRESHOLD = 30          # 触发鼠标手势的最小位移（像素）
        _state = {'t': 0, 'x': -9999, 'y': -9999}
        # 鼠标手势（类似 Mouse Gestures）：按住右键画线，支持多笔画序列
        #   ←后退 / →前进 / ↓关闭标签 / ↑新建标签
        #   ↑↓刷新 / ↓↑上级目录 / ↓→恢复关闭的标签页
        _gesture = {'active': False, 'sx': 0, 'sy': 0, 'cx': 0, 'cy': 0,
                    'ax': 0, 'ay': 0, 'seq': []}
        _STROKE_THRESHOLD = 36  # 单笔画方向判定的最小位移（像素）
        _ARROW = {'left': '←', 'right': '→', 'up': '↑', 'down': '↓'}
        # 手势序列 → (动作键, 动作名)
        _GESTURE_TABLE = {
            ('left',):           ('back',    tr('后退')),
            ('right',):          ('forward', tr('前进')),
            ('down',):           ('close',   tr('关闭标签页')),
            ('up',):             ('new',     tr('新建标签页')),
            ('up', 'down'):      ('refresh', tr('刷新')),
            ('down', 'up'):      ('up_dir',  tr('上级目录')),
            ('down', 'right'):   ('reopen',  tr('恢复关闭的标签页')),
        }

        def _seq_arrows(seq):
            return ''.join(_ARROW.get(d, '') for d in seq)

        def _lookup_gesture(seq):
            """返回 (动作键, 动作名)；未匹配返回 (None, '')。"""
            return _GESTURE_TABLE.get(tuple(seq), (None, ''))

        def _get_current_tab():
            """获取当前活动的浏览面板（含分屏面板，由最近一次鼠标按下位置决定）"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                if mw and hasattr(mw, 'get_active_pane'):
                    tab = mw.get_active_pane()
                    if tab and isinstance(getattr(tab, 'explorer', None), IExplorerBrowserWidget):
                        return tab
            except Exception:
                pass
            return None

        def _gestures_enabled():
            """鼠标手势是否启用（config.json: enable_mouse_gestures，默认开启）"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                if mw and hasattr(mw, 'config'):
                    return bool(mw.config.get('enable_mouse_gestures', True))
            except Exception:
                pass
            return True

        def _cursor_over_current_explorer(px, py):
            """光标是否位于某个浏览面板（当前标签或分屏面板）的资源管理器区域内，且本程序窗口为前台。

            仅在满足条件时才接管右键，避免影响其他程序或非文件区域的右键菜单。
            """
            try:
                mw = getattr(_self_ref, 'main_window', None)
                if mw is None:
                    return False
                # 仅当本程序为前台窗口时接管右键
                try:
                    fg = ctypes.windll.user32.GetForegroundWindow()
                    if int(fg) != int(mw.winId()):
                        return False
                except Exception:
                    pass
                return mw.pane_at_global_pos(px, py) is not None
            except Exception:
                return False

        def _run_gesture_action(action):
            """在主线程执行手势对应的操作"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                tab = _get_current_tab()
                if tab is not None:
                    mw = getattr(tab, 'main_window', None) or mw
                if not mw:
                    return
                if action == 'back':
                    mw.go_back_current_tab()
                elif action == 'forward':
                    mw.go_forward_current_tab()
                elif action == 'close':
                    mw.close_current_tab()
                elif action == 'new':
                    # 新建标签作用于手势所在的一侧（左/右分屏组）
                    mw.add_new_tab(target_tabwidget=mw.get_active_group_tabwidget())
                elif action == 'refresh':
                    mw.refresh_current_tab()
                elif action == 'up_dir':
                    mw.go_up_current_tab()
                elif action == 'reopen':
                    mw.reopen_closed_tab()
            except Exception as _e:
                debug_print(f"[Gesture] action error: {_e}")

        def _synth_right_click():
            """合成一次真实右键点击以弹出系统右键菜单（带标记，避免被本钩子再次拦截）"""
            try:
                MOUSEEVENTF_RIGHTDOWN = 0x0008
                MOUSEEVENTF_RIGHTUP   = 0x0010
                ctypes.windll.user32.mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, _GESTURE_MARKER)
                ctypes.windll.user32.mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, _GESTURE_MARKER)
            except Exception as _e:
                debug_print(f"[Gesture] synth right click error: {_e}")

        def _get_overlay():
            """获取/懒创建 MainWindow 上的手势覆盖层单例。"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                tab = _get_current_tab()
                if tab is not None:
                    mw = getattr(tab, 'main_window', None) or mw
                if mw is None:
                    return None
                ov = getattr(mw, '_gesture_overlay', None)
                if ov is None:
                    ov = GestureOverlay()
                    mw._gesture_overlay = ov
                return ov
            except Exception:
                return None

        def _hook_proc(nCode, wParam, lParam):
            # ── 鼠标手势处理（右键画线）─────────────────────────────────
            if nCode >= 0:
                try:
                    g_info = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                    if int(g_info.dwExtraInfo) == _GESTURE_MARKER:
                        pass  # 本程序合成的事件：直接放行
                    elif wParam == WM_RBUTTONDOWN and _gestures_enabled():
                        gx, gy = int(g_info.pt.x), int(g_info.pt.y)
                        # 手势起点处的面板即为本次手势目标（同时供后续键盘定位）
                        mw_g = getattr(_self_ref, 'main_window', None)
                        if mw_g is not None:
                            try:
                                mw_g.set_active_pane_from_global_pos(gx, gy)
                            except Exception:
                                pass
                        if _cursor_over_current_explorer(gx, gy):
                            _gesture['active'] = True
                            _gesture['sx'], _gesture['sy'] = gx, gy
                            _gesture['cx'], _gesture['cy'] = gx, gy
                            _gesture['ax'], _gesture['ay'] = gx, gy
                            _gesture['seq'] = []
                            ov = _get_overlay()
                            if ov is not None:
                                ov.start()
                                ov.add_point(gx, gy)
                            return 1  # 吞掉按下，松开时再决定执行手势或弹出菜单
                    elif wParam == WM_MOUSEMOVE and _gesture['active']:
                        nx, ny = int(g_info.pt.x), int(g_info.pt.y)
                        _gesture['cx'], _gesture['cy'] = nx, ny
                        # 多笔画方向识别：相对当前笔画锚点的位移超过阈值则提交一个方向
                        adx = nx - _gesture['ax']
                        ady = ny - _gesture['ay']
                        if max(abs(adx), abs(ady)) >= _STROKE_THRESHOLD:
                            if abs(adx) >= abs(ady):
                                d = 'right' if adx > 0 else 'left'
                            else:
                                d = 'down' if ady > 0 else 'up'
                            seq = _gesture['seq']
                            if not seq or seq[-1] != d:
                                seq.append(d)
                            _gesture['ax'], _gesture['ay'] = nx, ny
                        ov = _get_overlay()
                        if ov is not None:
                            ov.add_point(nx, ny)
                            _action, _label = _lookup_gesture(_gesture['seq'])
                            ov.set_hint(_seq_arrows(_gesture['seq']) if _label else '', _label)
                    elif wParam == WM_RBUTTONUP and _gesture['active']:
                        _gesture['active'] = False
                        ov = _get_overlay()
                        if ov is not None:
                            ov.finish()
                        action, label = _lookup_gesture(_gesture['seq'])
                        if action:
                            debug_print(f"[Gesture] {_seq_arrows(_gesture['seq'])} → {action}")
                            QTimer.singleShot(0, lambda a=action: _run_gesture_action(a))
                        else:
                            # 未识别为有效手势 → 合成真实右键以弹出菜单
                            QTimer.singleShot(0, _synth_right_click)
                        _gesture['seq'] = []
                        return 1  # 吞掉这次松开（已执行手势或即将合成右键）
                except Exception:
                    pass

            if nCode >= 0 and wParam == WM_LBUTTONDOWN:
                try:
                    info = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                    px = int(info.pt.x)
                    py = int(info.pt.y)
                    t  = int(info.time)

                    # 左键按下位置所在面板即为活动面板（供键盘快捷键/双击定位）
                    mw_l = getattr(_self_ref, 'main_window', None)
                    if mw_l is not None:
                        try:
                            mw_l.set_active_pane_from_global_pos(px, py)
                        except Exception:
                            pass

                    # 点击 explorer 区域内时，通知路径栏退出编辑模式
                    try:
                        from PyQt5.QtCore import QPoint
                        tab = _get_current_tab()
                        if tab:
                            pt_chk = tab.explorer.mapFromGlobal(QPoint(px, py))
                            if tab.explorer.rect().contains(pt_chk):
                                if hasattr(tab, 'path_bar') and getattr(tab.path_bar, '_in_edit', False):
                                    tab.path_bar.exit_edit_mode()
                                elif hasattr(tab, 'path_bar') and getattr(tab.path_bar, 'edit_mode', False):
                                    tab.path_bar.exit_edit_mode()
                    except Exception:
                        pass

                    dct = ctypes.windll.user32.GetDoubleClickTime()
                    dt  = (t - _state['t']) & 0xFFFFFFFF

                    is_dblclick = (
                        _state['t'] != 0 and
                        dt <= dct and
                        abs(px - _state['x']) <= 4 and
                        abs(py - _state['y']) <= 4
                    )

                    if is_dblclick:
                        _state['t'] = 0

                        def _check(_px=px, _py=py):
                            try:
                                tab = _get_current_tab()
                                if not tab:
                                    return

                                # Bounds check
                                from PyQt5.QtCore import QPoint
                                pt_local = tab.explorer.mapFromGlobal(QPoint(_px, _py))
                                if not tab.explorer.rect().contains(pt_local):
                                    return

                                # Navigation guard
                                last_nav = getattr(tab, '_last_nav_complete_time', 0)
                                if time.monotonic() - last_nav < 0.5:
                                    return

                                # LVM_HITTEST
                                lv = tab._find_syslistview_hwnd()
                                if lv:
                                    try:
                                        cpt = win32gui.ScreenToClient(lv, (_px, _py))

                                        class _PT(ctypes.Structure):
                                            _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]
                                        class _LVHI(ctypes.Structure):
                                            _fields_ = [('pt', _PT), ('flags', ctypes.c_uint),
                                                         ('iItem', ctypes.c_int), ('iSubItem', ctypes.c_int)]

                                        hi = _LVHI()
                                        hi.pt.x, hi.pt.y = int(cpt[0]), int(cpt[1])
                                        res = ctypes.windll.user32.SendMessageW(
                                            lv, 0x1012, 0, ctypes.byref(hi))
                                        if int(res) != -1:
                                            tab._suppress_auto_refresh = False
                                            tab._resume_path_sync_after_navigation()
                                            return
                                    except Exception:
                                        pass

                                cnt = tab._get_selected_count_safe()
                                if cnt is not None and int(cnt) > 0:
                                    return

                                debug_print("[DoubleClick/IEB] Blank area → go_up")
                                tab.go_up(force=True)
                            except Exception as _e:
                                debug_print(f"[DoubleClick/IEB] check: {_e}")

                        QTimer.singleShot(150, _check)
                    else:
                        _state['t'] = t
                        _state['x'] = px
                        _state['y'] = py
                except Exception:
                    pass

            try:
                hh = getattr(FileExplorerTab, '_global_mouse_hook_handle', None) or 0
                return ctypes.windll.user32.CallNextHookEx(hh, nCode, wParam, lParam)
            except Exception:
                return 0

        cb = _HOOKPROC(_hook_proc)
        handle = ctypes.windll.user32.SetWindowsHookExW(14, cb, None, 0)
        if handle:
            FileExplorerTab._global_mouse_hook_cb     = cb       # prevent GC
            FileExplorerTab._global_mouse_hook_handle = handle
            self._lv_hook_handle = handle  # backward compat for logs
            debug_print("[IEB] WH_MOUSE_LL hook installed (singleton)")
        else:
            err = ctypes.windll.kernel32.GetLastError()
            debug_print(f"[IEB] SetWindowsHookExW(WH_MOUSE_LL) failed err={err}")

    def _uninstall_listview_dblclick_hook(self):
        """Remove the singleton WH_MOUSE_LL hook only when the app is closing."""
        self._lv_hook_handle = None
        # Only actually unhook when the MainWindow is closing (no more tabs)
        handle = getattr(FileExplorerTab, '_global_mouse_hook_handle', None)
        if not handle:
            return
        # Check if any other IEB tab still exists
        mw = getattr(self, 'main_window', None)
        if mw and hasattr(mw, 'content_stack'):
            for i in range(mw.content_stack.count()):
                w = mw.content_stack.widget(i)
                if w is not self and isinstance(getattr(w, 'explorer', None), IExplorerBrowserWidget):
                    return  # other IEB tabs still alive, keep hook
        try:
            ctypes.windll.user32.UnhookWindowsHookEx(handle)
            debug_print("[IEB] WH_MOUSE_LL hook removed (singleton)")
        except Exception as e:
            debug_print(f"[IEB] UnhookWindowsHookEx error: {e}")
        finally:
            FileExplorerTab._global_mouse_hook_handle = None
            FileExplorerTab._global_mouse_hook_cb = None

    def _native_listview_hit_test(self, screen_x, screen_y):
        if not HAS_PYWIN:
            return False
        try:
            lv = self._find_syslistview_hwnd()
            if not lv:
                return False
            # convert screen -> client
            pt = (int(screen_x), int(screen_y))
            try:
                cx, cy = win32gui.ScreenToClient(lv, pt)
            except Exception:
                return False

            class POINT(ctypes.Structure):
                _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]

            class LVHITTESTINFO(ctypes.Structure):
                _fields_ = [('pt', POINT), ('flags', ctypes.c_uint), ('iItem', ctypes.c_int), ('iSubItem', ctypes.c_int)]

            info = LVHITTESTINFO()
            info.pt.x = int(cx)
            info.pt.y = int(cy)
            LVM_FIRST = 0x1000
            LVM_HITTEST = LVM_FIRST + 18
            res = ctypes.windll.user32.SendMessageW(lv, LVM_HITTEST, 0, ctypes.byref(info))
            try:
                if int(res) == -1:
                    return False
                return True
            except Exception:
                return False
        except Exception:
            return False

    def eventFilter(self, obj, event):
        # 通过事件过滤器捕获 Explorer 的鼠标按下与双击事件
        from PyQt5.QtCore import QEvent, QTimer, Qt
        
        # 注意：快捷键处理现在由MainWindow的轮询定时器处理，不在这里处理
        
        try:
            if obj is self.explorer:
                if event.type() == QEvent.MouseButtonPress:
                    # 记录按下时的选中项数
                    try:
                        cnt = self._get_selected_count_safe()
                        self._selected_before_click = int(cnt) if cnt is not None else None
                    except Exception:
                        self._selected_before_click = None
                    self._schedule_status_update(track_selection=True)
                elif event.type() == QEvent.ContextMenu:
                    self._schedule_status_update(track_selection=True)
                    global_pos = event.globalPos() if hasattr(event, 'globalPos') else QCursor.pos()
                    if self.show_selected_item_context_menu(global_pos):
                        return True
                elif event.type() == QEvent.MouseButtonRelease:
                    self._schedule_status_update(track_selection=True)
                elif event.type() == QEvent.MouseButtonDblClick:
                    self._schedule_status_update(track_selection=True)
                    # 取消所有之前待处理的双击检查
                    for _t in getattr(self, '_pending_double_click_timers', []):
                        try: _t.stop()
                        except Exception: pass
                    self._pending_double_click_timers = []

                    # 立即在双击位置做 native hit-test —— 这是最可靠的判断
                    # True  = 点中了列表项（文件夹/文件），不触发 go_up
                    # False = 点在空白区，触发 go_up
                    double_click_pos = QCursor.pos()
                    path_before = getattr(self, 'current_path', None)

                    if HAS_PYWIN:
                        hit = self._native_listview_hit_test(double_click_pos.x(), double_click_pos.y())
                        debug_print(f"[DoubleClick] hit-test={hit}, path_before='{path_before}'")
                        if hit:
                            # 点中了项目，让 Explorer 自己处理（打开文件夹/文件）
                            # 无论是否能立即检测到选中项，都启动路径同步定时器，
                            # 确保进入子目录后地址栏及时更新（Explorer可能在双击瞬间已清除选中状态）
                            sel = self._get_selected_count_safe()
                            if sel and int(sel) > 0:
                                self._navigating_folder = True
                            # 解除导航抑制标志，确保本次用户主动双击触发的 Explorer 内部
                            # 导航能被路径同步定时器检测到，不被之前 navigate_to 的 3s 抑制窗口阻断。
                            self._suppress_auto_refresh = False
                            # 50ms 后启动 polling 兜底（NavigateComplete2 直连信号优先触发）
                            QTimer.singleShot(50, self._resume_path_sync_after_navigation)
                        else:
                            # 空白区域双击 —— 用极短延迟（50ms）执行 go_up，
                            # 50ms 仅为让 Explorer 完成 dblclick 内部处理，不会有可见延迟
                            def _do_go_up_blank():
                                try:
                                    # 二次安全确认：路径未变且无选中，再 go_up
                                    if getattr(self, 'current_path', None) != path_before:
                                        return
                                    cnt = self._get_selected_count_safe()
                                    if cnt and int(cnt) > 0:
                                        return
                                    debug_print(f"[DoubleClick] Blank area confirmed, executing go_up")
                                    self.go_up(force=True)
                                except Exception as e:
                                    debug_print(f"[DoubleClick] go_up exception: {e}")
                            t = QTimer(self)
                            t.setSingleShot(True)
                            t.timeout.connect(_do_go_up_blank)
                            t.start(50)
                            self._pending_double_click_timers.append(t)
                    else:
                        # 无 pywin32：退回到 150ms 路径/选中检查（比原 400ms+700ms 仍快很多）
                        selected_before = getattr(self, '_selected_before_click', None)
                        def _fallback_check():
                            try:
                                if getattr(self, '_navigating_folder', False):
                                    self._navigating_folder = False
                                    return
                                cur_path = getattr(self, 'current_path', None)
                                if path_before is not None and cur_path != path_before:
                                    return
                                cnt = self._get_selected_count_safe()
                                if cnt is not None and int(cnt) > 0:
                                    return
                                if selected_before is not None and int(selected_before) > 0:
                                    return
                                debug_print(f"[DoubleClick] Fallback: blank area confirmed, executing go_up")
                                self.go_up(force=True)
                            except Exception as e:
                                debug_print(f"[DoubleClick] Fallback exception: {e}")
                            finally:
                                self._selected_before_click = None
                        t = QTimer(self)
                        t.setSingleShot(True)
                        t.timeout.connect(_fallback_check)
                        t.start(150)
                        self._pending_double_click_timers.append(t)
                elif event.type() in (QEvent.KeyRelease, QEvent.FocusIn, QEvent.Wheel):
                    self._schedule_status_update(track_selection=True)
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def blank_double_click(self, event):
        self.go_up(force=True)

    def _status_bar_mouse_press(self, event):
        """在状态栏按下左键时，开始拖动整个软件窗口。"""
        if event.button() == Qt.LeftButton:
            win = self.window()
            # 优先使用系统原生窗口移动（Qt 5.15+，最大化等情况处理更可靠）
            handle = win.windowHandle()
            if handle is not None and hasattr(handle, 'startSystemMove'):
                try:
                    if handle.startSystemMove():
                        self._status_drag_pos = None
                        event.accept()
                        return
                except Exception:
                    pass
            self._status_drag_pos = event.globalPos() - win.frameGeometry().topLeft()
            event.accept()
            return
        QLabel.mousePressEvent(self.status_bar, event)

    def _status_bar_mouse_move(self, event):
        """拖动状态栏时移动整个软件窗口。"""
        if event.buttons() == Qt.LeftButton and self._status_drag_pos is not None:
            win = self.window()
            if not win.isMaximized():
                win.move(event.globalPos() - self._status_drag_pos)
            event.accept()
            return
        QLabel.mouseMoveEvent(self.status_bar, event)

    def _status_bar_mouse_release(self, event):
        """结束状态栏拖动。"""
        self._status_drag_pos = None
        QLabel.mouseReleaseEvent(self.status_bar, event)

    def _status_bar_mouse_double_click(self, event):
        """双击底部状态栏：手动触发一次“缩放式”重排刷新，兜底修复路径栏偶发卡死。"""
        if event.button() != Qt.LeftButton:
            QLabel.mouseDoubleClickEvent(self.status_bar, event)
            return
        try:
            mw = getattr(self, 'main_window', None)
            if mw and hasattr(mw, '_manual_statusbar_reflow_refresh'):
                mw._manual_statusbar_reflow_refresh()
            else:
                self._manual_pathbar_rebuild()
            event.accept()
        except Exception as e:
            debug_print(f"[StatusBar] manual reflow refresh failed: {e}")
            QLabel.mouseDoubleClickEvent(self.status_bar, event)

    def _manual_pathbar_rebuild(self):
        """仅重建当前标签路径栏（MainWindow 不可用时的兜底）。"""
        pb = getattr(self, 'path_bar', None)
        current_path = getattr(self, 'current_path', '')
        if pb and current_path and not getattr(pb, '_in_edit', False):
            pb.set_path(current_path)
            try:
                pb.force_refresh()
            except Exception:
                pass

    # 移除 on_document_complete 和 eventFilter 相关内容

    def go_up(self, force=False):
        # 返回上一级目录，盘符根目录时导航到"此电脑"
        # 如果 force=True，则绕过鼠标位置检查（用于按钮或程序化调用）
        if not self.current_path:
            return
        
        debug_print(f"[go_up] Called with force={force}, current_path='{self.current_path}'")
        
        # 快速检查：如果是force模式，直接跳过widget检查
        if not force:
            # 仅在明确来自空白区域或路径栏的触发时执行，避免误由文件双击触发
            try:
                pos = QCursor.pos()
                w = QApplication.widgetAt(pos.x(), pos.y())
                # 允许的触发源：底部空白标签或路径栏
                if w is not self.blank and w is not getattr(self, 'path_bar', None):
                    debug_print(f"[go_up] Rejected: not from valid source")
                    return
            except Exception:
                debug_print(f"[go_up] Rejected: exception in widget check")
                return
        
        path = self.current_path
        # 判断是否为盘符根目录，导航到"此电脑"
        if path.endswith(":\\") or path.endswith(":/"):
            debug_print(f"[go_up] Root directory, navigate to MyComputer")
            self.navigate_to('shell:MyComputerFolder', is_shell=True, skip_async_check=True)
            return
        
        parent_path = os.path.dirname(path)
        if parent_path and os.path.exists(parent_path):
            debug_print(f"[go_up] Navigate to parent: {parent_path}")
            # 返回上一级时跳过异步检查，直接导航，提升响应速度
            self.navigate_to(parent_path, skip_async_check=True)
        else:
            debug_print(f"[go_up] Invalid parent path")

    def _normalize_local_path(self, path):
        if not isinstance(path, str) or not path:
            return path
        if path.startswith('shell:') or '::' in path:
            return path
        # 自愈：修复历史会话/缓存中被破坏的 UNC 路径。旧版本把 \\server\share
        # 误存成单反斜杠 \server\share（丢了一个 \），导致 SHParseDisplayName 失败。
        # 统一分隔符后，若以单个分隔符开头（非双）且至少有两段（\主机\共享…），
        # 判定为被破坏的 UNC 并还原 \\ 前缀。本应用只存绝对路径，不会有盘符相对路径。
        unified = path.replace('/', '\\')
        if unified.startswith('\\') and not unified.startswith('\\\\'):
            rest = unified.lstrip('\\')
            if rest and rest[1:2] != ':':  # 排除形如 \C:\ 的异常
                segs = [s for s in rest.split('\\') if s]
                if len(segs) >= 2:
                    unified = '\\\\' + rest
        try:
            return os.path.normpath(unified)
        except Exception:
            return unified

    def __init__(self, parent=None, path="", is_shell=False, select_file=None, defer_nav=False):
        super().__init__(parent)
        self.main_window = parent
        initial_path = path if path else QDir.homePath()
        self.current_path = self._normalize_local_path(initial_path)
        self.select_file = select_file  # 要选中的文件名
        # 延迟首次导航：会话恢复时后台标签用此模式，避免启动瞬间 N 个 IExplorerBrowser
        # 同时创建 COM/导航/overlay 预加载/scandir 造成的 CPU 洪峰。首次可见（showEvent）时才导航。
        self._deferred_nav = None  # (path, is_shell) 待首次可见时执行；None 表示无待处理导航
        self.notepad_plus_plus_path = detect_notepad_plus_plus()
        # 浏览历史记录
        self.history = []
        self.history_index = -1
        # 标志：是否正在程序化导航（用于防止sync时重复添加历史）
        self._navigating_programmatically = False
        # 用于跟踪待处理的双击检查定时器
        self._pending_double_click_timers = []
        # 双击事件唯一ID，用于区分不同的双击操作
        self._double_click_id = 0
        # Win32 WH_MOUSE thread hook for IExplorerBrowser double-click detection
        self._lv_hook_wndproc = None   # ctypes callback (must stay alive to prevent GC)
        self._lv_hook_handle  = None   # HHOOK handle returned by SetWindowsHookExW
        self._is_cleaning_up = False
        
        # 文件系统监控（监控当前路径的变化）
        self.file_watcher = QFileSystemWatcher(self)
        self.file_watcher.directoryChanged.connect(self.on_directory_changed)
        self.file_watcher.fileChanged.connect(self.on_file_changed)
        # 延迟刷新定时器（避免频繁刷新）
        self.refresh_timer = QTimer(self)
        self.refresh_timer.setSingleShot(True)
        self.refresh_timer.timeout.connect(self.delayed_refresh)
        self.refresh_delay_ms = 500  # 500ms延迟
        self._refresh_min_interval_ms = 3000  # 连续刷新最小间隔，避免COM刷新风暴卡界面
        self._last_refresh_ts_ms = 0
        # 防抖机制：记录最近处理的路径和时间，避免事件风暴
        self._last_watcher_event = {}  # {path: timestamp}
        self._watcher_debounce_ms = 3000  # 3s 内的重复事件会被忽略，进一步抑制风暴
        self._refresh_burst_times = []  # 刷新时间戳列表，用于检测风暴
        self._refresh_burst_suppressed_until = 0  # 风暴期间暂停刷新的截止时间戳(ms)
        self._last_file_event = {}  # {file_path: timestamp_ms}
        self._file_event_debounce_ms = 400
        self._watched_files = set()
        self._last_dir_snapshot = None
        # 后台目录快照（兜底轮询）：懒创建的信号对象 + in-flight 去重标志
        self._snapshot_signals = None
        self._snapshot_inflight = False
        self._refresh_active = False
        self._refresh_pending = False
        self._refresh_pending_reason = None
        self._manual_refresh_frozen = False
        # Ctrl+G 选中保护：选中文件后短时间内抑制自动刷新，避免后台目录变动触发的
        # Refresh()（完整重导航）立即清掉刚设置的选中项，导致“跳转选中总是失败”。
        # 注意：仅抑制自动刷新，不影响路径同步（回车进入子目录仍能更新地址栏）。
        self._selection_guard_until = 0.0
        self._status_tracking_deadline_ms = 0
        self.status_update_timer = QTimer(self)
        self.status_update_timer.setSingleShot(True)
        self.status_update_timer.timeout.connect(self.update_explorer_status)
        self.status_tracking_timer = QTimer(self)
        self.status_tracking_timer.setInterval(STATUS_TRACKING_INTERVAL_MS)
        self.status_tracking_timer.timeout.connect(self._poll_status_during_interaction)

        # 目录轮询刷新兜底：处理部分编辑器改文件但目录 watcher 不触发的情况
        self.dir_mtime = None
        self.dir_poll_timer = QTimer(self)
        self.dir_poll_timer.setInterval(8000)  # 8s 低频轮询，后台标签降低I/O开销
        self.dir_poll_timer.timeout.connect(self._poll_directory_changes)
        
        self.setup_ui()
        
        # 安装事件过滤器来处理快捷键（让Ctrl键能穿透到主窗口）
        self.installEventFilter(self)
        if hasattr(self, 'explorer'):
            self.explorer.installEventFilter(self)
        
        if defer_nav:
            # 延迟首次导航：仅在路径栏显示占位路径，暂存导航参数；
            # 不创建 IExplorerBrowser、不导航、不 scandir、不 overlay 预加载，
            # 直到该标签首次可见（showEvent）时才真正导航。
            self._deferred_nav = (self.current_path, is_shell)
            if hasattr(self, 'path_bar') and self.path_bar:
                try:
                    self.path_bar.set_path(self.current_path)
                except Exception:
                    pass
            self.update_tab_title()
        else:
            self.navigate_to(self.current_path, is_shell=is_shell)
        # 路径同步定时器已在setup_ui中启动，这里不需要重复启动
        
        # 如果指定了要选中的文件，延迟选中（等待导航完成）
        # 增加延迟时间确保文件夹完全加载
        if self.select_file:
            QTimer.singleShot(1500, lambda: self.select_file_in_explorer(self.select_file))

    def _release_folder_checker(self, wait_ms=150):
        checker = getattr(self, 'folder_checker', None)
        self.folder_checker = None
        if not checker:
            return
        try:
            checker.finished.disconnect()
        except Exception:
            pass
        try:
            checker.stop()
        except Exception:
            pass
        try:
            if checker.isRunning():
                checker.wait(wait_ms)
        except Exception:
            pass
        try:
            checker.deleteLater()
        except Exception:
            pass

    def cleanup(self):
        if self._is_cleaning_up:
            return
        self._is_cleaning_up = True
        # Remove Win32 subclass hook before any other cleanup
        self._uninstall_listview_dblclick_hook()

        if hasattr(self, '_pending_double_click_timers'):
            for timer in self._pending_double_click_timers:
                try:
                    timer.stop()
                    timer.deleteLater()
                except Exception:
                    pass
            self._pending_double_click_timers = []

        for timer_name in (
            'refresh_timer',
            'status_update_timer',
            'status_tracking_timer',
            'dir_poll_timer',
            '_path_sync_timer',
            '_path_sync_stop_timer',
            '_keepalive_sync_timer',
        ):
            timer = getattr(self, timer_name, None)
            if timer:
                try:
                    timer.stop()
                    timer.deleteLater()
                except Exception:
                    pass
                setattr(self, timer_name, None)

        watcher = getattr(self, 'file_watcher', None)
        if watcher:
            try:
                watched_paths = watcher.directories() + watcher.files()
                if watched_paths:
                    watcher.removePaths(watched_paths)
            except Exception:
                pass

        self._release_folder_checker(wait_ms=150)

        # 关闭 LocationURL 看门狗线程池
        ex = getattr(self, '_com_executor', None)
        if ex is not None:
            try:
                ex.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass
            self._com_executor = None

        explorer = getattr(self, 'explorer', None)
        if explorer:
            try:
                explorer.removeEventFilter(self)
            except Exception:
                pass
            try:
                explorer.dynamicCall('Stop()')
            except Exception:
                pass
            try:
                explorer.clear()
            except Exception:
                pass

        try:
            self.removeEventFilter(self)
        except Exception:
            pass

        self._watched_files.clear()
        self._last_watcher_event.clear()
        self._last_file_event.clear()
        self._last_dir_snapshot = None
        # 断开后台快照信号，避免延迟到达的 QThreadPool 结果回调已销毁的标签
        sig = getattr(self, '_snapshot_signals', None)
        if sig is not None:
            try:
                sig.done.disconnect()
            except Exception:
                pass
            try:
                sig.deleteLater()
            except Exception:
                pass
            self._snapshot_signals = None
        self._snapshot_inflight = False

    # 移除重复的setup_ui，保留带路径栏的实现

    def get_selected_filenames(self):
        """获取选中的文件名列表（仅文件名，含后缀）"""
        filenames = []
        try:
            # IExplorerBrowser 模式：通过 COM 接口直接获取选中路径
            if isinstance(self.explorer, IExplorerBrowserWidget):
                paths = self._get_ieb_selected_paths()
                for p in paths:
                    if p:
                        filenames.append(os.path.basename(str(p)))
                return filenames
            # 旧 QAxWidget 模式：通过Document接口获取SelectedItems
            doc = self.explorer.querySubObject('Document')
            if doc:
                selected = doc.querySubObject('SelectedItems()')
                if selected:
                    count = selected.dynamicCall('Count()')
                    if count and count > 0:
                        for i in range(count):
                            item = selected.querySubObject('Item(int)', i)
                            if item:
                                name = item.dynamicCall('Name()')
                                if name:
                                    filenames.append(str(name))
        except Exception as e:
            debug_print(f"[get_selected_filenames] Error: {e}")
        return filenames

    def showEvent(self, event):
        """标签首次可见时消费延迟导航：真正创建 IExplorerBrowser 并导航。

        会话恢复时后台标签以 defer_nav 模式创建（不导航），切换过去首次可见即在此导航，
        从而把 N 个 Shell 视图的创建/导航分摊到用户实际访问时，消除启动 CPU 洪峰。"""
        super().showEvent(event)
        deferred = getattr(self, '_deferred_nav', None)
        if deferred is not None:
            self._deferred_nav = None
            path, is_shell = deferred
            debug_print(f"[navigate_to] Deferred first navigation on show: '{path}' (is_shell={is_shell})")
            try:
                self.navigate_to(path, is_shell=is_shell)
            except Exception as e:
                debug_print(f"[navigate_to] Deferred navigation failed: {e}")

    def navigate_to(self, path, is_shell=False, add_to_history=True, skip_async_check=False):
        if not is_shell:
            path = self._normalize_local_path(path)
        debug_print(f"[navigate_to] To '{path}' (is_shell={is_shell}, skip_async={skip_async_check})")
        # 导航进行中（慢盘异步解析未完成）：忽略对同一目标的重复点击，避免频繁操作堆积后台线程。
        # 仅拦截相同目标；切换到不同路径仍放行（旧解析结果由导航代号自动作废）。
        if (not is_shell and getattr(self, '_nav_in_progress', False) and
                path == getattr(self, '_nav_in_progress_path', None)):
            debug_print(f"[navigate_to] Ignoring duplicate nav while in progress: {path}")
            return
        # 控制面板及其子目录用原生窗口打开，不嵌入
        if self._is_control_panel_path(path):
            try:
                import subprocess
                launch_detached(['explorer.exe', path])
                show_toast(self, tr("已打开"), tr("控制面板已在新窗口打开"), level="info", duration=2000)
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开控制面板: {}").format(e), level="error")
            if hasattr(self, 'path_bar'):
                self.path_bar.set_path(self.current_path)
            return

        # 快速取消所有待处理的双击检查定时器
        if hasattr(self, '_pending_double_click_timers'):
            for timer in self._pending_double_click_timers:
                try:
                    timer.stop()
                except Exception:
                    pass
            self._pending_double_click_timers = []

        # 停止之前的文件夹检查线程（减少等待时间）
        if hasattr(self, 'folder_checker') and self.folder_checker and self.folder_checker.isRunning():
            self._release_folder_checker(wait_ms=50)  # 减少等待时间从100ms到50ms

        # 支持本地路径和shell特殊路径
        if is_shell:
            # shell:OneDrive 解析为真实路径（Shell.Explorer无法正确显示内容）
            if path.lower() == 'shell:onedrive':
                onedrive_path = os.environ.get('OneDrive', '')
                if onedrive_path and os.path.exists(onedrive_path):
                    self.navigate_to(onedrive_path, is_shell=False, add_to_history=add_to_history, skip_async_check=True)
                    return
            self._hide_loading_indicator()
            self.explorer.dynamicCall("Navigate(const QString&)", path)
            self.current_path = path
            if hasattr(self, 'path_bar'):
                self.path_bar.set_path(path)
            self.update_tab_title()
            # 添加到历史记录
            if add_to_history:
                self._add_to_history(path)
            self.update_explorer_status()
        elif self._is_slow_path(path):
            # 慢盘（网络/UNC/映射盘/OneDrive）：os.path.exists / os.path.isdir 在挂起的
            # 网络路径上会同步阻塞 UI 线程，导致整个程序卡死。提前判定并跳过所有同步
            # 文件系统探测，直接进入导航流程（PIDL 解析已在后台线程异步执行）。
            debug_print(f"[navigate_to] Slow/network path, skipping sync fs checks: {path}")
            self._perform_navigation(path, add_to_history)
        elif os.path.exists(path):
            # 如果skip_async_check=True或禁用异步，直接导航
            # OneDrive/网络路径的os.scandir()可能永久阻塞，直接跳过异步检查
            is_slow = self._is_slow_path(path)
            if is_slow:
                debug_print(f"[navigate_to] OneDrive/network path detected, skipping async check: {path}")
            if skip_async_check or not ASYNC_LOAD_ENABLED or not os.path.isdir(path) or is_slow:
                self._perform_navigation(path, add_to_history)
            else:
                # 异步检查文件夹大小
                self._check_folder_size_async(path, add_to_history)
        elif path.startswith('\\\\') or path.startswith('//'):
            # UNC 网络路径可能因网络延迟导致 os.path.exists 返回 False，直接尝试导航
            debug_print(f"[navigate_to] UNC path, attempting navigation despite exists=False: {path}")
            self._perform_navigation(path, add_to_history)
        else:
            debug_print(f"[navigate_to] Path does not exist: {path}")

    def _is_slow_path(self, path):
        """检测OneDrive/网络/映射网络驱动器路径——这类路径上os.scandir()可能阻塞UI线程"""
        if not path:
            return False
        # 网络UNC路径
        if path.startswith('\\\\') or path.startswith('//'):
            return True
        # OneDrive同步文件夹（路径中含OneDrive关键字）
        path_lower = path.replace('\\', '/').lower()
        if 'onedrive' in path_lower:
            return True
        # 映射网络驱动器（盘符类型 DRIVE_REMOTE=4）——网络抖动时os.scandir()同样永久阻塞
        try:
            import ctypes
            drive = os.path.splitdrive(path)[0]  # e.g. 'D:'
            if drive:
                DRIVE_REMOTE = 4
                if ctypes.windll.kernel32.GetDriveTypeW(drive + '\\') == DRIVE_REMOTE:
                    return True
        except Exception:
            pass
        return False

    def _is_control_panel_path(self, path):
        """判断路径是否为控制面板或其子目录"""
        if not path:
            return False
        s = path.lower()
        # shell:ControlPanelFolder
        if s.startswith('shell:controlpanelfolder'):
            return True
        # 控制面板 CLSID
        if '::{26ee0668-a00a-44d7-9371-beb064c98683}' in s:
            return True
        # 控制面板的子项目通常以 control panel/ 或 control panel\\ 开头
        if s.startswith('control panel') or s.startswith('control panel/') or s.startswith('control panel\\'):
            return True
        # 也可能是 file:///C:/Windows/System32/control.exe 或类似
        if 'control.exe' in s:
            return True
        # 也可能是 shell:::{26ee0668-a00a-44d7-9371-beb064c98683} 或其子路径
        if s.startswith('shell:::{26ee0668-a00a-44d7-9371-beb064c98683}'):
            return True
        # 也可能是 explorer.exe 打开的控制面板子页面，带有 control panel 字样
        if '/control panel/' in s or '\\control panel\\' in s:
            return True
        return False
    
    def _check_folder_size_async(self, path, add_to_history):
        """异步检查文件夹大小并决定是否显示加载指示器"""
        # 显示加载指示器
        self._show_loading_indicator()
        self._folder_checker_done = False

        self._release_folder_checker(wait_ms=50)

        # 创建并启动检查线程
        self.folder_checker = FolderSizeChecker(path, self)

        def _on_checker_finished(p, count, is_large):
            if getattr(self, '_folder_checker_done', False):
                self._release_folder_checker(wait_ms=0)
                return  # 已由超时保护处理
            self._folder_checker_done = True
            self._on_folder_size_checked(p, count, is_large, add_to_history)
            self._release_folder_checker(wait_ms=0)

        self.folder_checker.finished.connect(_on_checker_finished)
        self.folder_checker.start()

        # 超时保护：若线程在 FOLDER_CHECK_TIMEOUT+500ms 内未完成则强制导航
        # 防止云存储/网络路径的os.scandir()永久阻塞
        def _checker_timeout():
            if getattr(self, '_folder_checker_done', True):
                return  # 线程已正常结束
            debug_print(f"[AsyncLoad] FolderSizeChecker timeout, forcing navigation: {path}")
            if hasattr(self, 'folder_checker') and self.folder_checker:
                self._release_folder_checker(wait_ms=50)
            self._folder_checker_done = True
            self._on_folder_size_checked(path, 0, False, add_to_history)

        QTimer.singleShot(FOLDER_CHECK_TIMEOUT + 500, _checker_timeout)
        debug_print(f"[AsyncLoad] Started checking folder: {path}")
    
    def _on_folder_size_checked(self, path, file_count, is_large, add_to_history):
        """文件夹大小检查完成的回调"""
        debug_print(f"[AsyncLoad] Folder checked: {path}, files={file_count}, large={is_large}")
        
        # 执行导航
        self._perform_navigation(path, add_to_history)
        
        # 隐藏加载指示器
        if is_large:
            # 大文件夹延迟隐藏指示器（等待Explorer加载）
            QTimer.singleShot(1000, self._hide_loading_indicator)
        else:
            # 小文件夹立即隐藏
            self._hide_loading_indicator()
    
    def _perform_navigation(self, path, add_to_history):
        """执行实际的导航操作"""
        path = self._normalize_local_path(path)
        old_path = getattr(self, 'current_path', None)
        url = QDir.toNativeSeparators(path)
        
        # 立即更新路径栏（不等到最后，确保先更新 UI）
        if hasattr(self, 'path_bar'):
            self.path_bar.set_path(path)
        
        # 更新当前路径
        self.current_path = path
        
        # 导航到新目录时清除 Git 状态缓存
        self._git_status_cache = None
        
        # 停止目录轮询定时器，避免误触发刷新
        if hasattr(self, 'dir_poll_timer') and self.dir_poll_timer.isActive():
            self.dir_poll_timer.stop()
            debug_print(f"[Navigation] Stopped dir polling timer")
        
        # 设置标志，防止导航期间的自动刷新
        self._suppress_auto_refresh = True
        
        # 使用Navigate2来获得更好的控制
        try:
            # 尝试使用Navigate2获得更好的刷新效果
            self.explorer.dynamicCall("Navigate2(QVariant,QVariant,QVariant,QVariant,QVariant)", 
                                     url, 0, "", None, None)
        except Exception:
            # 回退到普通Navigate
            self.explorer.dynamicCall("Navigate(const QString&)", url)
        
        # 导航完成后清理标志并恢复路径同步
        if getattr(self, '_navigating_folder', False):
            self._navigating_folder = False
        self._resume_path_sync_after_navigation()

        # 更新状态栏
        self.update_explorer_status()
        
        # 更新文件系统监控（只监控真实文件系统路径）
        # 慢盘（网络/UNC/映射盘）：os.path.exists/os.path.isdir/addPath/_build_dir_snapshot
        # 均为同步文件系统调用，在挂起的网络路径上会阻塞 UI 线程导致整个程序卡死，
        # 故对慢盘完全跳过 watcher 注册与快照构建（此类路径的自动刷新本就依赖轮询兜底，
        # 而 _update_dir_polling 已对慢盘跳过轮询）。
        path_is_slow = self._is_slow_path(path)
        if hasattr(self, 'file_watcher') and not path_is_slow:
            # 移除旧路径的监控（旧路径若为慢盘同样跳过，避免 os.path.exists 阻塞）
            if (old_path and not self._is_slow_path(old_path) and
                    os.path.exists(old_path) and os.path.isdir(old_path) and
                    not old_path.startswith('shell:')):
                self._force_remove_watcher(old_path)
            # 添加新路径的监控
            if os.path.isdir(path):
                self._force_remove_watcher(path)
                if self.file_watcher.addPath(path):
                    debug_print(f"[FileWatcher] Now watching: {path}")
                else:
                    debug_print(f"[FileWatcher] Failed to watch: {path}")
                self._refresh_file_watch_paths(path)
                self._last_dir_snapshot = self._build_dir_snapshot(path)
                debug_print(f"[FileWatcher] Now watching: {self.file_watcher.directories()}")
        elif path_is_slow:
            # 慢盘不注册 watcher，清空上次快照，避免下次轮询用旧快照误判
            self._last_dir_snapshot = None
            debug_print(f"[FileWatcher] Skipped watcher for slow path: {path}")

        # 启用低频轮询兜底，处理 watcher 未报告的文件修改时间变化
        self._update_dir_polling(path)
        
        self.update_tab_title()
        if self.main_window and hasattr(self.main_window, 'get_current_tab_widget'):
            try:
                if self.main_window.get_current_tab_widget() is self:
                    self.main_window.update_chat_context()
            except Exception:
                pass
        # 添加到历史记录
        if add_to_history:
            self._add_to_history(path)
        
        # 延迟3秒后允许自动刷新（避免刚导航完就因为文件监视器触发刷新）
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, lambda: setattr(self, '_suppress_auto_refresh', False))
    
    def _show_loading_indicator(self):
        """显示加载指示器（延迟显示）。

        快速导航（绝大多数情况）会在 hide 前取消该定时器，因此进度条不会出现，
        避免路径栏下方进度条一闪而过导致顶部区域"变高又恢复"的布局抖动。
        """
        if not hasattr(self, 'loading_bar'):
            return
        from PyQt5.QtCore import QTimer
        timer = getattr(self, '_loading_show_timer', None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.timeout.connect(self._do_show_loading_indicator)
            self._loading_show_timer = timer
        # 250ms 内完成的导航不显示进度条，消除闪烁
        timer.start(250)

    def _position_loading_bar(self):
        """将悬浮加载进度条定位到文件区域顶部（路径栏正下方），覆盖在 explorer 之上。"""
        if not hasattr(self, 'loading_bar'):
            return
        try:
            top = self.path_bar.height() if hasattr(self, 'path_bar') else 30
            self.loading_bar.setGeometry(0, top, self.width(), 20)
        except Exception:
            pass

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 悬浮加载进度条不在布局中，需随窗口尺寸变化手动跟随宽度
        if getattr(self, 'loading_bar', None) is not None and self.loading_bar.isVisible():
            self._position_loading_bar()

    def _do_show_loading_indicator(self):
        if hasattr(self, 'loading_bar'):
            self._position_loading_bar()
            self.loading_bar.show()
            self.loading_bar.raise_()
            debug_print("[AsyncLoad] Loading indicator shown")

    def _hide_loading_indicator(self):
        """隐藏加载指示器"""
        timer = getattr(self, '_loading_show_timer', None)
        if timer is not None:
            timer.stop()
        if hasattr(self, 'loading_bar'):
            self.loading_bar.hide()
            debug_print("[AsyncLoad] Loading indicator hidden")
    
    def _add_to_history(self, path):
        """添加路径到历史记录（应用内存优化限制）"""
        # 如果当前不在历史末尾，删除当前位置之后的所有历史
        if self.history_index < len(self.history) - 1:
            self.history = self.history[:self.history_index + 1]
        # 添加新路径（避免重复添加相同路径）
        if not self.history or self.history[-1] != path:
            self.history.append(path)
            self.history_index = len(self.history) - 1
            
            # 内存优化：限制历史记录长度
            if len(self.history) > MAX_NAVIGATION_HISTORY:
                # 删除最旧的记录
                remove_count = len(self.history) - MAX_NAVIGATION_HISTORY
                self.history = self.history[remove_count:]
                self.history_index = len(self.history) - 1
                debug_print(f"[Navigation History] Trimmed to {MAX_NAVIGATION_HISTORY} entries")
                
        # 更新主窗口按钮状态
        if self.main_window and hasattr(self.main_window, 'update_navigation_buttons'):
            self.main_window.update_navigation_buttons()
    
    def can_go_back(self):
        """是否可以后退"""
        return self.history_index > 0
    
    def can_go_forward(self):
        """是否可以前进"""
        return self.history_index < len(self.history) - 1
    
    def go_back(self):
        """后退到上一个位置"""
        if self.can_go_back():
            self.history_index -= 1
            path = self.history[self.history_index]
            is_shell = path.startswith('shell:')
            # 设置标志，防止sync时重复添加历史
            self._navigating_programmatically = True
            self.navigate_to(path, is_shell=is_shell, add_to_history=False)
            # 延迟重置标志，确保sync不会在导航完成前被触发
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1000, lambda: setattr(self, '_navigating_programmatically', False))
            # 更新主窗口按钮状态
            if self.main_window and hasattr(self.main_window, 'update_navigation_buttons'):
                self.main_window.update_navigation_buttons()
    
    def go_forward(self):
        """前进到下一个位置"""
        if self.can_go_forward():
            self.history_index += 1
            path = self.history[self.history_index]
            is_shell = path.startswith('shell:')
            # 设置标志，防止sync时重复添加历史
            self._navigating_programmatically = True
            self.navigate_to(path, is_shell=is_shell, add_to_history=False)
            # 延迟重置标志，确保sync不会在导航完成前被触发
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1000, lambda: setattr(self, '_navigating_programmatically', False))
            # 更新主窗口按钮状态
            if self.main_window and hasattr(self.main_window, 'update_navigation_buttons'):
                self.main_window.update_navigation_buttons()
    
    def on_directory_changed(self, path):
        """文件系统监控：目录内容发生变化（带防抖+风暴检测）"""
        import time, os
        current_time = time.time() * 1000  # 转为毫秒
        # 目录已不存在，强制移除watcher，防止事件风暴
        if not os.path.exists(path):
            debug_print(f"[FileWatcher] Directory not exist, remove watcher: {path}")
            self._force_remove_watcher(path)
            return

        # ── 风暴计数：记录短时间内收到的事件数 ──
        storm_times = getattr(self, '_watcher_storm_times', [])
        storm_times = [t for t in storm_times if current_time - t < 10000]  # 10s 窗口
        storm_times.append(current_time)
        self._watcher_storm_times = storm_times
        is_storm = len(storm_times) > 5  # 10s 内超过 5 个事件视为风暴

        if hasattr(self, '_last_watcher_event'):
            last_time = self._last_watcher_event.get(path, 0)
            time_since_last = current_time - last_time
            if path == getattr(self, 'current_path', None) and self.refresh_timer.isActive():
                return
            # 风暴期间使用更长的防抖时间（5s），正常情况 3s
            debounce = 5000 if is_storm else self._watcher_debounce_ms
            if time_since_last < debounce:
                return
            self._last_watcher_event[path] = current_time
            if len(self._last_watcher_event) > 10:
                sorted_items = sorted(self._last_watcher_event.items(), key=lambda x: x[1], reverse=True)
                self._last_watcher_event = dict(sorted_items[:10])
        else:
            self._last_watcher_event = {path: current_time}
        if os.path.isdir(path) and not self._is_slow_path(path):
            current_snapshot = self._build_dir_snapshot(path)
            if current_snapshot is not None:
                previous_snapshot = getattr(self, '_last_dir_snapshot', None)
                if previous_snapshot == current_snapshot:
                    debug_print(f"[FileWatcher] Ignored internal-only directory change: {path}")
                    return
                if path == getattr(self, 'current_path', None):
                    self._last_dir_snapshot = current_snapshot
        debug_print(f"[FileWatcher] Directory changed: {path} (storm={is_storm})")
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[FileWatcher] Auto-refresh suppressed during navigation")
            return
        if not getattr(self, '_refresh_active', True):
            self._request_refresh(reason="watcher")
            debug_print(f"[FileWatcher] Background tab marked dirty: {path}")
            return
        if path == self.current_path:
            self._request_refresh(reason="watcher")

    def _schedule_refresh(self, reason="manual"):
        """统一的刷新调度，避免重复代码"""
        import time
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[AutoRefresh] Suppressed during navigation (reason={reason})")
            return

        delay_ms = int(getattr(self, 'refresh_delay_ms', 500))
        min_interval_ms = int(getattr(self, '_refresh_min_interval_ms', 0))
        last_refresh_ms = float(getattr(self, '_last_refresh_ts_ms', 0) or 0)
        if min_interval_ms > 0 and last_refresh_ms > 0:
            elapsed_ms = (time.time() * 1000) - last_refresh_ms
            if elapsed_ms < min_interval_ms:
                delay_ms = max(delay_ms, int(min_interval_ms - elapsed_ms))

        if self.refresh_timer.isActive():
            remaining = self.refresh_timer.remainingTime()
            # 已有更早的刷新计划时不延后，避免连续事件导致“永远不刷新”
            if remaining > 0 and remaining <= delay_ms:
                return
            self.refresh_timer.stop()
            self.refresh_timer.start(delay_ms)
            debug_print(f"[AutoRefresh] Refresh timer adjusted to {delay_ms}ms (reason={reason})")
        else:
            debug_print(f"[AutoRefresh] Scheduling refresh in {delay_ms}ms (reason={reason})")
            self.refresh_timer.start(delay_ms)
    
    def delayed_refresh(self):
        """延迟刷新：避免频繁刷新。风暴期间跳过或进一步延后。"""
        import time
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[FileWatcher] Auto-refresh suppressed during navigation")
            return
        if self._selection_guard_active():
            debug_print(f"[FileWatcher] Auto-refresh suppressed during selection guard")
            return
        now_ms = time.time() * 1000

        # 风暴检测：如果仍在风暴中且距上次刷新太近，重新延后
        storm_times = getattr(self, '_watcher_storm_times', [])
        storm_count = sum(1 for t in storm_times if now_ms - t < 10000)
        is_storm = storm_count > 5
        last_refresh_ms = float(getattr(self, '_last_refresh_ts_ms', 0) or 0)
        if is_storm and last_refresh_ms > 0 and (now_ms - last_refresh_ms) < 8000:
            # 风暴中两次刷新间隔不到 8s，重新延后
            remain = int(8000 - (now_ms - last_refresh_ms))
            self.refresh_timer.start(remain)
            debug_print(f"[AutoRefresh] Storm active, deferring refresh by {remain}ms")
            return

        if getattr(self, '_manual_refresh_frozen', False):
            debug_print(f"[AutoRefresh] Manually frozen, skipping refresh execution")
            return
        self._last_refresh_ts_ms = now_ms
        self._refresh_pending = False
        self._refresh_pending_reason = None
        debug_print(f"[FileWatcher] Auto-refreshing: {self.current_path}")
        if hasattr(self, 'explorer') and self.current_path:
            try:
                try:
                    self.explorer.dynamicCall('Refresh()')
                except Exception:
                    is_shell = self.current_path.startswith('shell:')
                    if is_shell:
                        self.explorer.dynamicCall('Navigate(const QString&)', self.current_path)
                    elif self.current_path.startswith('\\\\'):
                        self.explorer.dynamicCall('Navigate(const QString&)', self.current_path)
                    else:
                        url = 'file:///' + self.current_path.replace('\\', '/')
                        self.explorer.dynamicCall('Navigate2(const QVariant&)', url)
                debug_print(f"[FileWatcher] Refresh completed")
            except Exception as e:
                debug_print(f"[FileWatcher] Refresh error: {e}")
        # 刷新后更新快照（风暴时延迟执行，避免阻塞 UI）
        if is_storm:
            QTimer.singleShot(500, self._deferred_post_refresh_snapshot)
        else:
            self._last_dir_snapshot = self._build_dir_snapshot(self.current_path)
            self.update_explorer_status()

    def _deferred_post_refresh_snapshot(self):
        """延迟更新快照，避免风暴期间连续 scandir 阻塞 UI"""
        try:
            self._last_dir_snapshot = self._build_dir_snapshot(self.current_path)
            self.update_explorer_status()
        except Exception:
            pass

    def _poll_directory_changes(self):
        """兜底轮询目录元数据，解决文件编辑后目录 watcher 不触发的问题"""
        # 如果设置了抑制标志，不触发刷新
        if getattr(self, '_suppress_auto_refresh', False):
            return
        if self._selection_guard_active():
            return
        if not getattr(self, '_refresh_active', True):
            return

        # 风暴期间跳过轮询（watcher 已在处理，避免额外 scandir 阻塞）
        import time
        now_ms = time.time() * 1000
        storm_times = getattr(self, '_watcher_storm_times', [])
        storm_count = sum(1 for t in storm_times if now_ms - t < 10000)
        if storm_count > 5:
            return

        path = self.current_path
        if not path or not os.path.isdir(path):
            return
        # OneDrive/网络路径的os.scandir()会阻塞UI线程，不做轮询
        if self._is_slow_path(path):
            return
        # 快照计算（scandir + 逐项 stat）放到后台线程池，避免大目录每 8s 在 UI 线程卡顿。
        # 上一次计算尚未返回时跳过本次，防止慢目录任务堆积。
        if getattr(self, '_snapshot_inflight', False):
            return
        try:
            if not hasattr(self, '_snapshot_signals') or self._snapshot_signals is None:
                self._snapshot_signals = _DirSnapshotSignals(self)
                self._snapshot_signals.done.connect(self._on_dir_snapshot_ready)
            self._snapshot_inflight = True
            runnable = _DirSnapshotRunnable(path, self._should_ignore_internal_dir_entry, self._snapshot_signals)
            QThreadPool.globalInstance().start(runnable)
        except Exception as e:
            self._snapshot_inflight = False
            debug_print(f"[DirPoll] Failed to start snapshot worker: {e}")

    def _on_dir_snapshot_ready(self, snap_path, current_snapshot):
        """后台快照计算完成（回到 UI 线程）：与上次快照比较，变化则调度刷新。"""
        self._snapshot_inflight = False
        # 计算期间已切换目录：丢弃过期结果
        if snap_path != getattr(self, 'current_path', None):
            return
        if current_snapshot is None:
            return
        if self._last_dir_snapshot is None:
            self._last_dir_snapshot = current_snapshot
            return
        if current_snapshot != self._last_dir_snapshot:
            debug_print(f"[DirPoll] Detected snapshot change for {snap_path}, scheduling refresh")
            self._last_dir_snapshot = current_snapshot
            self._request_refresh(reason="poll")

    def _update_dir_polling(self, path):
        """根据当前路径启动或停止兜底轮询"""
        if not hasattr(self, 'dir_poll_timer'):
            return
        # OneDrive/网络路径不做轮询，避免os.scandir()阻塞UI线程
        if path and self._is_slow_path(path):
            if self.dir_poll_timer.isActive():
                self.dir_poll_timer.stop()
                debug_print(f"[DirPoll] Stopped polling for slow path: {path}")
            return
        if path and os.path.isdir(path):
            # 立即更新快照，避免刚导航时误判为变化
            self.dir_mtime = self._get_dir_mtime(path)
            self._last_dir_snapshot = self._build_dir_snapshot(path)
            debug_print(f"[DirPoll] Updated mtime for {path}: {self.dir_mtime}")
            if getattr(self, '_refresh_active', False) and not self.dir_poll_timer.isActive():
                self.dir_poll_timer.start()
                debug_print(f"[DirPoll] Started polling {path}")
        else:
            if self.dir_poll_timer.isActive():
                self.dir_poll_timer.stop()
                debug_print(f"[DirPoll] Stopped polling")
            self.dir_mtime = None
            self._last_dir_snapshot = None

    def _get_dir_mtime(self, path):
        """安全获取目录修改时间"""
        try:
            return os.stat(path).st_mtime
        except Exception:
            return None

    def update_explorer_status(self):
        """更新嵌入 Explorer 下方状态栏（仅显示 Git 状态）"""
        if not hasattr(self, 'status_bar'):
            return
        path = getattr(self, 'current_path', None)
        if not path or path.startswith('shell:') or '::' in path:
            self.status_bar.setText('')
            return

        # 只显示 Git 状态摘要（仅读缓存，不阻塞 UI）
        cache = getattr(self, '_git_status_cache', None)
        git_summary = cache.get('result') if (cache and cache.get('path') == path) else None
        self.status_bar.setText(git_summary or '')
        # 异步刷新 Git 状态（后台线程）
        self._request_git_status_async(path)

    def _get_selection_entries(self):
        """返回选中条目列表，每项包含 is_file 与 size"""
        try:
            doc = self.explorer.querySubObject('Document')
            if not doc:
                return []
            selected = doc.querySubObject('SelectedItems()')
            if not selected:
                return []
            count = selected.dynamicCall('Count()')
            if not count or count <= 0:
                return []
            count_int = int(count)
            collect_file_sizes = count_int <= STATUS_SELECTION_METADATA_LIMIT
            entries = []
            for i in range(count_int):
                item = selected.querySubObject('Item(int)', i)
                if not item:
                    continue
                path = item.dynamicCall('Path()')
                name = item.dynamicCall('Name()')
                if not path and name:
                    # 部分场景仅返回名称，尝试拼接
                    path = os.path.join(self.current_path, str(name))
                if not path:
                    continue
                path_str = str(path)
                is_file = os.path.isfile(path_str)
                size = None
                if is_file and collect_file_sizes:
                    try:
                        size = os.path.getsize(path_str)
                    except Exception:
                        size = None
                entries.append({
                    'path': path_str,
                    'is_file': is_file,
                    'size': size,
                })
            return entries
        except Exception:
            return None

    def _get_selected_paths(self):
        entries = self._get_selection_entries()
        if not entries:
            return []
        return [entry.get('path') for entry in entries if entry and entry.get('path')]

    def _get_single_selected_path(self):
        # 先用廉价的计数（单次 COM Count）短路：仅在恪好选中 1 项时才逐项枚举，
        # 避免大目录 Ctrl+A 后右键时对成千上万项做 COM Item()+isfile 的无谓枚举。
        cnt = self._get_selected_count_safe()
        if cnt is not None and cnt != 1:
            return None
        paths = self._get_selected_paths()
        if len(paths) == 1:
            return paths[0]
        return None

    def _launch_selected_file_with_program(self, file_path, program_path, display_name):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("当前选中项不是可打开的文件"), level="warning")
            return False
        if not program_path:
            show_toast(self, tr("提示"), tr("未找到 {}").format(display_name), level="warning")
            return False
        try:
            launch_detached_async([program_path, file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法使用 {} 打开文件: {}").format(display_name, e), level="error")
            return False

    def open_selected_with_default_app(self, file_path):
        if not file_path or not os.path.exists(file_path):
            show_toast(self, tr("提示"), tr("当前选中项不可打开"), level="warning")
            return False
        try:
            if os.path.isdir(file_path):
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    self.main_window.add_new_tab(file_path)
            elif os.name == 'nt':
                os.startfile(file_path)
            else:
                launch_detached([file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开选中项: {}").format(e), level="error")
            return False

    def open_selected_with_notepad(self, file_path):
        return self._launch_selected_file_with_program(file_path, 'notepad.exe', tr('记事本'))

    def open_selected_with_notepad_plus_plus(self, file_path):
        return self._launch_selected_file_with_program(file_path, getattr(self, 'notepad_plus_plus_path', None), 'Notepad++')

    def open_selected_with_system_dialog(self, file_path):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("当前选中项不是可打开的文件"), level="warning")
            return False
        if os.name != 'nt':
            show_toast(self, tr("提示"), tr("当前系统不支持打开“选择其他应用”对话框"), level="warning")
            return False
        try:
            launch_detached_async(['rundll32.exe', 'shell32.dll,OpenAs_RunDLL', file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开“选择其他应用”对话框: {}").format(e), level="error")
            return False

    def open_selected_parent_folder(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return False
        if os.path.isdir(file_path):
            folder_path = file_path
            select_file = None
        else:
            folder_path = os.path.dirname(file_path)
            select_file = os.path.basename(file_path)
        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
            self.main_window.add_new_tab(folder_path, select_file=select_file)
            return True
        return False

    def show_selected_item_context_menu(self, global_pos):
        file_path = self._get_single_selected_path()
        if not file_path or not os.path.exists(file_path):
            return False


        menu = QMenu(self)
        open_action = menu.addAction(tr("打开"))
        open_action.triggered.connect(lambda: self.open_selected_with_default_app(file_path))

        open_folder_action = menu.addAction(tr("打开所在目录"))
        open_folder_action.triggered.connect(lambda: self.open_selected_parent_folder(file_path))

        if os.path.isfile(file_path):
            default_action = menu.addAction(tr("用系统默认程序打开"))
            default_action.triggered.connect(lambda: self.open_selected_with_default_app(file_path))

            menu.addSeparator()

            notepad_action = menu.addAction(tr("用记事本打开"))
            notepad_action.triggered.connect(lambda: self.open_selected_with_notepad(file_path))

            notepadpp_action = menu.addAction(tr("用 Notepad++ 打开"))
            notepadpp_action.setEnabled(bool(getattr(self, 'notepad_plus_plus_path', None)))
            if not getattr(self, 'notepad_plus_plus_path', None):
                notepadpp_action.setToolTip(tr("未检测到 Notepad++"))
            notepadpp_action.triggered.connect(lambda: self.open_selected_with_notepad_plus_plus(file_path))

            menu.addSeparator()

            system_dialog_action = menu.addAction(tr("选择其他应用..."))
            system_dialog_action.triggered.connect(lambda: self.open_selected_with_system_dialog(file_path))

        menu.exec_(global_pos)
        return True

    def select_file_in_explorer(self, filename, retries=6, delay_ms=250):
        """在Explorer控件中选中当前目录下指定的文件或文件夹。"""
        try:
            debug_print(f"[SelectFile] Attempting to select file: {filename}")
            
            # 构建完整路径
            full_path = os.path.join(self.current_path, filename)
            if not os.path.exists(full_path):
                debug_print(f"[SelectFile] File not found: {full_path}")
                return False

            if self._select_ieb_item_by_name(filename):
                return True
            
            # 使用Windows API选中文件（通过查找ListView控件并发送消息）
            try:
                import ctypes
                from ctypes import wintypes
                
                # Windows API常量
                LVM_SETITEMSTATE = 0x102B
                LVM_ENSUREVISIBLE = 0x1013
                LVM_GETITEMCOUNT = 0x1004
                LVM_GETITEMTEXT = 0x102D
                LVIF_STATE = 0x0008
                LVIS_SELECTED = 0x0002
                LVIS_FOCUSED = 0x0001
                
                # 获取当前窗口句柄
                hwnd = int(self.explorer.winId())
                
                # 查找ListView控件（通常类名是 SysListView32）
                user32 = ctypes.windll.user32
                
                def enum_child_windows(parent_hwnd):
                    """枚举所有子窗口"""
                    handles = []
                    def callback(hwnd, lParam):
                        handles.append(hwnd)
                        return True
                    
                    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
                    enum_proc = WNDENUMPROC(callback)
                    user32.EnumChildWindows(parent_hwnd, enum_proc, 0)
                    return handles
                
                # 查找ListView控件
                listview_hwnd = None
                for child_hwnd in enum_child_windows(hwnd):
                    class_name = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(child_hwnd, class_name, 256)
                    if 'SysListView32' in class_name.value:
                        listview_hwnd = child_hwnd
                        debug_print(f"[SelectFile] Found ListView control: {listview_hwnd}")
                        break
                
                if not listview_hwnd:
                    debug_print(f"[SelectFile] ListView control not found")
                    if retries > 0:
                        QTimer.singleShot(delay_ms, lambda: self.select_file_in_explorer(filename, retries - 1, delay_ms))
                    return False
                
                # 获取ListView中的项目数
                item_count = user32.SendMessageW(listview_hwnd, LVM_GETITEMCOUNT, 0, 0)
                debug_print(f"[SelectFile] ListView has {item_count} items")
                if item_count <= 0:
                    if retries > 0:
                        debug_print(f"[SelectFile] ListView not ready, retrying... remaining={retries}")
                        QTimer.singleShot(delay_ms, lambda: self.select_file_in_explorer(filename, retries - 1, delay_ms))
                    return False
                
                # 遍历所有项目，查找匹配的文件名
                for i in range(item_count):
                    # 获取项目文本需要使用进程间通信（因为ListView在不同进程）
                    # 这里使用简化方法：通过Document接口获取文件列表来匹配索引
                    try:
                        doc = self.explorer.querySubObject('Document')
                        if doc:
                            folder = doc.querySubObject('Folder')
                            if folder:
                                items = folder.querySubObject('Items()')
                                if items and i < items.dynamicCall('Count()'):
                                    item = items.querySubObject('Item(int)', i)
                                    if item:
                                        item_name = item.dynamicCall('Name()')
                                        if item_name == filename:
                                            debug_print(f"[SelectFile] Found file at index {i}: {filename}")
                                            
                                            # 定义LVITEM结构
                                            class LVITEM(ctypes.Structure):
                                                _fields_ = [
                                                    ('mask', wintypes.UINT),
                                                    ('iItem', ctypes.c_int),
                                                    ('iSubItem', ctypes.c_int),
                                                    ('state', wintypes.UINT),
                                                    ('stateMask', wintypes.UINT),
                                                    ('pszText', wintypes.LPWSTR),
                                                    ('cchTextMax', ctypes.c_int),
                                                    ('iImage', ctypes.c_int),
                                                    ('lParam', wintypes.LPARAM),
                                                ]
                                            
                                            # 取消所有项目的选中状态
                                            for j in range(item_count):
                                                lvi = LVITEM()
                                                lvi.mask = LVIF_STATE
                                                lvi.state = 0
                                                lvi.stateMask = LVIS_SELECTED | LVIS_FOCUSED
                                                user32.SendMessageW(listview_hwnd, LVM_SETITEMSTATE, j, ctypes.byref(lvi))
                                            
                                            # 选中并聚焦目标项
                                            lvi = LVITEM()
                                            lvi.mask = LVIF_STATE
                                            lvi.state = LVIS_SELECTED | LVIS_FOCUSED
                                            lvi.stateMask = LVIS_SELECTED | LVIS_FOCUSED
                                            result = user32.SendMessageW(listview_hwnd, LVM_SETITEMSTATE, i, ctypes.byref(lvi))
                                            debug_print(f"[SelectFile] SendMessage result: {result}")
                                            
                                            # 确保可见
                                            user32.SendMessageW(listview_hwnd, LVM_ENSUREVISIBLE, i, 0)
                                            
                                            # 设置焦点到ListView
                                            self.activateWindow()
                                            self.raise_()
                                            if hasattr(self, 'explorer') and self.explorer:
                                                self.explorer.setFocus(Qt.OtherFocusReason)
                                            user32.SetFocus(listview_hwnd)
                                            
                                            debug_print(f"[SelectFile] Successfully selected file via API: {filename}")
                                            # 用户随后可能按 Enter 进入选中的文件夹。
                                            # NavigateComplete2 信号在某些环境下不可用，
                                            # 通过启动路径同步定时器来兜底捕获导航变化。
                                            self._suppress_auto_refresh = False
                                            self._arm_selection_guard()
                                            self.start_path_sync_timer(duration_ms=8000)
                                            return True
                    except Exception as e:
                        debug_print(f"[SelectFile] Error matching item {i}: {e}")
                        continue
                
                debug_print(f"[SelectFile] File not found in ListView: {filename}")
                if retries > 0:
                    debug_print(f"[SelectFile] Target not visible yet, retrying... remaining={retries}")
                    QTimer.singleShot(delay_ms, lambda: self.select_file_in_explorer(filename, retries - 1, delay_ms))
                return False
                
            except Exception as e:
                debug_print(f"[SelectFile] Windows API method failed: {e}")
                import traceback
                traceback.print_exc()
                return False
        
        except Exception as e:
            debug_print(f"[SelectFile] Error: {e}")
            import traceback
            traceback.print_exc()
            return False


class DragDropTabWidget(QTabWidget):
    """支持拖放文件夹的自定义QTabWidget"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.main_window = parent
        self.currentChanged.connect(self._refresh_close_button)

    def _refresh_close_button(self, idx):
        tabbar = self.tabBar()
        if hasattr(tabbar, 'show_close_button_under_cursor'):
            tabbar.show_close_button_under_cursor()
    """支持拖放文件夹的自定义QTabWidget"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.main_window = parent

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def dragMoveEvent(self, event):
        """允许在整个 TabWidget 区域内拖动"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def mouseDoubleClickEvent(self, event):
        """捕获 TabWidget 区域的双击事件"""
        from PyQt5.QtCore import QPoint
        # 获取 TabBar 的几何位置
        tabbar = self.tabBar()
        # 将事件位置转换为 TabBar 的坐标系
        tabbar_pos = tabbar.mapFrom(self, event.pos())
        
        debug_print(f"[DEBUG] TabWidget double click: pos={event.pos()}, tabbar_pos={tabbar_pos}")
        debug_print(f"[DEBUG] TabBar rect: {tabbar.rect()}")
        
        # 检查点击是否在 TabBar 的矩形范围内（使用 TabBar 自己的坐标系）
        in_tabbar = tabbar.rect().contains(tabbar_pos)
        debug_print(f"[DEBUG] In TabBar: {in_tabbar}")
        
        if in_tabbar:
            # 在 TabBar 内，检查是否点击在空白区域
            clicked_tab = tabbar.tabAt(tabbar_pos)
            debug_print(f"[DEBUG] Clicked tab index: {clicked_tab}")
            
            if clicked_tab == -1:
                # 空白区域，打开新标签页（归属本标签组，左右分屏各自独立）
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    debug_print(f"[DEBUG] Opening new tab from TabBar blank area...")
                    self.main_window.add_new_tab(target_tabwidget=self)
                    return
        else:
            # 不在 TabBar 内，检查是否在标签页头部区域（TabBar 右侧的空白）
            # 获取 TabWidget 的 TabBar 所在的区域高度
            if event.pos().y() < tabbar.height():
                debug_print(f"[DEBUG] Click is in tab header area but outside TabBar")
                # 这是标签头和按钮之间的空白区域，打开新标签页（归属本标签组）
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    debug_print(f"[DEBUG] Opening new tab from header blank area...")
                    self.main_window.add_new_tab(target_tabwidget=self)
                    return
        
        super().mouseDoubleClickEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            
            # 获取拖拽位置
            drop_pos = event.pos()
            tabbar = self.tabBar()
            tabbar_pos = tabbar.mapFrom(self, drop_pos)
            
            # 检查是否拖拽到标签栏区域
            in_tabbar_area = drop_pos.y() < tabbar.height()
            
            debug_print(f"[DEBUG] Drop event: pos={drop_pos}, in_tabbar_area={in_tabbar_area}")
            
            for url in urls:
                path = None
                # 尝试获取本地文件路径
                if url.isLocalFile():
                    path = url.toLocalFile()
                else:
                    # 尝试从 URL 字符串中提取路径（支持网络路径）
                    url_str = url.toString()
                    if url_str.startswith('file:///'):
                        from urllib.parse import unquote
                        path = unquote(url_str[8:])
                        if os.name == 'nt' and path.startswith('/'):
                            path = path[1:]
                    elif url_str.startswith('file://'):
                        from urllib.parse import unquote
                        # 网络路径 file://server/share
                        path = '\\\\' + unquote(url_str[7:]).replace('/', '\\')
                
                if path and os.path.exists(path):
                    debug_print(f"[DEBUG] Processing dropped path: {path}")
                    if os.path.isdir(path):
                        # 如果是文件夹，打开新标签页（归属当前标签组）
                        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                            self.main_window.add_new_tab(path, target_tabwidget=self)
                    elif os.path.isfile(path):
                        # 如果是文件，打开其所在文件夹（归属本标签组）
                        folder = os.path.dirname(path)
                        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                            self.main_window.add_new_tab(folder, target_tabwidget=self)
            event.acceptProposedAction()
        else:
            event.ignore()



# 自定义MenuBar以支持右键菜单
from PyQt5.QtWidgets import QMenuBar
class CustomMenuBar(QMenuBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
    
    def mousePressEvent(self, event):
        """处理菜单栏的鼠标点击"""
        if event.button() == Qt.RightButton:
            pos = event.pos()
            action = self.actionAt(pos)
            
            debug_print(f"[DEBUG] CustomMenuBar right click at {pos}, action: {action}")
            
            if action and self.main_window:
                if hasattr(self.main_window, 'bookmark_actions') and action in self.main_window.bookmark_actions:
                    node = self.main_window.bookmark_actions[action]
                    bookmark_id = node.get('id')
                    bookmark_name = node.get('name', '')
                    
                    debug_print(f"[DEBUG] Found bookmark: {bookmark_name} (ID: {bookmark_id})")
                    
                    # 检查是否是特殊书签（不允许删除）
                    special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
                    is_special = any(bookmark_name.startswith(icon) for icon in special_icons)
                    
                    debug_print(f"[DEBUG] Is special bookmark: {is_special}")
                    
                    if not is_special:
                        global_pos = self.mapToGlobal(pos)
                        debug_print(f"[DEBUG] Showing context menu at: {global_pos}")
                        self.main_window.show_bookmark_context_menu(global_pos, bookmark_id, bookmark_name)
                        event.accept()
                        return
                else:
                    debug_print(f"[DEBUG] No bookmark action found")
        
        super().mousePressEvent(event)

# 自定义 TabBar 以支持双击空白区域打开新标签页和悬停显示关闭按钮
from PyQt5.QtWidgets import QTabBar, QToolButton
from PyQt5.QtCore import QEvent, QPoint
from PyQt5.QtGui import QIcon
class CustomTabBar(QTabBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = None
        self.owner_tabwidget = None  # 所属标签组的 QTabWidget（左侧或右侧分屏组）
        self.hovered_tab = -1  # 当前鼠标悬停的标签页索引
        self.setMouseTracking(True)  # 启用鼠标追踪
        # 自行处理拖拽（组内重排 + 跨组转移），不用原生 setMovable：
        # 原生拖拽会把被拖标签夹在标签栏自身右边缘，拖出本栏时“卡在右侧看不到”，
        # 且无法给跨组拖拽提供清晰的落点预览。手动实现可全程显示跟随光标的浮动预览。
        self.setMovable(False)
        # 连接标签移动信号（手动 moveTab 仍会触发 tabMoved → on_tab_moved 同步内容栈）
        self.tabMoved.connect(self.on_tab_moved)
        # 拖拽期间抑制重型 on_tab_changed 操作
        self._is_dragging_tab = False
        # 拖拽状态：内容面板(对象身份)、标题、起始索引、按下位置、是否已进入拖拽、浮动预览
        self._press_content = None
        self._press_title = ""
        self._press_index = -1
        self._press_pos = None
        self._dragging = False
        self._drag_preview = None
        from PyQt5.QtCore import QTimer
        self._drag_end_timer = QTimer(self)
        self._drag_end_timer.setSingleShot(True)
        self._drag_end_timer.setInterval(150)
        self._drag_end_timer.timeout.connect(self._on_drag_end_timeout)

    def _owner_tw(self):
        """返回所属标签组的 QTabWidget；优先显式 owner，回退到父控件。"""
        tw = getattr(self, 'owner_tabwidget', None)
        if tw is not None:
            return tw
        return self.parentWidget()

    def _on_drag_end_timeout(self):
        """拖拽结束后触发一次完整的 on_tab_changed（去抖后执行）"""
        self._is_dragging_tab = False
        if self.main_window:
            self.main_window._tab_drag_in_progress = False
            tw = self._owner_tw()
            idx = tw.currentIndex() if tw is not None else -1
            self.main_window._on_group_tab_changed(tw, idx)

    def mousePressEvent(self, event):
        # 记录拖拽起点：内容面板(对象身份)、标题、索引、按下位置；供组内重排与跨组转移使用
        self._press_content = None
        self._press_title = ""
        self._press_index = -1
        self._press_pos = None
        self._dragging = False
        try:
            if event.button() == Qt.LeftButton and self.main_window is not None:
                # 点击本组任一标签即把该组设为活动面板：即使不切换标签，右上角按钮/终端/
                # TortoiseGit 等也能作用于用户正在操作的这一侧，而非总是左侧。
                if hasattr(self.main_window, 'set_active_pane_to_group'):
                    try:
                        self.main_window.set_active_pane_to_group(self._owner_tw())
                    except Exception:
                        pass
                idx = self.tabAt(event.pos())
                if idx >= 0:
                    cs = self.main_window._content_stack_for(self._owner_tw())
                    if cs is not None and idx < cs.count():
                        self._press_content = cs.widget(idx)
                        self._press_title = self.tabText(idx)
                        self._press_index = idx
                        self._press_pos = event.pos()
        except Exception:
            self._press_content = None
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        # 悬停追踪：更新鼠标下的标签索引（供关闭按钮显示等）
        self.hovered_tab = self.tabAt(event.pos())
        # 拖拽判定：左键按住并移动超过阈值 → 进入拖拽，显示跟随光标的浮动预览
        if (self._press_index >= 0 and (event.buttons() & Qt.LeftButton)
                and self._press_pos is not None):
            if not self._dragging:
                try:
                    from PyQt5.QtWidgets import QApplication
                    threshold = QApplication.startDragDistance()
                except Exception:
                    threshold = 6
                if (event.pos() - self._press_pos).manhattanLength() >= threshold:
                    self._dragging = True
                    self._is_dragging_tab = True
                    if self.main_window is not None:
                        self.main_window._tab_drag_in_progress = True
            if self._dragging:
                self._update_drag_preview(event.globalPos())
        super().mouseMoveEvent(event)

    def _update_drag_preview(self, gpos):
        """拖拽中显示跟随光标的浮动预览（标签标题气泡），并按目标组区分提示样式。"""
        dest_tw = None
        try:
            if self.main_window is not None and hasattr(self.main_window, '_pane_group_hit_test'):
                dest_tw, _idx = self.main_window._pane_group_hit_test(gpos)
        except Exception:
            dest_tw = None
        prev = getattr(self, '_drag_preview', None)
        if prev is None:
            from PyQt5.QtWidgets import QLabel
            prev = QLabel(None)
            prev.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint)
            prev.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            prev.setAttribute(Qt.WA_ShowWithoutActivating, True)
            self._drag_preview = prev
        cross = (dest_tw is not None and dest_tw is not self._owner_tw())
        border = "#4a90d9" if cross else "#b0b0b0"
        title = getattr(self, '_press_title', '') or tr("标签页")
        prefix = "\u21a6 " if cross else ""  # ↦ 表示将转移到另一组
        prev.setStyleSheet(
            f"QLabel {{ background: #ffffff; border: 1px solid {border};"
            f" border-radius: 4px; padding: 4px 10px; color: #202020; font-size: 12px; }}"
        )
        prev.setText(prefix + title)
        prev.adjustSize()
        prev.move(int(gpos.x()) + 14, int(gpos.y()) + 14)
        if not prev.isVisible():
            prev.show()
        prev.raise_()

    def _hide_drag_preview(self):
        prev = getattr(self, '_drag_preview', None)
        if prev is not None and prev.isVisible():
            prev.hide()

    def mouseReleaseEvent(self, event):
        was_dragging = self._dragging
        press_content = self._press_content
        press_index = self._press_index
        self._press_content = None
        self._press_index = -1
        self._press_pos = None
        self._dragging = False
        self._hide_drag_preview()
        super().mouseReleaseEvent(event)

        if not was_dragging or press_content is None or event.button() != Qt.LeftButton:
            # 普通点击（未触发拖拽）：清理拖拽标志即可
            if self._is_dragging_tab:
                self._is_dragging_tab = False
                if self.main_window is not None:
                    self.main_window._tab_drag_in_progress = False
            return

        # 确定落点目标组与插入位置
        dest_tw, dest_index = None, -1
        try:
            if self.main_window is not None and hasattr(self.main_window, '_pane_group_hit_test'):
                dest_tw, dest_index = self.main_window._pane_group_hit_test(event.globalPos())
        except Exception:
            dest_tw, dest_index = None, -1
        owner = self._owner_tw()
        moved = False
        if dest_tw is not None and dest_tw is not owner:
            # 跨组转移
            try:
                moved = self.main_window.move_tab_across_groups(
                    owner, press_content, dest_tw, dest_index)
            except Exception as _e:
                debug_print(f"[CrossGroupDrag] transfer failed: {_e}")
        elif dest_tw is owner:
            # 组内重排：按光标位置计算目标索引，move 到该位置（tabMoved → on_tab_moved 同步内容栈）
            try:
                target = self.tabAt(self.mapFromGlobal(event.globalPos()))
                if target < 0:
                    target = self.count() - 1
                if 0 <= press_index < self.count() and target != press_index:
                    self.moveTab(press_index, target)
                    moved = True
            except Exception as _e:
                debug_print(f"[TabReorder] failed: {_e}")
        # dest_tw 为 None（释放在任何标签组之外，如标题栏）→ 视为取消，不移动

        # 结束拖拽：恢复正常刷新状态
        self._is_dragging_tab = False
        if self.main_window is not None:
            self.main_window._tab_drag_in_progress = False
        if not moved:
            # 未发生移动：补一次当前组的 tab_changed，恢复正常刷新
            try:
                self.main_window._on_group_tab_changed(owner, owner.currentIndex())
            except Exception:
                pass
    
    def event(self, event):
        # 拦截所有事件，确保双击事件能被处理
        if event.type() == QEvent.MouseButtonDblClick:
            debug_print(f"[DEBUG] TabBar event: MouseButtonDblClick")
            self.mouseDoubleClickEvent(event)
            return True
        return super().event(event)
    
    def mouseDoubleClickEvent(self, event):
        debug_print(f"[DEBUG] TabBar double click event triggered")
        # 获取点击位置
        pos = event.pos()
        # 判断是否点在空白区域（没有点在任何标签页上）
        clicked_tab = self.tabAt(pos)
        debug_print(f"[DEBUG] Clicked tab: {clicked_tab}, pos: ({pos.x()}, {pos.y()}), count: {self.count()}")
        
        # 如果点击在空白区域，或点击在最后一个标签右侧的空白处
        is_blank = clicked_tab == -1
        if not is_blank and self.count() > 0:
            # 检查是否点击在最后一个标签页的右侧
            last_tab_rect = self.tabRect(self.count() - 1)
            debug_print(f"[DEBUG] Last tab right edge: {last_tab_rect.right()}")
            if pos.x() > last_tab_rect.right():
                is_blank = True
        
        debug_print(f"[DEBUG] Is blank area: {is_blank}, has main_window: {self.main_window is not None}")
        
        if is_blank:
            # 点击在空白区域，打开新标签页（归属当前标签组）
            if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                debug_print(f"[DEBUG] Opening new tab from TabBar...")
                self.main_window.add_new_tab(target_tabwidget=self._owner_tw())
                event.accept()
                return
        
        # 如果点击在标签页上，调用默认行为
        super().mouseDoubleClickEvent(event)
    
    def leaveEvent(self, event):
        """鼠标离开标签栏"""
        self.hovered_tab = -1
        super().leaveEvent(event)
    
    def _make_close_btn(self, index):
        """创建关闭按钮，始终可见，hover 时变灰"""
        close_btn = QToolButton(self)
        close_btn.setText("×")
        tab_close_btn_size = int(16 * getattr(self, 'parent_window', self).dpi_scale if hasattr(getattr(self, 'parent_window', self), 'dpi_scale') else 16)
        close_btn.setFixedSize(tab_close_btn_size, tab_close_btn_size)
        close_btn.setStyleSheet("""
            QToolButton {
                border: none;
                background: transparent;
                color: #999;
                font-size: 18px;
                font-weight: bold;
                padding: 0px;
                margin: 0px;
            }
            QToolButton:hover {
                background: #cccccc;
                color: #333;
                border-radius: 4px;
            }
        """)
        def _on_click(_checked=False, btn=close_btn, tabbar=self):
            # 通过按钮在 TabBar 中的位置反查标签索引，避免 PyQt 对象包装导致的身份比较失效
            tab_index = tabbar.tabAt(btn.mapToParent(QPoint(1, 1)))
            if tab_index < 0:
                tab_index = tabbar.tabAt(btn.pos())
            if tab_index >= 0 and tabbar.main_window and hasattr(tabbar.main_window, 'close_tab'):
                tabbar.main_window.close_tab(tab_index, target_tabwidget=tabbar._owner_tw())
        close_btn.clicked.connect(_on_click)
        return close_btn

    def tabInserted(self, index):
        """标签添加时自动附加关闭按钮"""
        super().tabInserted(index)
        close_btn = self._make_close_btn(index)
        self.setTabButton(index, QTabBar.RightSide, close_btn)

    def close_tab_at_index(self, index):
        """关闭指定索引的标签页"""
        if self.main_window and hasattr(self.main_window, 'close_tab'):
            self.main_window.close_tab(index)

    def show_close_button_under_cursor(self):
        """关闭标签后重建所有按钮（索引已变化）"""
        for i in range(self.count()):
            if self.tabButton(i, QTabBar.RightSide) is None:
                close_btn = self._make_close_btn(i)
                self.setTabButton(i, QTabBar.RightSide, close_btn)
    
    def on_tab_moved(self, from_index, to_index):
        """标签页移动后的处理，同步所属组的 content_stack；固定标签逻辑仅适用于左侧组。"""
        if not self.main_window:
            return
        debug_print(f"[TabMoved] Moving tab from {from_index} to {to_index}")
        # 标记拖拽进行中，通知 on_tab_changed 跳过重型操作
        self._is_dragging_tab = True
        self.main_window._tab_drag_in_progress = True
        self._drag_end_timer.start(150)  # 重置去抖计时器
        tw = self._owner_tw()
        # 同步移动所属组 content_stack 中的对应内容
        content_stack = self.main_window._content_stack_for(tw)
        if content_stack is not None:
            moved_widget = content_stack.widget(from_index)
            if moved_widget:
                content_stack.removeWidget(moved_widget)
                content_stack.insertWidget(to_index, moved_widget)
                debug_print(f"[TabMoved] Synced content_stack: moved widget from {from_index} to {to_index}")
        # 移动后自动检测鼠标下的tab并显示关闭按钮
        self.show_close_button_under_cursor()
        # 固定标签纠正仅适用于左侧主标签组
        if tw is not self.main_window.tab_widget:
            return
        # 获取被移动的标签页
        moved_tab = self.main_window.tab_widget.widget(to_index)
        if not moved_tab:
            return
        is_pinned = getattr(moved_tab, 'is_pinned', False)
        pinned_count = 0
        for i in range(self.count()):
            tab = self.main_window.tab_widget.widget(i)
            if tab and getattr(tab, 'is_pinned', False):
                pinned_count += 1
        
        # 如果是固定标签页移动到非固定区域，或非固定标签页移动到固定区域，需要纠正
        if is_pinned and to_index >= pinned_count:
            # 固定标签页不能移动到非固定区域，移回固定区域末尾
            self.moveTab(to_index, pinned_count - 1)
        elif not is_pinned and to_index < pinned_count - 1:
            # 非固定标签页不能移动到固定区域，移到非固定区域开头
            self.moveTab(to_index, pinned_count)


# ─────────────────────────────────────────────────────────────────────────────
# AI 聊天面板：ChatWorker（异步 API 线程）+ ChatPanel（UI 面板）
# ─────────────────────────────────────────────────────────────────────────────


class ChatWorker(QThread):
    """在后台线程中发起 OpenAI 兼容 API 请求，避免阻塞 UI。支持流式输出。"""
    response_received = pyqtSignal(str)
    token_received = pyqtSignal(str)   # 流式逐块 token
    error_occurred = pyqtSignal(str)

    def __init__(self, messages, api_url, api_key, model, stream=True, parent=None):
        super().__init__(parent)
        self.messages = messages
        self.api_url = api_url.rstrip('/')
        self.api_key = api_key
        self.model = model
        self.stream = stream

    def run(self):
        try:
            import requests, json as _json
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}" if self.api_key else "",
            }
            payload = {
                "model": self.model,
                "messages": self.messages,
                "stream": self.stream,
            }
            url = self.api_url + '/chat/completions'
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            if self.stream:
                try:
                    resp = requests.post(url, headers=headers, json=payload,
                                         timeout=120, verify=False, stream=True)
                    resp.raise_for_status()
                except requests.exceptions.ConnectionError as _ce:
                    if 'prematurely' in str(_ce).lower() or 'protocol' in str(_ce).lower():
                        self.error_occurred.emit(
                            tr('❌ 流式响应被服务器提前关闭（上下文可能过长）。\n') + str(_ce)
                        )
                        return
                    raise
                full_content = []
                try:
                    for raw_line in resp.iter_lines():
                        if self.isInterruptionRequested():
                            break
                        if not raw_line:
                            continue
                        line = raw_line.decode('utf-8', errors='replace') if isinstance(raw_line, bytes) else raw_line
                        if not line.startswith('data: '):
                            continue
                        data_str = line[6:].strip()
                        if data_str == '[DONE]':
                            break
                        try:
                            data = _json.loads(data_str)
                            chunk = ((data.get('choices') or [{}])[0]
                                     .get('delta', {}).get('content')) or ''
                            if chunk:
                                full_content.append(chunk)
                                self.token_received.emit(chunk)
                        except Exception:
                            pass
                except Exception as _se:
                    _msg = str(_se)
                    if 'prematurely' in _msg.lower() or 'protocol' in _msg.lower():
                        partial = ''.join(full_content)
                        if partial:
                            # 已收到部分内容，仍然返回
                            self.response_received.emit(partial)
                        else:
                            self.error_occurred.emit(
                                tr('❌ 流式响应被服务器提前关闭（上下文可能过长，请清空聊天记录重试）。\n') + _msg
                            )
                        return
                    raise
                self.response_received.emit(''.join(full_content))
            else:
                resp = requests.post(url, headers=headers, json=payload, timeout=120, verify=False)
                resp.raise_for_status()
                data = resp.json()
                content = data['choices'][0]['message']['content']
                self.response_received.emit(content)
        except Exception as e:
            self.error_occurred.emit(str(e))

class ChatPanel(QWidget):
    """右侧 AI 聊天侧边栏面板。"""

    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.messages = []   # 对话历史（不含 system prompt）
        self.worker = None
        self.history_file = get_app_data_path("chat_history.json")
        self._is_loading_history = False
        # 给 ChatPanel 独立的 Win32 HWND，避免 QAxWidget(Shell Explorer) 抢占鼠标/键盘头
        self.setAttribute(Qt.WA_NativeWindow, True)
        self.setFocusPolicy(Qt.ClickFocus)
        # 启用拖拽
        self.setAcceptDrops(True)
        self._setup_ui()
        self._load_history()  # 启动时加载聊天记录

    # ── UI 构建 ──────────────────────────────────────────────────────────────
    def _setup_ui(self):
        from PyQt5.QtWidgets import QTextBrowser, QPlainTextEdit
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # 标题栏
        header = QWidget()
        header.setStyleSheet("background: #E3F2FD; border-radius: 4px;")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(8, 4, 8, 4)
        title_lbl = QLabel(tr("🤖 AI 助手"))
        title_lbl.setStyleSheet("font-weight: bold; font-size: 10.5pt; background: transparent;")
        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        clear_btn = QPushButton(tr("清空"))
        clear_btn.setFixedHeight(22)
        clear_btn.setStyleSheet(
            "QPushButton{background:#fff;border:1px solid #90CAF9;border-radius:3px;font-size:9pt;padding:0 6px;}"
            "QPushButton:hover{background:#BBDEFB;}"
        )
        clear_btn.clicked.connect(self.clear_chat)
        header_layout.addWidget(clear_btn)
        layout.addWidget(header)

        # 当前目录提示条
        self.context_label = QLabel(tr("当前目录: —"))
        self.context_label.setStyleSheet(
            "color:#555; font-size:8.5pt; padding:2px 6px;"
            "background:#f0f4f8; border-radius:3px;"
        )
        self.context_label.setWordWrap(True)
        layout.addWidget(self.context_label)

        # 聊天历史显示区
        self.chat_display = QTextBrowser()
        self.chat_display.setOpenExternalLinks(False)
        self.chat_display.setTextInteractionFlags(
            Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard | Qt.TextBrowserInteraction
        )
        self.chat_display.setStyleSheet(
            "QTextBrowser{background:#fafafa;border:1px solid #e0e0e0;"
            "border-radius:4px;font-family:'Microsoft YaHei UI','Segoe UI',Arial;"
            "font-size:9.5pt;}"
        )
        self.chat_display.setContextMenuPolicy(Qt.CustomContextMenu)
        self.chat_display.customContextMenuRequested.connect(self._show_context_menu)
        layout.addWidget(self.chat_display, 1)

        # 输入框
        self.input_box = QPlainTextEdit()
        self.input_box.setPlaceholderText(tr("输入问题… (Enter 发送，Shift+Enter 换行)"))
        self.input_box.setFixedHeight(72)
        self.input_box.setReadOnly(False)
        self.input_box.setFocusPolicy(Qt.StrongFocus)
        self.input_box.setAttribute(Qt.WA_InputMethodEnabled, True)
        self.input_box.setStyleSheet(
            "QPlainTextEdit{border:1px solid #d0d0d0;border-radius:4px;padding:4px;"
            "font-family:'Microsoft YaHei UI','Segoe UI',Arial;font-size:9.5pt;}"
            "QPlainTextEdit:focus{border:1px solid #42A5F5;}"
        )
        self.input_box.installEventFilter(self)
        # QPlainTextEdit 的鼠标事件实际落在 viewport 上，这里一并监听
        self.input_box.viewport().installEventFilter(self)
        layout.addWidget(self.input_box)

        # 发送按钮行
        btn_row = QHBoxLayout()
        self.send_btn = QPushButton(tr("发 送"))
        self.send_btn.setFixedHeight(28)
        self.send_btn.setStyleSheet(
            "QPushButton{background:#1976D2;color:white;border:none;border-radius:4px;"
            "font-size:9.5pt;font-weight:bold;padding:0 16px;}"
            "QPushButton:hover{background:#1565C0;}"
            "QPushButton:disabled{background:#BDBDBD;}"
        )
        self.send_btn.clicked.connect(self.send_message)
        btn_row.addStretch()
        btn_row.addWidget(self.send_btn)
        layout.addLayout(btn_row)

        # 状态提示
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color:#888;font-size:8.5pt;")
        layout.addWidget(self.status_lbl)

    # ── 鼠标点击：确保焦点转移到输入框 ──────────────────────────────────────
    def mousePressEvent(self, event):
        """点击面板空白区域时，强制 input_box 获取焦点（解决 QAxWidget 抢焦点问题）。"""
        super().mousePressEvent(event)
        self.main_window.activateWindow()
        self.input_box.setFocus(Qt.MouseFocusReason)

    # ── 事件过滤（Enter 发送 / 鼠标点击夺回焦点）────────────────────────────
    def eventFilter(self, obj, event):
        from PyQt5.QtCore import QEvent, QTimer
        if obj is self.input_box or obj is self.input_box.viewport():
            if event.type() == QEvent.KeyPress:
                if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                    if not (event.modifiers() & Qt.ShiftModifier):
                        self.send_message()
                        return True
            elif event.type() == QEvent.MouseButtonPress:
                # input_box 被点击时，确保主窗口重新激活并把焦点交给输入框
                # 用 singleShot(0) 让 Qt 先处理完点击事件再设焦点
                self.main_window.activateWindow()
                QTimer.singleShot(0, lambda: self.input_box.setFocus(Qt.MouseFocusReason))
                QTimer.singleShot(20, lambda: self.input_box.setFocus(Qt.MouseFocusReason))
        return super().eventFilter(obj, event)

    # ── 公共方法 ─────────────────────────────────────────────────────────────
    def update_context(self, path: str):
        """更新当前目录提示。"""
        if path:
            display = path if len(path) <= 55 else '…' + path[-52:]
            self.context_label.setText(tr("当前目录: {}").format(display))

    def _show_context_menu(self, pos):
        """显示右键菜单（复制、全选等）。"""
        menu = QMenu(self)
        
        copy_action = menu.addAction(tr("复制"))
        copy_action.triggered.connect(self.chat_display.copy)
        
        select_all_action = menu.addAction(tr("全选"))
        select_all_action.triggered.connect(self.chat_display.selectAll)
        
        menu.addSeparator()
        clear_action = menu.addAction(tr("清空聊天"))
        clear_action.triggered.connect(self.clear_chat)
        
        menu.exec_(self.chat_display.mapToGlobal(pos))

    def clear_chat(self):
        self.messages.clear()
        self.chat_display.clear()
        self._delete_history()  # 清空时删除保存文件

    _MAX_AGENTIC_STEPS = 0  # 0 = 无限制；>0 = 硬性轮数上限

    def _cleanup_worker(self):
        # 用 sender() 避免与 agentic loop 新建的 worker 混淆
        finished_worker = self.sender()
        target = finished_worker if finished_worker else self.worker
        if self.worker is target:
            self.worker = None
        if target:
            try:
                target.deleteLater()
            except Exception:
                pass

    def cleanup(self):
        try:
            self.input_box.removeEventFilter(self)
        except Exception:
            pass
        try:
            self.input_box.viewport().removeEventFilter(self)
        except Exception:
            pass

        worker = self.worker
        self.worker = None
        if worker:
            for signal_name, handler in (
                ('response_received', self._on_response),
                ('error_occurred', self._on_error),
                ('finished', self._cleanup_worker),
                ('token_received', self._on_token),
            ):
                try:
                    getattr(worker, signal_name).disconnect(handler)
                except Exception:
                    pass

            try:
                worker.requestInterruption()
            except Exception:
                pass

            if worker.isRunning():
                try:
                    worker.finished.connect(worker.deleteLater)
                except Exception:
                    pass
                try:
                    worker.setParent(None)
                except Exception:
                    pass
            else:
                try:
                    worker.deleteLater()
                except Exception:
                    pass

    def closeEvent(self, event):
        self.cleanup()
        super().closeEvent(event)

    def _sanitize_chat_message(self, message):
        if not isinstance(message, dict):
            return None
        role = str(message.get("role", "system") or "system")
        if role not in ("system", "user", "assistant"):
            role = "system"
        content = str(message.get("content", "") or "")
        if len(content) > MAX_CHAT_MESSAGE_CHARS:
            omitted = len(content) - MAX_CHAT_MESSAGE_CHARS
            content = content[:MAX_CHAT_MESSAGE_CHARS] + tr("\n\n【已截断 {} 个字符】").format(omitted)
        return {"role": role, "content": content}

    def _trim_chat_history(self):
        trimmed = []
        for message in self.messages[-MAX_CHAT_HISTORY_MESSAGES:]:
            sanitized = self._sanitize_chat_message(message)
            if sanitized is not None:
                trimmed.append(sanitized)
        # 总字符预算：超出时从最老的消息开始丢弃，至少保留最新 1 条
        while len(trimmed) > 1:
            if sum(len(m.get('content', '')) for m in trimmed) <= MAX_CONTEXT_TOTAL_CHARS:
                break
            trimmed.pop(0)
        self.messages = trimmed

    @staticmethod
    def _md_to_html(text: str) -> str:
        """Convert common Markdown patterns to HTML for chat display."""
        import html as _html
        import re as _re

        def inline_fmt(s):
            # s is already HTML-escaped; apply inline formatting
            parts = _re.split(r'(`[^`]+`)', s)
            result = []
            for part in parts:
                if part.startswith('`') and part.endswith('`') and len(part) > 1:
                    inner = part[1:-1]
                    result.append(
                        f'<code style="background:#f0f0f0;padding:1px 4px;'
                        f'border-radius:3px;font-family:Consolas,monospace;font-size:12px">'
                        f'{inner}</code>'
                    )
                else:
                    p = _re.sub(r'\*\*\*(.+?)\*\*\*', r'<b><i>\1</i></b>', part)
                    p = _re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', p)
                    p = _re.sub(r'\*(.+?)\*', r'<i>\1</i>', p)
                    result.append(p)
            return ''.join(result)

        lines = text.split('\n')
        out = []
        i = 0
        while i < len(lines):
            line = lines[i]

            # --- Fenced code block ---
            if line.startswith('```'):
                lang = _html.escape(line[3:].strip())
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].startswith('```'):
                    code_lines.append(_html.escape(lines[i]))
                    i += 1
                code = '\n'.join(code_lines)
                lang_html = (
                    f'<div style="color:#aaa;font-size:11px;margin-bottom:3px">{lang}</div>'
                    if lang else ''
                )
                out.append(
                    f'<div style="background:#1e1e1e;color:#d4d4d4;padding:8px 10px;margin:6px 0;'
                    f'border-radius:4px;font-family:Consolas,monospace;font-size:12px;'
                    f'white-space:pre;overflow-x:auto">{lang_html}{code}</div>'
                )
                i += 1
                continue

            # --- Horizontal rule ---
            if _re.match(r'^[-*_]{3,}\s*$', line):
                out.append('<hr style="border:0;border-top:1px solid #ccc;margin:8px 0">')
                i += 1
                continue

            # --- Heading ---
            m = _re.match(r'^(#{1,6})\s+(.+)', line)
            if m:
                level = len(m.group(1))
                sizes = {1: '18px', 2: '16px', 3: '14px', 4: '13px'}
                size = sizes.get(level, '13px')
                border = 'border-bottom:1px solid #ccc;padding-bottom:3px;' if level <= 2 else ''
                content = inline_fmt(_html.escape(m.group(2)))
                out.append(
                    f'<div style="font-size:{size};font-weight:bold;'
                    f'margin:10px 0 4px 0;{border}">{content}</div>'
                )
                i += 1
                continue

            # --- Table ---
            if line.strip().startswith('|') and line.strip().endswith('|'):
                tbl_lines = []
                while i < len(lines) and lines[i].strip().startswith('|') and lines[i].strip().endswith('|'):
                    tbl_lines.append(lines[i])
                    i += 1
                rows = []
                sep_idx = -1
                for ri, tl in enumerate(tbl_lines):
                    cells = [c.strip() for c in tl.strip()[1:-1].split('|')]
                    if all(_re.match(r'^:?-+:?$', c) for c in cells if c):
                        sep_idx = ri
                    else:
                        rows.append((ri, cells))
                if rows:
                    tbl = ['<table style="border-collapse:collapse;width:100%;margin:6px 0;font-size:12px">']
                    header_rows = [r for r in rows if sep_idx < 0 or r[0] < sep_idx]
                    data_rows   = [r for r in rows if sep_idx >= 0 and r[0] > sep_idx]
                    if header_rows:
                        tbl.append('<thead>')
                        for _, cells in header_rows:
                            tbl.append('<tr>')
                            for cell in cells:
                                tbl.append(
                                    f'<th style="padding:4px 8px;border:1px solid #ccc;'
                                    f'background:#e8e8e8;text-align:left">'
                                    f'{inline_fmt(_html.escape(cell))}</th>'
                                )
                            tbl.append('</tr></thead>')
                    if data_rows:
                        tbl.append('<tbody>')
                        for _, cells in data_rows:
                            tbl.append('<tr>')
                            for cell in cells:
                                tbl.append(
                                    f'<td style="padding:4px 8px;border:1px solid #ccc">'
                                    f'{inline_fmt(_html.escape(cell))}</td>'
                                )
                            tbl.append('</tr>')
                        tbl.append('</tbody>')
                    tbl.append('</table>')
                    out.append(''.join(tbl))
                continue

            # --- Unordered list ---
            if _re.match(r'^\s*[-*+]\s+', line):
                out.append('<ul style="margin:2px 0;padding-left:20px">')
                while i < len(lines) and _re.match(r'^\s*[-*+]\s+', lines[i]):
                    m2 = _re.match(r'^\s*[-*+]\s+(.+)', lines[i])
                    out.append(f'<li>{inline_fmt(_html.escape(m2.group(1)))}</li>')
                    i += 1
                out.append('</ul>')
                continue

            # --- Ordered list ---
            if _re.match(r'^\s*\d+\.\s+', line):
                out.append('<ol style="margin:2px 0;padding-left:20px">')
                while i < len(lines) and _re.match(r'^\s*\d+\.\s+', lines[i]):
                    m2 = _re.match(r'^\s*\d+\.\s+(.+)', lines[i])
                    out.append(f'<li>{inline_fmt(_html.escape(m2.group(1)))}</li>')
                    i += 1
                out.append('</ol>')
                continue

            # --- Empty line ---
            if not line.strip():
                out.append('<div style="height:4px"></div>')
                i += 1
                continue

            # --- Normal line ---
            out.append(f'<div>{inline_fmt(_html.escape(line))}</div>')
            i += 1

        return ''.join(out)

    def _build_bubble_html(self, role: str, content: str):
        import html as _html
        if role == "user":
            color, prefix, bg = "#1976D2", tr("👤 你"), "#E3F2FD"
        elif role == "assistant":
            color, prefix, bg = "#2E7D32", "🤖 AI", "#F1F8E9"
        else:
            color, prefix, bg = "#888", tr("ℹ 系统"), "#F5F5F5"

        if role == "assistant":
            body = self._md_to_html(content)
        else:
            body = _html.escape(content).replace('\n', '<br>')
        return (
            f'<div style="margin:4px 0;padding:6px 10px;background:{bg};'
            f'border-radius:6px;border-left:3px solid {color};">'
            f'<b style="color:{color};">{prefix}</b><br>{body}</div>'
        )

    def _rebuild_chat_display(self):
        self.chat_display.clear()
        for message in self.messages:
            self.chat_display.append(self._build_bubble_html(message.get("role", "system"), message.get("content", "")))
        self.chat_display.verticalScrollBar().setValue(
            self.chat_display.verticalScrollBar().maximum()
        )

    # 气泡过多时 QTextDocument 不会自动截断——超过此块数则重建显示
    _MAX_CHAT_DISPLAY_BLOCKS = 400

    def append_bubble(self, role: str, content: str, save_history=True):
        """向聊天区追加一条消息气泡（HTML 格式）。"""
        bubble = self._build_bubble_html(role, content)
        self.chat_display.append(bubble)
        self.chat_display.verticalScrollBar().setValue(
            self.chat_display.verticalScrollBar().maximum()
        )
        # 防止 QTextDocument 无限增长：超块数时重建（已有 _rebuild_chat_display 函数）
        if self.chat_display.document().blockCount() > self._MAX_CHAT_DISPLAY_BLOCKS:
            self._rebuild_chat_display()
        if save_history and not self._is_loading_history:
            self._schedule_save_history()  # 防抖写盘，避免每条气泡都触发 I/O

    def _schedule_save_history(self):
        """聊天记录防抖写盘（500ms 内多次追加只写一次磁盘）。"""
        if not hasattr(self, '_history_save_timer') or self._history_save_timer is None:
            from PyQt5.QtCore import QTimer
            self._history_save_timer = QTimer(self)
            self._history_save_timer.setSingleShot(True)
            self._history_save_timer.timeout.connect(self._save_history)
        self._history_save_timer.start(500)

    # ── 发送消息 ─────────────────────────────────────────────────────────────
    def send_message(self):
        text = self.input_box.toPlainText().strip()
        if not text:
            return

        # 检查用户输入是否包含文件路径或链接，提前处理
        self._handle_user_file_input(text)

        cfg = self.main_window.config.get("ai_chat", {})
        api_url = cfg.get("api_url", "").strip()
        api_key = cfg.get("api_key", "").strip()
        model   = cfg.get("model", "gpt-3.5-turbo").strip() or "gpt-3.5-turbo"

        if not api_url:
            self.append_bubble("system", tr("❌ 请先在 设置 → AI 助手 中填写 API 地址"))
            return

        self.input_box.clear()
        self.append_bubble("user", text)

        # 获取当前路径作为上下文
        current_path = ""
        try:
            tab = self.main_window.get_current_tab_widget()
            if tab and hasattr(tab, 'current_path'):
                current_path = tab.current_path or ""
        except Exception:
            pass

        # 构建 system prompt
        system_prompt = cfg.get("system_prompt", "").strip()
        if not system_prompt:
            system_prompt = (
                tr("你是一个智能文件管理助手，帮助用户管理文件系统、打开目录和运行脚本。\n") +
                tr("⚠️ 只有当用户明确要求‘切换/打开目录’时，才在回复末尾添加：[OPEN_DIR: 目录完整路径]。") +
                tr("普通问答不要输出任何操作指令。\n") +
                tr("⚠️ 只有当用户明确要求运行脚本时，才添加：[RUN_SCRIPT: 脚本完整路径]。\n") +
                tr("⚠️ 仅当用户明确要求时，才可使用以下指令：\n") +
                tr("[READ_FILE: 文件路径] 读取文件（系统自动分段读取大文件，无需手动指定偏移）；\n") +
                tr("[PATCH_FILE: 路径|旧文本|新文本] 局部修改已有文件（安全，仅替换目标内容段）；\n") +
                tr("⚠️ 修改已有文件时必须使用 PATCH_FILE，禁止用 WRITE_FILE 覆盖已有大文件；\n") +
                tr("⚠️ PATCH_FILE 使用规范：\n") +
                tr("  1. 旧文本只需包含被修改的行及前后各1-2行（足够唯一定位即可），不要复制整个函数体；\n") +
                tr("  2. 对同一文件做多个 PATCH_FILE 时，第二个 patch 的旧文本必须是第一个 patch 应用后的实际内容；\n") +
                tr("  3. 最小改动原则：新文本只改动必要的行，不要重写周围未变更的代码。\n") +
                tr("[WRITE_FILE: 文件路径|文件内容] 仅用于创建全新文件（已有文件禁止使用）；\n") +
                tr("[LIST_DIR: 目录路径] 列出目录；\n") +
                tr("[MKDIR: 目录路径] 创建目录；\n") +
                tr("[DELETE: 路径] 删除文件或目录（需用户确认）。\n") +
                tr("路径使用 Windows 格式，例如 D:\\project\\src。\n") +
                tr("可以在同一回复中包含多个操作命令。")
            )

        ctx_prefix = tr("[当前目录: {}]\n").format(current_path) if current_path else ""
        full_messages = [{"role": "system", "content": system_prompt}]
        full_messages.extend(self.messages)
        full_messages.append({"role": "user", "content": ctx_prefix + text})

        # 保存对话历史（不含 system，不含 ctx_prefix）
        self.messages.append({"role": "user", "content": text})
        self._trim_chat_history()

        # 启动工作线程（流式输出，每个 token 实时更新状态栏）
        self._agentic_steps = 0
        self._last_agentic_feed = None  # 重复调用检测用
        self._streaming_chars = 0
        self.send_btn.setEnabled(False)
        self.status_lbl.setText(tr("⏳ AI 思考中…"))
        self.worker = ChatWorker(full_messages, api_url, api_key, model, stream=True, parent=self)
        self.worker.token_received.connect(self._on_token)
        self.worker.response_received.connect(self._on_response)
        self.worker.error_occurred.connect(self._on_error)
        self.worker.finished.connect(self._cleanup_worker)
        self.worker.start()

    def _handle_user_file_input(self, text):
        """检查用户输入中是否包含文件链接，直接打开。仅处理显式的 file:// 链接，避免误触发。"""
        import re
        # 只匹配 file:/// 或 file:// 链接（显式指示）
        file_links = re.findall(r'file:///?([^\s\]]+)', text)
        for link in file_links:
            path = link.replace('%20', ' ').replace('%5C', '\\')
            if os.path.isfile(path):
                # 打开文件所在目录
                folder = os.path.dirname(path)
                if os.path.isdir(folder):
                    self.append_bubble("system", tr("📂 打开目录: {}").format(folder))
                    self.main_window.add_new_tab(folder)
            elif os.path.isdir(path):
                # 直接打开目录
                self.append_bubble("system", tr("📂 打开目录: {}").format(path))
                self.main_window.add_new_tab(path)

    def _on_token(self, chunk: str):
        """流式 token 到达——实时更新聊天气泡（VS Code Copilot 效果）。"""
        self._streaming_buffer = getattr(self, '_streaming_buffer', '') + chunk
        step = getattr(self, '_agentic_steps', 0)
        label = tr("第 {} 轮").format(step + 1) if step > 0 else ""
        self.status_lbl.setText(tr("⏳ AI 推理中{}…").format(label))

        if not getattr(self, '_streaming_bubble_open', False):
            # 第一个 token：记录插入位置，后续刷新替换这段内容
            from PyQt5.QtGui import QTextCursor
            cursor = QTextCursor(self.chat_display.document())
            cursor.movePosition(QTextCursor.End)
            self._streaming_pre_char_pos = cursor.position()
            self._streaming_bubble_open = True

        # 50ms 防抖刷新（约 20fps），避免每 token 都重建 HTML
        if not hasattr(self, '_streaming_flush_timer') or self._streaming_flush_timer is None:
            from PyQt5.QtCore import QTimer
            self._streaming_flush_timer = QTimer(self)
            self._streaming_flush_timer.setSingleShot(True)
            self._streaming_flush_timer.timeout.connect(self._flush_streaming_bubble)
        self._streaming_flush_timer.start(50)

    def _flush_streaming_bubble(self):
        """将累积的 token 刷新到聊天气泡（替换式更新，保持样式和光标）。"""
        if not getattr(self, '_streaming_bubble_open', False):
            return
        from PyQt5.QtGui import QTextCursor
        doc = self.chat_display.document()
        cursor = QTextCursor(doc)
        cursor.setPosition(self._streaming_pre_char_pos)
        cursor.movePosition(QTextCursor.End, QTextCursor.KeepAnchor)
        cursor.removeSelectedText()
        buf = getattr(self, '_streaming_buffer', '')
        cursor.insertHtml(self._build_bubble_html("assistant", buf + "▌"))
        self.chat_display.verticalScrollBar().setValue(
            self.chat_display.verticalScrollBar().maximum()
        )

    def _close_streaming_bubble(self):
        """移除流式临时气泡（由 _on_response / _on_error 调用，后续追加最终气泡）。"""
        if not getattr(self, '_streaming_bubble_open', False):
            return
        if hasattr(self, '_streaming_flush_timer') and self._streaming_flush_timer:
            self._streaming_flush_timer.stop()
        from PyQt5.QtGui import QTextCursor
        doc = self.chat_display.document()
        cursor = QTextCursor(doc)
        cursor.setPosition(self._streaming_pre_char_pos)
        cursor.movePosition(QTextCursor.End, QTextCursor.KeepAnchor)
        cursor.removeSelectedText()
        self._streaming_bubble_open = False
        self._streaming_buffer = ""

    def _on_response(self, content: str):
        self.status_lbl.setText("")
        self._streaming_chars = 0
        self.messages.append({"role": "assistant", "content": content})
        self._trim_chat_history()
        display_content, feedable = self._apply_actions(content)

        if feedable:
            # 中间轮次：不追加气泡，清空缓冲区让流式气泡原地刷新为“等待”状态
            last_feed = getattr(self, '_last_agentic_feed', None)
            stuck = (last_feed is not None and feedable == last_feed)
            self._last_agentic_feed = feedable
            if stuck:
                self.status_lbl.setText(tr("⚠️ 检测到重复工具调用，强制要求给出结论…"))
            if hasattr(self, '_streaming_flush_timer') and self._streaming_flush_timer:
                self._streaming_flush_timer.stop()
            self._streaming_buffer = ""
            if getattr(self, '_streaming_bubble_open', False):
                self._flush_streaming_bubble()
            self._start_agentic_step(feedable, force_final=stuck)
        else:
            # 最终结果：关闭流式气泡，追加正式气泡
            self._close_streaming_bubble()
            self.append_bubble("assistant", display_content)
            self._last_agentic_feed = None
            self.send_btn.setEnabled(True)
    def _start_agentic_step(self, feedable_results: list, force_final: bool = False):
        """将工具调用结果回传给 AI，启动下一轮自动推理。
        force_final=True 时禁止 AI 继续调工具，强制给出结论。
        """
        cfg = self.main_window.config.get("ai_chat", {})
        api_url = cfg.get("api_url", "").strip()
        api_key = cfg.get("api_key", "").strip()
        model   = cfg.get("model", "gpt-3.5-turbo").strip() or "gpt-3.5-turbo"
        if not api_url:
            self.send_btn.setEnabled(True)
            return

        self._agentic_steps = getattr(self, '_agentic_steps', 0) + 1
        step = self._agentic_steps

        # 工具结果限制单条 8000 字，防止上下文爆炸
        _MAX_FEED = 8000
        feed_parts = [r[:_MAX_FEED] + (f"\n…（已截断，共 {len(r)} 字）" if len(r) > _MAX_FEED else "")
                      for r in feedable_results]
        feed_text = "\n\n".join(feed_parts)

        if force_final:
            continuation = (
                tr("以上是最后一批工具调用结果。不要再调用任何工具指令，") +
                tr("请直接基于已收集的全部信息给出具体的结论、优化建议或解决方案。")
            )
        else:
            continuation = (
                tr("以上是工具调用结果。请高效利用工具，") +
                tr("优先阅读源代码文件（.c/.h）而非反复列目录，") +
                tr("并尽快基于收集到的信息给出具体结论或优化建议。")
            )
        feed_msg = {"role": "user", "content": feed_text + "\n\n" + continuation}

        # 保存进历史（下一轮 AI 能看到工具结果）
        self.messages.append(feed_msg)
        self._trim_chat_history()

        # ── Agentic 续推专用 system prompt ─────────────────────────────────
        # 与初始 send_message 的提示词分开：这里鼓励 AI 自由使用工具，
        # 而不是重复“只有用户明确要求时才使用”的限制性提示词。
        system_prompt = cfg.get("system_prompt", "").strip()
        if not system_prompt:
            if force_final:
                system_prompt = (
                    tr("你是一个智能文件管理助手。") +
                    tr("不能将任何工具指令放入回复，") +
                    tr("必须直接基于已收集的所有信息给出具体的结论、优化建议或解决方案。") +
                    tr("路径使用 Windows 格式。")
                )
            else:
                system_prompt = (
                    tr("你是一个智能文件管理助手，正在通过工具调用完成用户任务。\n") +
                    tr("你可以自由使用以下工具指令：\n") +
                    tr("[READ_FILE: 路径] 读文件（优先读 .c/.h 源代码）；") +
                    tr("[PATCH_FILE: 路径|旧文本|新文本] 局部修改文件（修改已有文件时必须用此，禁止用 WRITE_FILE 覆盖已有大文件）；") +
                    tr("[LIST_DIR: 路径] 列目录（仅在不知道源码位置时使用）；") +
                    tr("[OPEN_DIR: 路径] [MKDIR: 路径] [WRITE_FILE: 路径|内容]（仅用于创建新文件） [DELETE: 路径]。\n") +
                    tr("ℹ️ 工作原则：得到目录列表后尽快选择源代码文件直接阅读，") +
                    tr("而不要反复展开子目录；修改文件时优先使用 PATCH_FILE 安全替换具体代码段。\n") +
                    tr("⚠️ PATCH_FILE 最小改动规范：\n") +
                    tr("  1. 旧文本只取被修改行及前后各1-2行（能唯一定位即可），不要复制整个函数；\n") +
                    tr("  2. 对同一文件连续多个 PATCH_FILE 时，后一个的旧文本必须反映前一个 patch 已应用后的文件内容；\n") +
                    tr("  3. 新文本只改动必要的行，保留其余未变行，使 diff 最小化。\n") +
                    tr("路径使用 Windows 格式，例如 D:\\project\\src。")
                )
        full_messages = [{"role": "system", "content": system_prompt}]
        full_messages.extend(self.messages)

        self._streaming_chars = 0
        self.status_lbl.setText(tr("⏳ AI 第 {} 轮推理中…").format(step))
        self.worker = ChatWorker(full_messages, api_url, api_key, model, stream=True, parent=self)
        self.worker.token_received.connect(self._on_token)
        self.worker.response_received.connect(self._on_response)
        self.worker.error_occurred.connect(self._on_error)
        self.worker.finished.connect(self._cleanup_worker)
        self.worker.start()

    def _on_error(self, msg: str):
        self._close_streaming_bubble()   # 移除未完成的流式气泡
        self.send_btn.setEnabled(True)
        self.status_lbl.setText("")
        self.append_bubble("system", tr("❌ 请求失败: {}").format(msg))

    # ── 拖拽文件支持 ──────────────────────────────────────────────────────────
    def dragEnterEvent(self, event):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dragMoveEvent(self, event):
        """拖拽移动事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        """拖拽释放事件 - 读取文件内容"""
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    self._read_and_display_file(file_path)
                else:
                    # 尝试从 URL 字符串中提取本地路径
                    url_str = url.toString()
                    if url_str.startswith('file:///'):
                        from urllib.parse import unquote
                        file_path = unquote(url_str[8:])
                        if os.name == 'nt' and file_path.startswith('/'):
                            file_path = file_path[1:]
                        self._read_and_display_file(file_path)
            event.acceptProposedAction()
    
    def _read_and_display_file(self, file_path):
        """读取文件并在聊天窗口中显示"""
        if not os.path.isfile(file_path):
            self.append_bubble("system", tr("❌ 文件不存在: {}").format(file_path))
            return
        
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        if file_size_mb > 10:  # 超过10MB不读取
            self.append_bubble("system", tr("⚠️ 文件太大（{:.1f}MB），无法读取").format(file_size_mb))
            return
        
        try:
            # 尝试以 UTF-8 编码读取
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            try:
                # 降级为 GBK/GB2312
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()
            except Exception:
                try:
                    # 最后尝试二进制显示
                    with open(file_path, 'rb') as f:
                        raw = f.read(1000)
                    content = f"[二进制文件，前 {len(raw)} 字节]\n{raw[:200]}"
                except Exception as e:
                    self.append_bubble("system", tr("❌ 无法读取文件: {}").format(e))
                    return
        
        # 截断长内容（超过5000字符）
        max_chars = 5000
        if len(content) > max_chars:
            content = content[:max_chars] + f"\n\n【省略 {len(content) - max_chars} 个字符】"
        
        # 获取文件名
        file_name = os.path.basename(file_path)
        
        # 显示文件内容
        display_text = f"📄 {file_name}\n" + "="*50 + "\n" + content
        self.append_bubble("system", display_text)

    # ── 聊天记录持久化 ──────────────────────────────────────────────────────────
    def _save_history(self):
        """保存聊天记录到 JSON 文件。"""
        import json
        try:
            before_len = len(self.messages)
            self._trim_chat_history()
            if len(self.messages) != before_len:
                self._rebuild_chat_display()
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.messages, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(tr("保存聊天记录失败: {}").format(e))

    def _load_history(self):
        """从 JSON 文件加载聊天记录。"""
        import json
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    loaded_messages = json.load(f)
                self.messages = loaded_messages if isinstance(loaded_messages, list) else []
                self._trim_chat_history()
                # 重新显示所有消息
                self._is_loading_history = True
                self._rebuild_chat_display()
                self._is_loading_history = False
                self._save_history()
            except Exception as e:
                self._is_loading_history = False
                print(tr("加载聊天记录失败: {}").format(e))

    def _delete_history(self):
        """删除保存的聊天记录文件。"""
        try:
            if os.path.exists(self.history_file):
                os.remove(self.history_file)
        except Exception as e:
            print(tr("删除聊天记录文件失败: {}").format(e))

    def _get_action_base_dir(self):
        """获取动作执行的基准目录（当前标签目录）。"""
        try:
            tab = self.main_window.get_current_tab_widget()
            p = getattr(tab, 'current_path', '') if tab else ''
            if p and os.path.isdir(p):
                return os.path.normpath(p)
        except Exception:
            pass
        return os.path.normpath(get_app_base_dir())

    def _resolve_action_path(self, raw_path: str):
        """解析动作路径，限制在当前目录范围内。"""
        p = (raw_path or '').strip().strip('"\'')
        if not p:
            return None, tr("空路径")

        p = p.replace('/', '\\')
        base_dir = self._get_action_base_dir()

        # 相对路径按当前目录解析
        if not os.path.isabs(p):
            p = os.path.join(base_dir, p)

        p = translate_common_path(os.path.normpath(p))

        try:
            base_norm = os.path.normcase(os.path.normpath(base_dir))
            path_norm = os.path.normcase(os.path.normpath(p))
            if not (path_norm == base_norm or path_norm.startswith(base_norm + os.sep)):
                return None, tr("超出当前目录范围: {}").format(p)
        except Exception:
            return None, tr("路径非法: {}").format(p)

        return p, None

    def _confirm_danger_action(self, title: str, text: str) -> bool:
        """危险操作二次确认。"""
        try:
            from PyQt5.QtWidgets import QMessageBox
            ret = QMessageBox.question(
                self,
                title,
                text,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            return ret == QMessageBox.Yes
        except Exception:
            # 任何异常都按拒绝处理，确保安全
            return False

    # ── 解析并执行 AI 操作命令 ────────────────────────────────────────────────
    def _apply_actions(self, content: str) -> tuple:
        """解析 AI 回复中的操作标记并按出现顺序执行。
        返回 (display_content, feedable)：
        - display_content: 附带执行注释的显示文本
        - feedable: READ_FILE/LIST_DIR 结果列表，供 agentic loop 回传给 AI
        """
        import re, subprocess
        notes = []

        # ── 收集所有指令及其在文本中的位置，按顺序执行 ──────────────────────
        feedable = []  # 可回传给 AI 的工具结果（READ_FILE / LIST_DIR）
        actions = []  # list of (start_pos, action_type, match_obj)

        patterns = {
            'OPEN_DIR':   re.compile(r'\[OPEN_DIR:\s*([^\]]+)\]'),
            'RUN_SCRIPT': re.compile(r'\[RUN_SCRIPT:\s*([^\]]+)\]'),
            'LIST_DIR':   re.compile(r'\[LIST_DIR:\s*([^\]]+)\]'),
            'READ_FILE':  re.compile(r'\[READ_FILE:\s*([^\]]+)\]'),
            'MKDIR':      re.compile(r'\[MKDIR:\s*([^\]]+)\]'),
            # PATCH_FILE / WRITE_FILE 内容可含代码中的 ] (如 arr[i])，
            # 用 .*? + 行末 lookahead 确保只在真正的指令结束处停止
            'PATCH_FILE': re.compile(r'\[PATCH_FILE:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'WRITE_FILE': re.compile(r'\[WRITE_FILE:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'DELETE':     re.compile(r'\[DELETE:\s*([^\]]+)\]'),
        }
        for atype, pat in patterns.items():
            for m in pat.finditer(content):
                actions.append((m.start(), atype, m))

        # 按出现位置排序，保证 WRITE_FILE 在 RUN_SCRIPT 之前（若 AI 如此安排）
        actions.sort(key=lambda x: x[0])

        for _, atype, m in actions:

            # ── OPEN_DIR ────────────────────────────────────────────────────
            if atype == 'OPEN_DIR':
                raw = m.group(1).strip().strip('"\'')
                path = translate_common_path(os.path.normpath(raw))
                if os.path.isdir(path):
                    try:
                        current_tab = self.main_window.get_current_tab_widget()
                        if current_tab and hasattr(current_tab, 'navigate_to'):
                            current_path = getattr(current_tab, 'current_path', '')
                            same_path = False
                            if hasattr(self.main_window, '_normalize_path_for_compare'):
                                same_path = (
                                    self.main_window._normalize_path_for_compare(current_path)
                                    == self.main_window._normalize_path_for_compare(path)
                                )
                            if same_path:
                                notes.append(tr("ℹ 当前标签已在该目录: {}").format(path))
                            else:
                                current_tab.navigate_to(path, skip_async_check=True)
                                notes.append(tr("✅ 已在当前标签切换到目录: {}").format(path))
                        else:
                            self.main_window.add_new_tab(path)
                            notes.append(tr("✅ 已打开目录: {}").format(path))
                    except Exception as e:
                        notes.append(tr("❌ 打开目录失败: {}（{}）").format(path, e))
                else:
                    notes.append(tr("⚠️ 目录不存在: {}").format(path))

            # ── RUN_SCRIPT ──────────────────────────────────────────────────
            elif atype == 'RUN_SCRIPT':
                raw = m.group(1).strip().strip('"\'')
                script = os.path.normpath(raw)
                if os.path.isfile(script):
                    try:
                        if not self._confirm_danger_action(
                            tr("确认运行脚本"),
                            tr("AI 请求运行脚本：\n{}\n\n是否继续？").format(script)
                        ):
                            notes.append(tr("⏸ 已取消运行脚本: {}").format(script))
                            continue
                        launch_detached(script if isinstance(script, list) else [script], cwd=os.path.dirname(script))
                        notes.append(tr("✅ 已启动脚本: {}").format(script))
                    except Exception as e:
                        notes.append(tr("❌ 运行脚本失败: {}（{}）").format(script, e))
                else:
                    notes.append(tr("⚠️ 脚本不存在: {}").format(script))

            # ── LIST_DIR ────────────────────────────────────────────────────
            elif atype == 'LIST_DIR':
                raw = m.group(1)
                path, err = self._resolve_action_path(raw)
                if err:
                    _emsg = tr("❌ 列目录失败: {}").format(err)
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                if not os.path.isdir(path):
                    _emsg = tr("❌ 目录不存在: {}").format(path)
                    if os.path.isfile(path):
                        _emsg += tr("（这是文件，请用 READ_FILE）")
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                try:
                    items = os.listdir(path)
                    preview = "\n".join(items[:50]) if items else tr("(空目录)")
                    if len(items) > 50:
                        preview += f"\n... 其余 {len(items) - 50} 项省略"
                    notes.append(tr("📁 目录列表: {}\n{}").format(path, preview))
                    feedable.append(tr("[LIST_DIR 结果] 目录 {}:\n{}").format(path, preview))
                except Exception as e:
                    notes.append(tr("❌ 列目录失败: {}（{}）").format(path, e))

            # ── READ_FILE ───────────────────────────────────────────────────
            elif atype == 'READ_FILE':
                raw = m.group(1)
                parts = raw.split('|')
                raw_path = parts[0].strip()
                start_offset = 0
                if len(parts) > 1:
                    try:
                        start_offset = int(parts[1].strip())
                    except ValueError:
                        pass
                path, err = self._resolve_action_path(raw_path)
                if err:
                    _emsg = tr("❌ 读取失败: {}").format(err)
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                if not os.path.isfile(path):
                    _emsg = tr("❌ 文件不存在: {}").format(path)
                    if os.path.isdir(path):
                        _emsg += tr("（这是目录，请用 LIST_DIR 列目录，或直接指定具体 .c/.h 文件路径）")
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                CHUNK = 5000
                MAX_TOTAL = 50000
                try:
                    enc = 'utf-8'
                    try:
                        with open(path, 'r', encoding='utf-8') as _probe_f:
                            _probe_f.read(1)
                    except UnicodeDecodeError:
                        enc = 'gbk'
                    file_size = os.path.getsize(path)
                    all_text = []
                    offset = start_offset
                    with open(path, 'r', encoding=enc, errors='replace') as f:
                        f.seek(offset)
                        while True:
                            chunk = f.read(CHUNK)
                            if not chunk:
                                break
                            all_text.append(chunk)
                            offset += len(chunk.encode(enc, errors='replace'))
                            if sum(len(t) for t in all_text) >= MAX_TOTAL:
                                break
                    full_text = ''.join(all_text)
                    total_read = sum(len(t) for t in all_text)
                    truncated = (start_offset + total_read) < file_size
                    info = tr("📄 文件内容（{}-{}/{}字节，{}）: {}").format(start_offset, start_offset+total_read, file_size, enc, path)
                    if truncated:
                        info += tr("\n⚠️ 文件过大，仅读取前 {} 字符，剩余 {} 字节未读").format(MAX_TOTAL, file_size - start_offset - total_read)
                    notes.append(f"{info}\n{full_text}")
                    feedable.append(f"{info}\n{full_text}")
                except Exception as e:
                    notes.append(tr("❌ 读取失败: {}（{}）").format(path, e))

            # ── MKDIR ───────────────────────────────────────────────────────
            elif atype == 'MKDIR':
                raw = m.group(1)
                path, err = self._resolve_action_path(raw)
                if err:
                    notes.append(tr("❌ 创建目录失败: {}").format(err))
                    continue
                try:
                    os.makedirs(path, exist_ok=True)
                    notes.append(tr("✅ 已创建目录: {}").format(path))
                except Exception as e:
                    notes.append(tr("❌ 创建目录失败: {}（{}）").format(path, e))

            # ── PATCH_FILE ──────────────────────────────────────────────────
            elif atype == 'PATCH_FILE':
                raw = m.group(1)
                parts = raw.split('|', 2)
                if len(parts) < 3:
                    notes.append(tr("❌ PATCH_FILE 格式: [PATCH_FILE: 路径|旧文本|新文本]"))
                    continue
                raw_path, old_text, new_text = parts
                path, err = self._resolve_action_path(raw_path.strip())
                if err:
                    notes.append(tr("❌ 补丁失败: {}").format(err))
                    continue
                if not os.path.isfile(path):
                    notes.append(tr("❌ 文件不存在: {}").format(path))
                    continue
                try:
                    enc = 'utf-8'
                    try:
                        with open(path, 'r', encoding='utf-8') as _f: _f.read(1)
                    except UnicodeDecodeError:
                        enc = 'gbk'
                    with open(path, 'r', encoding=enc, errors='replace') as _f:
                        file_src = _f.read()
                    count = file_src.count(old_text)
                    if count == 0:
                        notes.append(tr("❌ 补丁失败：在 {} 中未找到目标文本").format(path))
                        continue
                    if count > 1:
                        notes.append(f"⚠️ 目标文本在 {os.path.basename(path)} 中出现 {count} 次，仅替换第一处")
                    patched = file_src.replace(old_text, new_text, 1)
                    with open(path, 'w', encoding=enc) as _f:
                        _f.write(patched)
                    notes.append(tr("✅ 补丁成功: {}").format(path))
                except Exception as e:
                    notes.append(tr("❌ 补丁失败: {}（{}）").format(path, e))

            # ── WRITE_FILE ──────────────────────────────────────────────────
            elif atype == 'WRITE_FILE':
                raw = m.group(1)
                if '|' not in raw:
                    notes.append(tr("❌ 写入失败: 格式应为 [WRITE_FILE: 路径|内容]"))
                    continue
                raw_path, file_content = raw.split('|', 1)
                path, err = self._resolve_action_path(raw_path)
                if err:
                    notes.append(tr("❌ 写入失败: {}").format(err))
                    continue
                try:
                    file_exists = os.path.exists(path)
                    # 已有文件且较大时，拒绝全量覆盖，引导用 PATCH_FILE
                    if file_exists and os.path.getsize(path) > 5000:
                        notes.append(
                            f"❌ 已拒绝覆盖：{os.path.basename(path)} 已存在且大于 5KB，" +
                            tr("全量覆盖会丢失未读内容。") +
                            tr("请改用 [PATCH_FILE: 路径|旧文本|新文本] 进行局部修改。")
                        )
                        continue
                    if file_exists:
                        if not self._confirm_danger_action(
                            tr("确认覆盖文件"),
                            tr("AI 请求覆盖文件：\n{}\n\n是否继续？").format(path)
                        ):
                            notes.append(tr("⏸ 已取消覆盖文件: {}").format(path))
                            continue
                    parent = os.path.dirname(path)
                    if parent and not os.path.exists(parent):
                        os.makedirs(parent, exist_ok=True)
                    with open(path, 'w', encoding='utf-8') as f:
                        f.write(file_content)
                    notes.append(tr("✅ 已写入文件: {}").format(path))
                except Exception as e:
                    notes.append(tr("❌ 写入失败: {}（{}）").format(path, e))

            # ── DELETE ──────────────────────────────────────────────────────
            elif atype == 'DELETE':
                raw = m.group(1).strip()
                path, err = self._resolve_action_path(raw)
                if err:
                    notes.append(tr("❌ 删除失败: {}").format(err))
                    continue
                if not os.path.exists(path):
                    notes.append(tr("❌ 路径不存在: {}").format(path))
                    continue
                is_dir = os.path.isdir(path)
                type_label = tr("目录（及其所有内容）") if is_dir else tr("文件")
                if not self._confirm_danger_action(
                    tr("确认删除"),
                    tr("AI 请求删除{}：\n{}\n\n⚠️ 此操作不可撤销！是否继续？").format(type_label, path)
                ):
                    notes.append(tr("⏸ 已取消删除: {}").format(path))
                    continue
                try:
                    if is_dir:
                        import shutil
                        shutil.rmtree(path)
                    else:
                        os.remove(path)
                    notes.append(tr("✅ 已删除{}: {}").format(type_label, path))
                except Exception as e:
                    notes.append(tr("❌ 删除失败: {}（{}）").format(path, e))

        if notes:
            content = content + "\n\n" + "\n".join(notes)
        return content, feedable


class TitleShortcutBar(QWidget):
    """标题栏快捷方式区域：支持拖拽常用启动文件并点击运行。"""

    shortcutDropped = pyqtSignal(str)
    shortcutClicked = pyqtSignal(str)
    shortcutsChanged = pyqtSignal(list)

    def __init__(self, parent=None, icon_size=18, button_size=28):
        super().__init__(parent)
        self._icon_size = icon_size
        self._button_size = button_size
        self._paths = []
        self._icon_cache = {}  # path → QIcon (avoids re-calling SHGetFileInfo)
        self._icon_provider = QFileIconProvider()
        self._drag_start_pos = None
        self._drag_source_index = -1
        self.setAcceptDrops(True)

        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(8)

        self._hint_label = QLabel(tr("拖入应用或快捷方式"))
        self._hint_label.setStyleSheet("QLabel { color: #666; font-size: 9pt; padding: 0 4px; }")
        self._hint_label.setAcceptDrops(True)
        self._hint_label.installEventFilter(self)
        self._layout.addWidget(self._hint_label)

    def set_shortcuts(self, paths):
        cleaned = []
        for path in paths or []:
            if isinstance(path, str) and path and path not in cleaned:
                cleaned.append(path)
        self._paths = cleaned[:20]
        self._rebuild_buttons()

    def get_shortcuts(self):
        return list(self._paths)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-tabex-shortcut-index"):
            event.acceptProposedAction()
            return
        if self._collect_shortcuts_from_mime(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat("application/x-tabex-shortcut-index"):
            event.acceptProposedAction()
            return
        if self._collect_shortcuts_from_mime(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasFormat("application/x-tabex-shortcut-index"):
            try:
                source_index = int(bytes(event.mimeData().data("application/x-tabex-shortcut-index")).decode("utf-8"))
            except Exception:
                event.ignore()
                return
            target_index = self._get_insert_index(event.pos().x())
            self._reorder_shortcut(source_index, target_index)
            event.acceptProposedAction()
            return

        paths = self._collect_shortcuts_from_mime(event.mimeData())
        if not paths:
            event.ignore()
            return
        for path in paths:
            self.shortcutDropped.emit(path)
        event.acceptProposedAction()

    def eventFilter(self, obj, event):
        if isinstance(obj, QToolButton) and obj.parent() is self:
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                path = obj.property("shortcut_path")
                if path in self._paths:
                    self._drag_start_pos = event.globalPos()
                    self._drag_source_index = self._paths.index(path)
                return False
            if event.type() == QEvent.MouseMove and (event.buttons() & Qt.LeftButton):
                if self._drag_start_pos is not None and self._drag_source_index >= 0:
                    if (event.globalPos() - self._drag_start_pos).manhattanLength() >= QApplication.startDragDistance():
                        source_index = self._drag_source_index
                        self._drag_start_pos = None
                        self._drag_source_index = -1
                        self._start_drag_from_index(source_index)
                        return True
                return False
            if event.type() == QEvent.MouseButtonRelease:
                self._drag_start_pos = None
                self._drag_source_index = -1
                return False

        if event.type() in (QEvent.DragEnter, QEvent.DragMove):
            mime_data = event.mimeData() if hasattr(event, 'mimeData') else None
            if mime_data and (mime_data.hasFormat("application/x-tabex-shortcut-index") or self._collect_shortcuts_from_mime(mime_data)):
                event.acceptProposedAction()
                return True
        elif event.type() == QEvent.Drop:
            mime_data = event.mimeData() if hasattr(event, 'mimeData') else None
            if not mime_data:
                return super().eventFilter(obj, event)

            if mime_data.hasFormat("application/x-tabex-shortcut-index"):
                try:
                    source_index = int(bytes(mime_data.data("application/x-tabex-shortcut-index")).decode("utf-8"))
                except Exception:
                    event.ignore()
                    return True
                local_pos = obj.mapTo(self, event.pos()) if isinstance(obj, QWidget) else event.pos()
                target_index = self._get_insert_index(local_pos.x())
                self._reorder_shortcut(source_index, target_index)
                event.acceptProposedAction()
                return True

            paths = self._collect_shortcuts_from_mime(mime_data)
            if paths:
                for path in paths:
                    self.shortcutDropped.emit(path)
                event.acceptProposedAction()
                return True

        return super().eventFilter(obj, event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = event.pos()
            self._drag_source_index = self._button_index_at(event.pos())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.LeftButton):
            super().mouseMoveEvent(event)
            return
        if self._drag_start_pos is None or self._drag_source_index < 0:
            super().mouseMoveEvent(event)
            return
        if (event.pos() - self._drag_start_pos).manhattanLength() < QApplication.startDragDistance():
            super().mouseMoveEvent(event)
            return

        source_index = self._drag_source_index
        self._drag_start_pos = None
        self._drag_source_index = -1
        self._start_drag_from_index(source_index)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_start_pos = None
        self._drag_source_index = -1
        super().mouseReleaseEvent(event)

    def _collect_shortcuts_from_mime(self, mime_data):
        if not mime_data or not mime_data.hasUrls():
            return []
        paths = []
        for url in mime_data.urls():
            if not url.isLocalFile():
                continue
            path = url.toLocalFile()
            if not path:
                continue
            if is_supported_title_shortcut_path(path):
                paths.append(path)
        return paths

    def _resolve_shortcut_icon(self, path):
        if not isinstance(path, str) or not path:
            return QIcon()

        lower_path = path.lower()

        # For .exe files, use ExtractIconEx (fast, no shell timeout)
        if lower_path.endswith('.exe') and os.name == 'nt':
            icon = self._extract_icon_fast(path)
            if icon and not icon.isNull():
                return icon
            return self._icon_provider.icon(QFileInfo(path))

        if lower_path.endswith('.lnk') and os.name == 'nt':
            try:
                from win32com.client import Dispatch

                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(path)

                icon_location = getattr(shortcut, 'IconLocation', '') or ''
                if icon_location:
                    icon_path = icon_location.split(',', 1)[0].strip().strip('"')
                    icon_path = os.path.expandvars(icon_path)
                    if os.path.exists(icon_path):
                        icon = self._extract_icon_fast(icon_path) if icon_path.lower().endswith('.exe') else None
                        if not icon or icon.isNull():
                            icon = self._icon_provider.icon(QFileInfo(icon_path))
                        if not icon.isNull():
                            return icon

                target_path = getattr(shortcut, 'Targetpath', '') or getattr(shortcut, 'TargetPath', '') or ''
                target_path = os.path.expandvars(str(target_path).strip().strip('"'))
                if target_path and os.path.exists(target_path):
                    icon = self._extract_icon_fast(target_path) if target_path.lower().endswith('.exe') else None
                    if not icon or icon.isNull():
                        icon = self._icon_provider.icon(QFileInfo(target_path))
                    if not icon.isNull():
                        return icon
            except Exception:
                pass

        return self._icon_provider.icon(QFileInfo(path))

    @staticmethod
    def _extract_icon_fast(exe_path):
        """Extract icon from .exe using ExtractIconExW (fast, no shell timeout)."""
        try:
            import ctypes
            import ctypes.wintypes
            from PyQt5.QtGui import QPixmap
            from PyQt5.QtWinExtras import QtWin

            _ExtractIconExW = ctypes.windll.shell32.ExtractIconExW
            _ExtractIconExW.argtypes = [ctypes.c_wchar_p, ctypes.c_int,
                                        ctypes.POINTER(ctypes.wintypes.HICON),
                                        ctypes.POINTER(ctypes.wintypes.HICON),
                                        ctypes.c_uint]
            _ExtractIconExW.restype = ctypes.c_uint
            _DestroyIcon = ctypes.windll.user32.DestroyIcon

            hicon_large = ctypes.wintypes.HICON()
            hicon_small = ctypes.wintypes.HICON()
            count = _ExtractIconExW(exe_path, 0,
                                    ctypes.byref(hicon_large),
                                    ctypes.byref(hicon_small), 1)
            if count and hicon_large.value:
                try:
                    pixmap = QtWin.fromHICON(hicon_large.value)
                    if not pixmap.isNull():
                        return QIcon(pixmap)
                finally:
                    _DestroyIcon(hicon_large)
                    if hicon_small.value:
                        _DestroyIcon(hicon_small)
            elif hicon_small.value:
                _DestroyIcon(hicon_small)
        except Exception:
            pass
        return None

    def _rebuild_buttons(self):
        while self._layout.count() > 0:
            item = self._layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        # 提示文字固定在最左侧
            self._hint_label = QLabel(tr("+ 拖入应用或快捷方式"))
        self._hint_label.setStyleSheet("QLabel { color: #888; font-size: 9pt; padding: 0 4px; }")
        self._hint_label.setAcceptDrops(True)
        self._hint_label.installEventFilter(self)
        self._layout.addWidget(self._hint_label)

        if not self._paths:
            return

        self._layout.addSpacing(6)

        for path in self._paths:
            btn = QToolButton(self)
            btn.setAutoRaise(True)
            btn.setFixedSize(self._button_size, self._button_size)
            icon_px = max(12, min(self._icon_size + 1, self._button_size - 4))
            btn.setIconSize(QSize(icon_px, icon_px))
            btn.setToolTip(f"{os.path.basename(path)}\n{path}\n左键启动应用/快捷方式，右键移除，拖拽可排序")
            # Use cached icon or defer loading to avoid SHGetFileInfo blocking startup
            cached_icon = self._icon_cache.get(path)
            if cached_icon is not None:
                if not cached_icon.isNull():
                    btn.setIcon(cached_icon)
                else:
                    btn.setText("↗")
            else:
                btn.setText("…")
                # Defer icon loading well after startup to avoid SHGetFileInfo blocking
                QTimer.singleShot(1500, lambda b=btn, p=path: self._load_icon_deferred(b, p))
            btn.setStyleSheet(
                "QToolButton { background: transparent; border: none; border-radius: 4px; padding: 0px; margin: 0px; }"
                "QToolButton:hover { background: #e0e0e0; }"
                "QToolButton:pressed { background: #d0d0d0; }"
            )
            btn.setAcceptDrops(True)
            btn.installEventFilter(self)
            btn.setProperty("shortcut_path", path)
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda pos, b=btn: self._show_context_menu_for_button(b, pos))
            btn.clicked.connect(lambda _checked=False, p=path: self.shortcutClicked.emit(p))
            self._layout.addWidget(btn)

    def _load_icon_deferred(self, btn, path):
        """Load shortcut icon asynchronously and update button."""
        try:
            if not btn or not btn.isVisible():
                return
            icon = self._resolve_shortcut_icon(path)
            self._icon_cache[path] = icon
            if not icon.isNull():
                btn.setIcon(icon)
                btn.setText("")
            else:
                btn.setText("↗")
        except RuntimeError:
            pass  # Button already deleted

    def _show_context_menu_for_button(self, btn, pos):
        path = btn.property("shortcut_path")
        if not path:
            return
        menu = QMenu(self)
        menu.setStyleSheet(
            """
            QMenu {
                background-color: #ffffff;
                border: 1px solid #c7c7c7;
                border-radius: 6px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 24px 6px 12px;
                background: transparent;
                border-radius: 4px;
                color: #303030;
                margin: 2px 4px;
            }
            QMenu::item:selected {
                background: #3f5f7a;
                color: #f6f8fa;
            }
            QMenu::item:pressed {
                background: #324c62;
                color: #ffffff;
            }
            """
        )
        remove_action = menu.addAction(tr("移除该快捷方式"))
        action = menu.exec_(btn.mapToGlobal(pos))
        if action == remove_action:
            self._remove_shortcut(path)

    def _remove_shortcut(self, path):
        if path not in self._paths:
            return
        self._paths = [p for p in self._paths if p != path]
        self._rebuild_buttons()
        self.shortcutsChanged.emit(list(self._paths))

    def _button_index_at(self, pos):
        widget = self.childAt(pos)
        while widget and widget is not self:
            if isinstance(widget, QToolButton):
                path = widget.property("shortcut_path")
                if path in self._paths:
                    return self._paths.index(path)
            widget = widget.parentWidget()
        return -1

    def _get_insert_index(self, x_pos):
        button_centers = []
        for i in range(self._layout.count()):
            item = self._layout.itemAt(i)
            widget = item.widget()
            if isinstance(widget, QToolButton):
                g = widget.geometry()
                button_centers.append((g.left() + g.right()) // 2)
        if not button_centers:
            return 0
        for idx, center in enumerate(button_centers):
            if x_pos < center:
                return idx
        return len(button_centers)

    def _reorder_shortcut(self, source_index, target_index):
        if source_index < 0 or source_index >= len(self._paths):
            return
        if target_index > source_index:
            target_index -= 1
        if target_index < 0:
            target_index = 0
        if target_index >= len(self._paths):
            target_index = len(self._paths) - 1
        if target_index == source_index:
            return
        moving = self._paths.pop(source_index)
        self._paths.insert(target_index, moving)
        self._rebuild_buttons()
        self.shortcutsChanged.emit(list(self._paths))

    def _start_drag_from_index(self, source_index):
        if source_index < 0 or source_index >= len(self._paths):
            return
        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setData("application/x-tabex-shortcut-index", str(source_index).encode("utf-8"))
        drag.setMimeData(mime_data)
        drag.exec_(Qt.MoveAction)


from PyQt5.QtWidgets import QSplitterHandle as _QSplitterHandle


class _ResizableSplitterHandle(_QSplitterHandle):
    """分割条手柄：作为原生窗口并强制显示左右调整光标，并在中央绘制可见抓取条。

    内嵌的 IExplorerBrowser 是原生窗口（HWND），夹在两个原生面板之间的普通（alien）
    手柄常常收不到鼠标悬停、光标不会变成调整光标。将手柄本身设为原生窗口可与相邻的
    Shell 原生窗口正确竞争命中测试，确保能拖动并显示左右调整光标。
    """

    def __init__(self, orientation, parent):
        super().__init__(orientation, parent)
        # 原生窗口：解决夹在原生 Shell 窗口之间时光标/命中测试失效的问题
        self.setAttribute(Qt.WA_NativeWindow, True)
        self.setCursor(Qt.SplitHCursor if orientation == Qt.Horizontal else Qt.SplitVCursor)
        self._hovered = False

    def enterEvent(self, event):
        self._hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hovered = False
        self.update()
        super().leaveEvent(event)

    def paintEvent(self, event):
        from PyQt5.QtGui import QPainter, QColor
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#aab2c0") if self._hovered else QColor("#d2d7e0"))
        # 中央竖向抓取条（三段短竖线），提升可识别度
        cx = self.width() // 2
        cy = self.height() // 2
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#6b7686"))
        if self.orientation() == Qt.Horizontal:
            for dy in (-12, 0, 12):
                p.drawRect(cx - 1, cy + dy - 6, 2, 12)
        else:
            for dx in (-12, 0, 12):
                p.drawRect(cx + dx - 6, cy - 1, 12, 2)
        p.end()


class ResizableSplitter(QSplitter):
    """使用更宽、带原生窗口、明确光标与可见抓取条的手柄，便于在内嵌原生 Shell 窗口旁拖动调整宽度。"""

    def createHandle(self):
        return _ResizableSplitterHandle(self.orientation(), self)


class MainWindow(QMainWindow):
    def _is_control_panel_path_for_monitor(self, path):
        """判断路径是否为控制面板或其子目录（供Explorer Monitor用）"""
        if not path:
            return False
        s = path.lower()
        if s.startswith('shell:controlpanelfolder'):
            return True
        if '::{26ee0668-a00a-44d7-9371-beb064c98683}' in s:
            return True
        if s.startswith('control panel') or s.startswith('control panel/') or s.startswith('control panel\\'):
            return True
        if 'control.exe' in s:
            return True
        if s.startswith('shell:::{26ee0668-a00a-44d7-9371-beb064c98683}'):
            return True
        if '/control panel/' in s or '\\control panel\\' in s:
            return True
        return False

    # 定义信号用于从服务器线程通知主线程打开新标签
    open_path_signal = pyqtSignal(str)

    def ensure_default_icons_on_bookmark_bar(self):
        """确保四个常用书签（带图标）始终在最左侧且不会被覆盖。"""
        bm = self.bookmark_manager
        tree = bm.get_tree()
        bar = tree.get('bookmark_bar')
        if not bar or 'children' not in bar:
            return
        import time
        from PyQt5.QtCore import QStandardPaths
        downloads_path = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
        if not downloads_path or not os.path.exists(downloads_path):
            downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
        icon_map = [
            ("🖥️", tr("此电脑"), "shell:MyComputerFolder"),
            ("🗔", "桌面", "shell:Desktop"),
            ("🗑️", tr("回收站"), "shell:RecycleBinFolder"),
            ("⬇️", "下载", downloads_path),
        ]
        names_set = set([n for _, n, _ in icon_map])
        bar['children'] = [c for c in bar['children'] if not (c.get('type') == 'url' and any(c.get('name', '').replace(icon, '').strip() == n for icon, n, _ in icon_map))]
        now = int(time.time() * 1000000)
        def make_bm(icon, name, url):
            nonlocal now
            now += 1
            return {
                "date_added": str(now),
                "id": str(now),
                "name": f"{icon} {name}",
                "type": "url",
                "url": url
            }
        bar['children'] = [make_bm(icon, name, url) for icon, name, url in icon_map] + bar['children']
        bm.save_bookmarks()

    def tabbar_mouse_double_click(self, event):
        tabbar = self.tab_widget.tabBar()
        pos = event.pos()
        # 判断是否点在tab右侧空白区（包括tabbar宽度范围内和超出tab的区域）
        if tabbar.tabAt(pos) == -1 or pos.x() > tabbar.tabRect(tabbar.count() - 1).right():
            self.add_new_tab()
    
    def get_tab_widget(self, index):
        """获取指定索引的实际标签页内容（从content_stack）"""
        if hasattr(self, 'content_stack') and index >= 0 and index < self.content_stack.count():
            return self.content_stack.widget(index)
        return self.tab_widget.widget(index)
    
    def get_current_tab_widget(self):
        """获取当前标签页的实际内容（从content_stack）"""
        current_index = self.tab_widget.currentIndex()
        return self.get_tab_widget(current_index)

    def _resolve_group(self, target_tabwidget):
        """根据目标标签控件返回 (tab_widget, content_stack, is_right)。默认左侧主组。"""
        if target_tabwidget is not None and target_tabwidget is getattr(self, 'split_tab_widget', None):
            return self.split_tab_widget, self.split_content_stack, True
        return self.tab_widget, self.content_stack, False

    def _content_stack_for(self, target_tabwidget):
        """返回目标标签组对应的 content_stack。"""
        _tw, cs, _is_right = self._resolve_group(target_tabwidget)
        return cs

    def _all_groups(self):
        """返回当前存在的所有标签组 (tab_widget, content_stack) 列表（左侧主组 + 右侧分屏组）。"""
        groups = [(self.tab_widget, self.content_stack)]
        stw = getattr(self, 'split_tab_widget', None)
        scs = getattr(self, 'split_content_stack', None)
        if stw is not None and scs is not None:
            groups.append((stw, scs))
        return groups

    def _pane_group_hit_test(self, gpos):
        """判断全局坐标 gpos 落在哪个标签组，用于跨组拖拽落点。

        以各组内容区(content_stack)的水平范围界定归属：命中内容区 → (tab_widget, -1) 追加；
        命中内容区正上方的标签栏/书签栏行 → (tab_widget, 插入索引)。分屏激活时右侧组优先，
        未命中 → (None, None)。用内容区宽度界定标签栏归属，不依赖窄窄的 tabBar 实际宽度，命中更宽松。"""
        from PyQt5.QtCore import QPoint, QRect
        groups = []
        if getattr(self, '_split_active', False):
            stw = getattr(self, 'split_tab_widget', None)
            scs = getattr(self, 'split_content_stack', None)
            if stw is not None and scs is not None:
                groups.append((stw, scs))
        groups.append((self.tab_widget, self.content_stack))
        for tw, cs in groups:
            try:
                if cs is None or not cs.isVisible():
                    continue
                cs_tl = cs.mapToGlobal(QPoint(0, 0))
                cs_rect = QRect(cs_tl, cs.size())
                # 内容区：追加到末尾
                if cs_rect.contains(gpos):
                    return tw, -1
                # 内容区正上方（标签栏/书签栏行）：按该组内容的水平范围归属，使拖放命中更容易
                if cs_rect.left() <= gpos.x() <= cs_rect.right() and gpos.y() < cs_rect.top():
                    bar = tw.tabBar() if tw is not None else None
                    insert_idx = bar.tabAt(bar.mapFromGlobal(gpos)) if bar is not None else -1
                    return tw, insert_idx
            except Exception:
                continue
        return None, None

    def move_tab_across_groups(self, source_tabwidget, content_widget, dest_tabwidget, dest_index):
        """把一个标签（连同其嵌入内容/历史/状态）从源标签组转移到目标标签组。

        content_widget 为被拖拽标签对应的 FileExplorerTab（以对象身份定位，避免索引漂移）。
        保持“固定标签在前”不变量；左侧主组不允许被拖空；右侧组拖空后自动收起分屏。
        返回 True 表示已成功转移。"""
        if content_widget is None:
            return False
        src_tw, src_cs, src_is_right = self._resolve_group(source_tabwidget)
        dst_tw, dst_cs, dst_is_right = self._resolve_group(dest_tabwidget)
        if src_tw is dst_tw or src_cs is None or dst_cs is None:
            return False
        src_index = src_cs.indexOf(content_widget)
        if src_index < 0:
            return False
        # 左侧主组必须至少保留一个标签页
        if not src_is_right and src_tw.count() <= 1:
            show_toast(self, tr("拖拽标签"), tr("左侧至少需要保留一个标签页。"), level="warning", duration=2000)
            return False
        # 跨组转移是一次明确的最终操作，清除拖拽中标志，确保两组标签切换处理完整执行
        self._tab_drag_in_progress = False
        title = src_tw.tabText(src_index)
        is_pinned = getattr(content_widget, 'is_pinned', False)
        # 计算目标插入位置，并保持“固定标签在前”不变量
        dst_count = dst_tw.count()
        if dest_index is None or dest_index < 0 or dest_index > dst_count:
            dest_index = dst_count
        pinned_count = 0
        for i in range(dst_count):
            w = dst_cs.widget(i)
            if getattr(w, 'is_pinned', False):
                pinned_count += 1
        if is_pinned:
            dest_index = min(dest_index, pinned_count)
        else:
            dest_index = max(dest_index, pinned_count)
        # 从源组移除（先内容栈后标签栏，保持索引同步）
        src_cs.removeWidget(content_widget)
        src_tw.removeTab(src_index)
        # 插入到目标组（占位标签 + 实际内容，索引保持同步）
        dst_cs.insertWidget(dest_index, content_widget)
        dst_tw.insertTab(dest_index, QWidget(), title)
        dst_tw.setCurrentIndex(dest_index)
        try:
            content_widget.update_tab_title()
        except Exception:
            pass
        try:
            content_widget.set_refresh_active(True)
        except Exception:
            pass
        self._active_pane = content_widget
        # 源为右侧分屏组且已被拖空 → 收起分屏，回到单组
        if src_is_right and src_tw.count() == 0:
            self._teardown_split_group()
        try:
            self.update_navigation_buttons()
        except Exception:
            pass
        self.save_pinned_tabs()
        self._schedule_session_snapshot()
        return True

    def _on_group_tab_changed(self, target_tabwidget, index):
        """将标签切换分派到对应组的处理函数（左侧 on_tab_changed / 右侧 _on_split_tab_changed）。"""
        if target_tabwidget is not None and target_tabwidget is getattr(self, 'split_tab_widget', None):
            self._on_split_tab_changed(index)
        else:
            self.on_tab_changed(index)

    def _explorer_panes(self):
        """返回当前所有可交互的浏览面板：当前标签 + 分屏面板（若存在）。"""
        panes = []
        try:
            cur = self.get_current_tab_widget()
            if cur is not None and hasattr(cur, 'explorer'):
                panes.append(cur)
        except Exception:
            pass
        sp = self._get_split_pane()
        if sp is not None and hasattr(sp, 'explorer'):
            panes.append(sp)
        return panes

    def pane_at_global_pos(self, gx, gy):
        """返回屏幕坐标 (gx, gy) 命中的浏览面板（当前标签或分屏面板），未命中返回 None。"""
        from PyQt5.QtCore import QPoint
        pt = QPoint(int(gx), int(gy))
        for p in self._explorer_panes():
            try:
                ex = getattr(p, 'explorer', None)
                if ex is not None and ex.isVisible():
                    if ex.rect().contains(ex.mapFromGlobal(pt)):
                        return p
            except Exception:
                continue
        return None

    def set_active_pane_from_global_pos(self, gx, gy):
        """根据鼠标按下位置更新“活动面板”，供键盘快捷键/手势/双击定位目标面板。"""
        p = self.pane_at_global_pos(gx, gy)
        if p is not None and p is not getattr(self, '_active_pane', None):
            self._active_pane = p
            try:
                self.update_navigation_buttons()
            except Exception:
                pass

    def set_active_pane_to_group(self, target_tabwidget):
        """把“活动面板”显式设为指定标签组的当前面板。

        点击某一组的标签栏（即使未切换标签、不改变索引）也能可靠地把该组设为活动，
        使右上角按钮/终端/TortoiseGit 等作用于用户正在操作的一侧，而非总是左侧。"""
        try:
            if target_tabwidget is not None and target_tabwidget is getattr(self, 'split_tab_widget', None):
                p = self._get_split_pane()
            else:
                p = self.get_current_tab_widget()
            if p is not None and p is not getattr(self, '_active_pane', None):
                self._active_pane = p
                try:
                    self.update_navigation_buttons()
                except Exception:
                    pass
        except Exception:
            pass

    def get_active_pane(self):
        """返回当前操作应作用的浏览面板：最近交互的面板（含分屏），默认当前标签。

        分屏面板被关闭或引用失效后回退到当前标签，避免引用已销毁对象。
        以 `is` 身份比较，不解引用底层 C++ 对象，对已销毁 QWidget 安全。"""
        pane = getattr(self, '_active_pane', None)
        if pane is not None:
            if pane is self._get_split_pane():
                return pane
            try:
                if pane is self.get_current_tab_widget():
                    return pane
            except Exception:
                pass
            self._active_pane = None
        return self.get_current_tab_widget()

    def get_active_group_tabwidget(self):
        """返回当前活动面板所属的标签组 QTabWidget（右侧分屏当前面板 → split_tab_widget，否则左侧）。

        供手势/快捷键的“关闭当前标签、新建标签、恢复关闭标签”等操作定位到用户正在操作的一侧，
        避免在右侧分屏画手势时这些动作错误地作用于左侧组。"""
        try:
            pane = getattr(self, '_active_pane', None)
            if pane is not None and pane is self._get_split_pane():
                return self.split_tab_widget
        except Exception:
            pass
        return self.tab_widget

    def _set_restore_nav_guard(self):
        """窗口从最小化恢复时，设置 guard 抑制 IEB 树面板自动展开的虚假导航"""
        tab = self.get_current_tab_widget()
        if tab:
            tab._restore_guard_until = time.monotonic() + 2.0

    def _normalize_path_for_compare(self, path):
        """将路径标准化用于比较（Windows 不区分大小写）。"""
        if not path:
            return ""
        try:
            if path.startswith('shell:'):
                return path.lower()
            return os.path.normcase(os.path.normpath(path))
        except Exception:
            return str(path).lower()

    def find_tab_index_by_path(self, path):
        """查找已打开路径对应的标签索引，不存在返回 -1。"""
        target = self._normalize_path_for_compare(path)
        if not target:
            return -1
        for i in range(self.tab_widget.count()):
            tab = self.get_tab_widget(i)
            current_path = getattr(tab, 'current_path', '') if tab else ''
            if self._normalize_path_for_compare(current_path) == target:
                return i
        return -1

    def is_path_open(self, path):
        """路径是否已在任一标签中打开。"""
        return self.find_tab_index_by_path(path) >= 0
    

    def go_up_current_tab(self):
        current_tab = self.get_active_pane()
        if hasattr(current_tab, 'go_up'):
            current_tab.go_up(force=True)
    
    def go_back_current_tab(self):
        """后退当前标签页"""
        current_tab = self.get_active_pane()
        if current_tab and hasattr(current_tab, 'go_back'):
            current_tab.go_back()
    
    def go_forward_current_tab(self):
        """前进当前标签页"""
        current_tab = self.get_active_pane()
        if current_tab and hasattr(current_tab, 'go_forward'):
            current_tab.go_forward()
    
    def update_navigation_buttons(self):
        """更新前进后退按钮状态"""
        current_tab = self.get_active_pane()
        if current_tab and hasattr(current_tab, 'can_go_back'):
            self.back_button.setEnabled(current_tab.can_go_back())
        else:
            self.back_button.setEnabled(False)
        
        if current_tab and hasattr(current_tab, 'can_go_forward'):
            self.forward_button.setEnabled(current_tab.can_go_forward())
        else:
            self.forward_button.setEnabled(False)
    
    def open_tortoisegit_log_current_tab(self):
        """打开当前标签页的 TortoiseGit 日志（作用于活动面板，分屏时跟随最近交互的一侧）"""
        current_tab = self.get_active_pane()
        if current_tab and hasattr(current_tab, 'open_tortoisegit_log'):
            current_tab.open_tortoisegit_log()
    
    def open_tortoisegit_commit_current_tab(self):
        """打开当前标签页的 TortoiseGit 提交窗口（作用于活动面板，分屏时跟随最近交互的一侧）"""
        current_tab = self.get_active_pane()
        if current_tab and hasattr(current_tab, 'open_tortoisegit_commit'):
            current_tab.open_tortoisegit_commit()

    def _get_current_tab_launch_dir(self):
        """获取适合外部终端启动的当前标签页目录（作用于活动面板，分屏时跟随最近交互的一侧）。"""
        current_tab = self.get_active_pane()
        current_path = getattr(current_tab, 'current_path', '') if current_tab else ''

        if not current_path:
            show_toast(self, tr("提示"), tr("当前没有可用的标签页路径"), level="warning")
            return None
        if isinstance(current_path, str) and current_path.startswith('shell:'):
            show_toast(self, tr("提示"), tr("当前为特殊路径，无法定位到 cmd 或 PowerShell"), level="warning")
            return None
        if os.path.isdir(current_path):
            return normalize_external_launch_dir(current_path)
        if os.path.exists(current_path):
            return normalize_external_launch_dir(os.path.dirname(current_path))
        show_toast(self, tr("提示"), tr("当前标签页路径无效"), level="warning")
        return None

    def get_preferred_terminal_tool(self):
        return normalize_terminal_tool_name(self.config.get('preferred_terminal_tool', 'cmd'))

    def open_preferred_terminal_current_tab(self):
        current_dir = self._get_current_tab_launch_dir()
        if not current_dir:
            return
        try:
            launch_shell_tool(self.get_preferred_terminal_tool(), current_dir)
        except FileNotFoundError as e:
            show_toast(self, tr("提示"), str(e), level="warning")
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开默认终端: {}").format(e), level="error")

    def open_cmd_current_tab(self):
        """在当前标签页目录打开 cmd。"""
        current_dir = self._get_current_tab_launch_dir()
        if not current_dir:
            return
        try:
            launch_shell_tool('cmd', current_dir)
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开 cmd: {}").format(e), level="error")

    def open_powershell_current_tab(self):
        """在当前标签页目录打开 PowerShell。"""
        current_dir = self._get_current_tab_launch_dir()
        if not current_dir:
            return
        try:
            launch_shell_tool('powershell', current_dir)
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开 PowerShell: {}").format(e), level="error")

    def open_git_bash_current_tab(self):
        """在当前标签页目录打开 Git Bash。"""
        current_dir = self._get_current_tab_launch_dir()
        if not current_dir:
            return
        try:
            launch_shell_tool('git-bash', current_dir)
        except FileNotFoundError as e:
            show_toast(self, tr("提示"), str(e), level="warning")
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开 Git Bash: {}").format(e), level="error")

    def open_calculator(self):
        """打开系统计算器。"""
        try:
            launch_shell_tool('calculator')
        except OSError as e:
            show_toast(self, tr("提示"), str(e), level="warning")
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开计算器: {}").format(e), level="error")
    
    def apply_tortoisegit_buttons_config(self):
        """根据配置显示/隐藏 TortoiseGit 按钮"""
        enable = self.config.get("enable_tortoisegit_buttons", False)
        if hasattr(self, 'git_log_button'):
            self.git_log_button.setVisible(enable)
        if hasattr(self, 'git_commit_button'):
            self.git_commit_button.setVisible(enable)
        if hasattr(self, 'git_bash_button'):
            self.git_bash_button.setVisible(enable)
        if hasattr(self, 'git_tools_separator'):
            self.git_tools_separator.setVisible(enable)

    def apply_title_shortcuts_config(self):
        """根据配置显示/隐藏标题栏快捷方式区域。"""
        enable = self.config.get("enable_title_shortcuts", True)
        if hasattr(self, 'title_shortcut_bar'):
            self.title_shortcut_bar.setVisible(enable)
        if hasattr(self, 'shortcut_git_separator'):
            self.shortcut_git_separator.setVisible(enable)

    def refresh_title_shortcuts_ui(self):
        """刷新标题栏快捷方式按钮。"""
        if not hasattr(self, 'title_shortcut_bar'):
            return
        shortcuts = self.config.get("title_shortcuts", [])
        if not isinstance(shortcuts, list):
            shortcuts = []
        cleaned = []
        for path in shortcuts:
            if isinstance(path, str) and path and os.path.exists(path) and path not in cleaned:
                cleaned.append(path)
        if cleaned != shortcuts:
            self.config["title_shortcuts"] = cleaned
            self.save_config()
        self.title_shortcut_bar.set_shortcuts(cleaned)

    def on_title_shortcut_dropped(self, path):
        """处理拖入标题栏快捷方式区域的启动文件。"""
        if not is_supported_title_shortcut_path(path):
            return
        shortcuts = self.config.get("title_shortcuts", [])
        if not isinstance(shortcuts, list):
            shortcuts = []
        if path in shortcuts:
            return
        # 新拖入的快捷方式优先放在左侧
        shortcuts.insert(0, path)
        self.config["title_shortcuts"] = shortcuts[:20]
        self.refresh_title_shortcuts_ui()
        self.save_config()

    def on_title_shortcuts_changed(self, paths):
        """处理标题栏快捷方式顺序/删除变更。"""
        self.config["title_shortcuts"] = list(paths or [])[:20]
        self.save_config()

    def open_title_shortcut(self, path):
        """点击标题栏快捷方式后启动对应程序。"""
        if not path or not os.path.exists(path):
            show_toast(self, tr("提示"), tr("快捷方式不存在，已从列表移除"), level="warning")
            shortcuts = self.config.get("title_shortcuts", [])
            if path in shortcuts:
                shortcuts = [p for p in shortcuts if p != path]
                self.config["title_shortcuts"] = shortcuts
                self.refresh_title_shortcuts_ui()
                self.save_config()
            return
        try:
            display_name = os.path.splitext(os.path.basename(path))[0] or os.path.basename(path)
            lower_path = path.lower()
            if os.name == 'nt' and lower_path.endswith('.ps1'):
                launch_detached([
                    'powershell.exe',
                    '-ExecutionPolicy',
                    'Bypass',
                    '-File',
                    path,
                ], cwd=os.path.dirname(path) or None)
            elif os.name == 'nt' and lower_path.endswith(('.bat', '.cmd')):
                launch_detached(['cmd.exe', '/c', 'start', '', path], cwd=os.path.dirname(path) or None)
            elif os.name == 'nt':
                os.startfile(path)
            else:
                launch_detached([path])
            show_toast(self, tr("已启动"), tr("运行: {}").format(display_name), level="info")
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法启动快捷方式: {}").format(e), level="error")

    def _ensure_chat_panel_created(self):
        """确保聊天面板已创建（延迟创建）"""
        if self.chat_panel is None:
            # 首次创建聊天面板
            self.chat_panel = ChatPanel(self)
            ai_panel_width = int(self.config.get("ai_chat", {}).get("panel_width", 360) * self.dpi_scale)
            self.chat_panel.setMinimumWidth(260)
            self.chat_panel.setFixedWidth(ai_panel_width)
            self.chat_panel.setVisible(False)
            # 根据当前是否分屏，决定插入位置
            if getattr(self, '_split_active', False):
                # 如果已经分屏，chat_panel 应该在索引 2
                self.splitter.insertWidget(2, self.chat_panel)
                self.splitter.setCollapsible(2, True)
            else:
                # 没有分屏，chat_panel 在索引 1
                self.splitter.addWidget(self.chat_panel)
                self.splitter.setCollapsible(1, True)

    def toggle_chat_panel(self):
        """切换 AI 聊天面板的显示/隐藏。"""
        self._ensure_chat_panel_created()  # 确保面板已创建
        visible = not self.chat_panel.isVisible()
        self.chat_panel.setVisible(visible)
        if hasattr(self, 'ai_chat_btn'):
            self.ai_chat_btn.setChecked(visible)
        # 保存AI面板的显示状态到配置
        if "ai_chat" not in self.config:
            self.config["ai_chat"] = {}
        self.config["ai_chat"]["panel_visible"] = visible
        self.save_config()
        if visible:
            # 保存面板宽度到 splitter
            ai_panel_width = int(
                self.config.get("ai_chat", {}).get("panel_width", 360) * self.dpi_scale
            )
            sizes = self.splitter.sizes()
            total = sum(sizes)
            self.splitter.setSizes([total - ai_panel_width, ai_panel_width])
            # 更新当前目录提示
            self.update_chat_context()
            self.chat_panel.input_box.setFocus()

    def toggle_split_view(self):
        """切换左右分屏：把当前标签移动到右侧标签组；再次触发则把右侧标签移回左侧并关闭分屏。

        与“复制视图”不同，这里移动的是真实标签内容（含其嵌入的资源管理器、历史与状态），
        右侧拥有自己的标签栏（位于书签栏上方，与左侧并排），并支持双击空白处新建标签、
        关闭、拖拽等原生功能。左右两侧均为完整可交互标签组，快捷键/手势作用于最近点击的面板。"""
        if getattr(self, '_split_active', False):
            self._merge_split_back()
            show_toast(self, tr("分屏对比"), tr("已将右侧标签合并回左侧。"), level="info", duration=1500)
        else:
            self._enter_split_view()

    def _on_split_tab_changed(self, index):
        """右侧标签组当前项变化：同步右侧内容栈，仅当前标签保持刷新，其余暂停以避免卡顿。"""
        scs = getattr(self, 'split_content_stack', None)
        if scs is None:
            return
        if 0 <= index < scs.count():
            scs.setCurrentIndex(index)
        # 拖拽重排期间仅同步显示，跳过重型刷新切换（与左侧 on_tab_changed 一致）
        if getattr(self, '_tab_drag_in_progress', False):
            return
        # 仅右侧当前可见标签保持高频刷新；其余右侧后台标签暂停轮询，
        # 避免多次切换后累积多个活跃面板导致 COM/scandir 高频轮询卡顿
        for i in range(scs.count()):
            tab_item = scs.widget(i)
            if tab_item and hasattr(tab_item, 'set_refresh_active'):
                try:
                    tab_item.set_refresh_active(i == index)
                except Exception:
                    pass
        content = scs.widget(index) if 0 <= index < scs.count() else None
        if content is not None:
            self._active_pane = content
        try:
            self.update_navigation_buttons()
        except Exception:
            pass

    def _get_split_pane(self):
        """返回右侧分屏标签组当前显示的面板（FileExplorerTab），未分屏返回 None。"""
        if not getattr(self, '_split_active', False):
            return None
        stw = getattr(self, 'split_tab_widget', None)
        scs = getattr(self, 'split_content_stack', None)
        if stw is None or scs is None or stw.count() == 0:
            return None
        idx = stw.currentIndex()
        if 0 <= idx < scs.count():
            return scs.widget(idx)
        return None

    def _sync_split_tabbar_width(self):
        """让右侧标签栏宽度跟随右侧内容面板宽度，使其与右侧内容上下对齐。"""
        if not getattr(self, '_split_active', False):
            return
        stw = getattr(self, 'split_tab_widget', None)
        scs = getattr(self, 'split_content_stack', None)
        if stw is None or scs is None:
            return
        try:
            w = scs.width()
            # 仅在宽度变化时才 setFixedWidth，避免拖动分隔条时高频触发布局重算导致卡顿
            if w > 0 and stw.maximumWidth() != w:
                stw.setFixedWidth(w)
        except Exception:
            pass

    def _activate_split_layout(self):
        """显示右侧分屏组的 UI 布局（加入分割器、显示、设置尺寸与折叠属性）。

        供“进入分屏”与“启动/崩溃恢复”共用，只负责布局，不涉及具体标签内容。"""
        if self.splitter.indexOf(self.split_content_stack) < 0:
            self.splitter.insertWidget(1, self.split_content_stack)
        self.split_content_stack.setVisible(True)
        self.split_tab_widget.setVisible(True)
        self._split_active = True
        # 分屏后：左侧内容(0)、右侧内容(1) 不可折叠，AI 面板(2) 可折叠（如果存在）
        try:
            self.splitter.setCollapsible(0, False)
            self.splitter.setCollapsible(1, False)
            if self.chat_panel is not None and self.splitter.count() > 2:
                self.splitter.setCollapsible(2, True)
        except Exception:
            pass
        # 左右各半，保留 AI 面板宽度
        try:
            chat_w = self.chat_panel.width() if (self.chat_panel is not None and self.chat_panel.isVisible()) else 0
            total = max(self.splitter.width() - chat_w, 200)
            half = max(total // 2, 100)
            if self.chat_panel is not None:
                self.splitter.setSizes([half, half, chat_w])
            else:
                self.splitter.setSizes([half, half])
        except Exception:
            pass
        from PyQt5.QtCore import QTimer as _QTimer
        _QTimer.singleShot(0, self._sync_split_tabbar_width)
        if hasattr(self, 'split_view_btn'):
            self.split_view_btn.setChecked(True)

    def _enter_split_view(self):
        """把当前标签从左侧标签组移动到右侧标签组（含其嵌入内容、历史与状态）。"""
        cur = self.tab_widget.currentIndex()
        if cur < 0:
            return
        content = self.content_stack.widget(cur)
        if content is None:
            return
        title = self.tab_widget.tabText(cur)
        # 仅有一个标签时，先在左侧补一个默认标签，避免移走后左侧空白
        if self.tab_widget.count() < 2:
            self.add_new_tab()
            if self.tab_widget.count() < 2:
                show_toast(self, tr("分屏失败"), tr("无法创建用于左侧的新标签页。"), level="warning", duration=2500)
                return
        # 重新计算被移动标签的当前索引（content_stack 与 tab_widget 索引保持一致）
        idx = self.content_stack.indexOf(content)
        if idx < 0:
            return
        # 从左侧分离（先 content_stack 再 tab_widget，保持索引同步）
        self.content_stack.removeWidget(content)
        self.tab_widget.removeTab(idx)
        # 显示右侧分屏 UI 布局（加入分割器并设置尺寸）
        self._activate_split_layout()
        # 放入右侧标签组（占位标签 + 实际内容，索引保持同步）
        self.split_content_stack.addWidget(content)
        new_idx = self.split_tab_widget.addTab(QWidget(), title)
        self.split_tab_widget.setCurrentIndex(new_idx)
        try:
            content.set_refresh_active(True)
        except Exception:
            pass
        self._active_pane = content
        # 持久化分屏状态，确保重启/崩溃后可恢复
        self._schedule_session_snapshot()
        show_toast(self, tr("分屏对比"), tr("已将当前标签移到右侧。再次按 F3 合并回左侧。"), level="info", duration=2000)

    def _merge_split_back(self):
        """把右侧标签组中的标签全部移回左侧标签组并关闭分屏（静默，供关闭流程复用）。"""
        stw = getattr(self, 'split_tab_widget', None)
        scs = getattr(self, 'split_content_stack', None)
        if stw is None or scs is None:
            self._teardown_split_group()
            return
        last_idx = -1
        while stw.count() > 0:
            content = scs.widget(0)
            title = stw.tabText(0)
            stw.removeTab(0)
            if content is None:
                continue
            scs.removeWidget(content)
            self.content_stack.addWidget(content)
            last_idx = self.tab_widget.addTab(QWidget(), title)
        if last_idx >= 0:
            self.tab_widget.setCurrentIndex(last_idx)
        self._teardown_split_group()
        # 合并回左侧后持久化，清除已保存的分屏状态
        self._schedule_session_snapshot()

    def _teardown_split_group(self):
        """收起右侧标签组：隐藏其标签栏与内容栈，从分割器移除并恢复单组布局。"""
        self._split_active = False
        stw = getattr(self, 'split_tab_widget', None)
        scs = getattr(self, 'split_content_stack', None)
        if stw is not None:
            stw.setVisible(False)
            stw.setMinimumWidth(0)
            stw.setMaximumWidth(16777215)
        if scs is not None:
            scs.setVisible(False)
            try:
                scs.setParent(None)  # 从分割器移除，恢复 [content_stack, chat_panel] 两元素布局
            except Exception:
                pass
        self._active_pane = None
        # 恢复 splitter 折叠属性：content_stack(0) 不可折叠，AI 面板(1) 可折叠
        try:
            self.splitter.setCollapsible(0, False)
            self.splitter.setCollapsible(1, True)
        except Exception:
            pass
        try:
            chat_w = self.chat_panel.width() if (self.chat_panel is not None and self.chat_panel.isVisible()) else 0
            total = max(self.splitter.width() - chat_w, 200)
            self.splitter.setSizes([total, chat_w])
        except Exception:
            pass
        if hasattr(self, 'split_view_btn'):
            self.split_view_btn.setChecked(False)

    def update_chat_context(self):
        """将当前标签页路径同步到 AI 聊天面板的上下文提示。"""
        if self.chat_panel is None or not self.chat_panel.isVisible():
            return
        try:
            tab = self.get_current_tab_widget()
            if tab and hasattr(tab, 'current_path'):
                self.chat_panel.update_context(tab.current_path or "")
        except Exception:
            pass

    def _update_window_title(self, current_path: str = None):
        """根据当前状态更新窗口标题和自定义标题栏文本。
        当处于恢复状态时，在标题中附加“窗口恢复中”。
        """
        base_title = f"TabExplorer v{APP_VERSION}"
        # 更新自定义标题栏文本
        if hasattr(self, 'title_label'):
            if self.is_restoring:
                self.title_label.setText(tr("{} - 窗口恢复中").format(base_title))
            else:
                self.title_label.setText(base_title)

        # 组合窗口标题
        if current_path:
            title = f"{base_title} - {current_path}"
        else:
            title = base_title
        if self.is_restoring:
            title = tr("{} - 窗口恢复中").format(title)
        try:
            self.setWindowTitle(title)
        except Exception:
            pass


    @pyqtSlot()
    @pyqtSlot(str)
    @pyqtSlot(str, bool)
    def add_new_tab(self, path="", is_shell=False, select_file=None, target_tabwidget=None, activate=True):
        # 默认新建标签页为“此电脑”
        if not path:
            path = 'shell:MyComputerFolder'
            is_shell = True
        
        tab_widget, content_stack, _is_right = self._resolve_group(target_tabwidget)

        try:
            # activate=False（会话恢复用）：延迟首次导航到该标签首次可见时，避免启动瞬间
            # 大量 IExplorerBrowser 同时创建/导航导致的 CPU 洪峰。
            tab = FileExplorerTab(self, path, is_shell=is_shell, select_file=select_file,
                                  defer_nav=(not activate))
        except Exception as e:
            debug_print(f"[MainWindow] Failed to create embedded explorer tab for '{path}': {e}")
            try:
                if path:
                    launch_detached(['explorer.exe', path])
            except Exception as open_error:
                debug_print(f"[MainWindow] Fallback explorer launch failed for '{path}': {open_error}")
            show_toast(self, tr("打开失败"), tr("无法嵌入该窗口，已尝试用系统资源管理器打开。\n{}").format(e), level="error", duration=3500)
            return -1
        tab.is_pinned = False
        short = path[-16:] if len(path) > 16 else path
        
        # 同时添加到 tab_widget（占位标签）和 content_stack（实际内容）
        tab_index = tab_widget.addTab(QWidget(), short)
        content_stack.addWidget(tab)
        
        # activate=True：切到该标签（触发 showEvent → 首次导航）。
        # activate=False：不激活，标签保持隐藏，其导航延迟到用户首次切过去（懒加载）。
        if activate:
            tab_widget.setCurrentIndex(tab_index)
        
        # 更新导航按钮状态（确保新标签页的按钮状态正确）
        self.update_navigation_buttons()

        self._schedule_session_snapshot()
        
        return tab_index


    def close_tab(self, index, target_tabwidget=None):
        tab_widget, content_stack, is_right = self._resolve_group(target_tabwidget)
        tab = content_stack.widget(index) if (content_stack and 0 <= index < content_stack.count()) else None
        if tab and hasattr(tab, 'cleanup'):
            try:
                tab.cleanup()
            except Exception as e:
                debug_print(f"[ClosedTabs] Tab cleanup failed: {e}")
        if tab and hasattr(tab, 'set_refresh_active'):
            try:
                tab.set_refresh_active(False)
            except Exception:
                pass
        
        # 调试：打印标签页信息
        debug_print(f"[ClosedTabs] Closing tab at index {index}")
        debug_print(f"[ClosedTabs] Tab type: {type(tab)}")
        debug_print(f"[ClosedTabs] Has current_path: {hasattr(tab, 'current_path')}")
        if hasattr(tab, 'current_path'):
            debug_print(f"[ClosedTabs] current_path value: {tab.current_path}")
        
        # 保存到关闭历史（在移除之前）
        if hasattr(tab, 'current_path') and tab.current_path:
            tab_info = {
                'path': tab.current_path,
                'title': tab_widget.tabText(index),
                'is_shell': tab.current_path.startswith('shell:') if hasattr(tab, 'current_path') else False
            }
            # 添加到历史列表开头
            self.closed_tabs_history.insert(0, tab_info)
            # 限制历史数量
            if len(self.closed_tabs_history) > self.max_closed_tabs_history:
                self.closed_tabs_history = self.closed_tabs_history[:self.max_closed_tabs_history]
            debug_print(f"[ClosedTabs] Saved to history: {tab_info['path']}, total history: {len(self.closed_tabs_history)}")
            
            # 更新恢复按钮状态
            if hasattr(self, 'reopen_tab_button'):
                self.reopen_tab_button.setEnabled(True)
        else:
            debug_print(f"[ClosedTabs] Not saved - no valid current_path")
        
        # 如果是固定标签页，关闭时自动移除固定
        if hasattr(tab, 'is_pinned') and tab.is_pinned:
            tab.is_pinned = False
            self.save_pinned_tabs()
        if tab_widget.count() > 1:
            # 先从 content_stack 移除（这样 on_tab_changed 触发时两者已同步）
            if content_stack is not None and index < content_stack.count():
                widget = content_stack.widget(index)
                content_stack.removeWidget(widget)
                if widget:
                    widget.deleteLater()
            tab_widget.removeTab(index)
            self._schedule_session_snapshot()
        else:
            # 该组仅剩一个标签
            if is_right:
                # 关闭右侧分屏组的最后一个标签 → 折叠分屏，回到单组
                if content_stack is not None and index < content_stack.count():
                    widget = content_stack.widget(index)
                    content_stack.removeWidget(widget)
                    if widget:
                        widget.deleteLater()
                tab_widget.removeTab(index)
                self._teardown_split_group()
                self._schedule_session_snapshot()
            else:
                self.close()


    def close_current_tab(self):
        # 作用于活动面板所属组：在右侧分屏操作时关闭右侧当前标签，而非总是左侧
        tw = self.get_active_group_tabwidget()
        self.close_tab(tw.currentIndex(), target_tabwidget=tw)
    
    def reopen_closed_tab(self):
        """恢复最近关闭的标签页"""
        if not self.closed_tabs_history:
            debug_print("[ClosedTabs] No closed tabs to restore")
            return
        
        # 取出最近关闭的标签页
        tab_info = self.closed_tabs_history.pop(0)
        debug_print(f"[ClosedTabs] Restoring tab: {tab_info['path']}, remaining history: {len(self.closed_tabs_history)}")
        
        # 重新打开标签页（作用于活动面板所属组，右侧分屏操作时恢复到右侧）
        self.add_new_tab(tab_info['path'], is_shell=tab_info.get('is_shell', False),
                         target_tabwidget=self.get_active_group_tabwidget())
        
        # 更新恢复按钮状态
        if hasattr(self, 'reopen_tab_button'):
            self.reopen_tab_button.setEnabled(len(self.closed_tabs_history) > 0)

    def on_tab_changed(self, index):
        if index >= 0:
            # 调试信息：检查同步状态
            debug_print(f"[TabSwitch] Tab changed to index {index}")
            if hasattr(self, 'content_stack'):
                debug_print(f"[TabSwitch] content_stack has {self.content_stack.count()} widgets, tab_widget has {self.tab_widget.count()} tabs")

            # 同步 content_stack 的显示（拖拽期间也需要保持同步）
            if hasattr(self, 'content_stack') and index < self.content_stack.count():
                self.content_stack.setCurrentIndex(index)
                debug_print(f"[TabSwitch] Set content_stack to index {index}")
            else:
                debug_print(f"[TabSwitch] WARNING: Cannot sync - content_stack count is {self.content_stack.count() if hasattr(self, 'content_stack') else 'N/A'}")

            # 拖拽进行中：仅同步 content_stack，跳过一切重型操作，等拖拽结束后统一执行
            if getattr(self, '_tab_drag_in_progress', False):
                return

            # 切换标签页后重置“活动面板”为当前标签，避免快捷键仍指向旧分屏/旧标签
            self._active_pane = None

            # 仅当前可见标签保持高频刷新；后台标签暂停轮询并仅记录脏状态
            for i in range(self.tab_widget.count()):
                tab_item = self.get_tab_widget(i)
                if tab_item and hasattr(tab_item, 'set_refresh_active'):
                    try:
                        tab_item.set_refresh_active(i == index)
                    except Exception:
                        pass
            
            # 切换标签后立即把资源占用刷到新活动标签的标签上（若功能开启）
            self._update_resource_usage_display()

            # 从 content_stack 获取实际的标签页内容
            tab = self.content_stack.widget(index) if hasattr(self, 'content_stack') else self.tab_widget.widget(index)
            if hasattr(tab, 'current_path'):
                # 切换标签时强制刷新该标签路径栏，避免显示上一个标签路径
                try:
                    if hasattr(tab, 'path_bar') and tab.path_bar:
                        try:
                            tab.path_bar.exit_edit_mode()
                        except Exception:
                            pass
                        tab.path_bar.set_path(tab.current_path)
                except Exception:
                    pass
                # 统一通过内部方法更新窗口标题（可带“窗口恢复中”标记）
                self._update_window_title(tab.current_path)
            # 更新导航按钮状态
            self.update_navigation_buttons()
            # 同步 AI 聊天面板的当前目录提示
            self.update_chat_context()
            self._schedule_session_snapshot()
        
        # 选中/非选中标签样式统一由创建时的共享样式表控制（左右两组一致，淡黄色选中）。
        # 此处不再每次切换重设左侧样式，避免与右侧分屏组样式分叉，并省去每次切换的重绘开销。


    def dragEnterEvent(self, event):
        """主窗口拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            debug_print("[DEBUG] MainWindow: Drag enter accepted")
        else:
            event.ignore()
    
    def dragMoveEvent(self, event):
        """主窗口拖拽移动事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def dropEvent(self, event):
        """主窗口拖拽释放事件"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            debug_print(f"[DEBUG] MainWindow: Drop event, urls count: {len(urls)}")
            
            for url in urls:
                path = None
                # 尝试获取本地文件路径
                if url.isLocalFile():
                    path = url.toLocalFile()
                else:
                    # 尝试从 URL 字符串中提取路径（支持网络路径）
                    url_str = url.toString()
                    if url_str.startswith('file:///'):
                        from urllib.parse import unquote
                        path = unquote(url_str[8:])
                        if os.name == 'nt' and path.startswith('/'):
                            path = path[1:]
                    elif url_str.startswith('file://'):
                        from urllib.parse import unquote
                        # 网络路径 file://server/share
                        path = '\\\\' + unquote(url_str[7:]).replace('/', '\\')
                
                if path and os.path.exists(path):
                    debug_print(f"[DEBUG] MainWindow: Processing dropped path: {path}")
                    if os.path.isdir(path):
                        # 如果是文件夹，打开新标签页
                        self.add_new_tab(path)
                    elif os.path.isfile(path):
                        # 如果是文件，打开其所在文件夹
                        folder = os.path.dirname(path)
                        self.add_new_tab(folder)
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def create_custom_titlebar(self, main_layout):
        """创建工具栏（系统原生标题栏下方的功能按钮区域）"""
        # 根据DPI调整工具栏高度
        titlebar_height = int(32 * getattr(self, 'dpi_scale', 1.0))
        titlebar = QWidget()
        titlebar.setFixedHeight(titlebar_height)
        titlebar.setStyleSheet("background-color: #f3f3f3;")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(10, 0, 0, 0)
        titlebar_layout.setSpacing(0)
        
        # 保存引用（兼容其他代码对 titlebar_widget 的引用）
        self.titlebar_widget = titlebar
        
        # TortoiseGit 按钮（可在设置中启用/禁用）
        btn_size = int(32 * getattr(self, 'dpi_scale', 1.0))
        btn_font_size = int(14 * getattr(self, 'dpi_scale', 1.0))
        btn_radius = int(4 * getattr(self, 'dpi_scale', 1.0))

        # 标题栏快捷方式区域（位于 Git 按钮左侧）
        shortcut_btn_size = max(int(24 * getattr(self, 'dpi_scale', 1.0)), btn_size - int(6 * getattr(self, 'dpi_scale', 1.0)))
        icon_size = max(14, int(shortcut_btn_size * 0.62))
        self.title_shortcut_bar = TitleShortcutBar(self, icon_size=icon_size, button_size=shortcut_btn_size)
        self.title_shortcut_bar.shortcutDropped.connect(self.on_title_shortcut_dropped)
        self.title_shortcut_bar.shortcutClicked.connect(self.open_title_shortcut)
        self.title_shortcut_bar.shortcutsChanged.connect(self.on_title_shortcuts_changed)
        self.refresh_title_shortcuts_ui()
        titlebar_layout.addWidget(self.title_shortcut_bar)

        git_btn_style = f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                font-weight: bold;
                color: #333;
            }}
            QPushButton:hover {{
                background: #e0e0e0;
            }}
            QPushButton:pressed {{
                background: #d0d0d0;
            }}
        """

        # 快捷方式区域与 Git 区域分隔线
        self.shortcut_git_separator = QFrame()
        self.shortcut_git_separator.setFrameShape(QFrame.VLine)
        self.shortcut_git_separator.setFrameShadow(QFrame.Plain)
        self.shortcut_git_separator.setStyleSheet("background-color: #d0d0d0; max-width: 1px;")
        self.shortcut_git_separator.setFixedHeight(int(20 * getattr(self, 'dpi_scale', 1.0)))
        titlebar_layout.addWidget(self.shortcut_git_separator)
        
        # Git Log 按钮
        self.git_log_button = QPushButton("🐢")
        self.git_log_button.setToolTip(tr("打开 TortoiseGit 日志"))
        self.git_log_button.setFixedSize(btn_size, btn_size)
        self.git_log_button.setStyleSheet(git_btn_style)
        self.git_log_button.clicked.connect(self.open_tortoisegit_log_current_tab)
        titlebar_layout.addWidget(self.git_log_button)
        
        # Git Commit 按钮
        self.git_commit_button = QPushButton("📤")
        self.git_commit_button.setToolTip(tr("打开 TortoiseGit 提交窗口"))
        self.git_commit_button.setFixedSize(btn_size, btn_size)
        self.git_commit_button.setStyleSheet(git_btn_style)
        self.git_commit_button.clicked.connect(self.open_tortoisegit_commit_current_tab)
        titlebar_layout.addWidget(self.git_commit_button)

        # 工具按钮文字样式
        tool_btn_font_size = max(int(8 * getattr(self, 'dpi_scale', 1.0)), 8)
        def _tool_btn_style(color):
            return f"""
                QPushButton {{
                    background: transparent;
                    border: none;
                    border-radius: {btn_radius}px;
                    font-size: {tool_btn_font_size}pt;
                    font-weight: bold;
                    color: {color};
                    padding: 0px;
                }}
                QPushButton:hover {{
                    background: #e0e0e0;
                }}
                QPushButton:pressed {{
                    background: #d0d0d0;
                }}
            """

        # Git Bash 按钮
        self.git_bash_button = QPushButton("GB")
        self.git_bash_button.setToolTip(tr("在当前标签页路径打开 Git Bash"))
        self.git_bash_button.setFixedSize(btn_size, btn_size)
        self.git_bash_button.setStyleSheet(_tool_btn_style("#e44d26"))
        self.git_bash_button.clicked.connect(self.open_git_bash_current_tab)
        titlebar_layout.addWidget(self.git_bash_button)

        # Git 与终端/工具按钮组之间的分隔线
        self.git_tools_separator = QFrame()
        self.git_tools_separator.setFrameShape(QFrame.VLine)
        self.git_tools_separator.setFrameShadow(QFrame.Plain)
        self.git_tools_separator.setStyleSheet("background-color: #d0d0d0; max-width: 1px;")
        self.git_tools_separator.setFixedHeight(int(20 * getattr(self, 'dpi_scale', 1.0)))
        titlebar_layout.addWidget(self.git_tools_separator)

        # 终端/工具按钮组（位于 Git 按钮右侧）
        self.cmd_button = QPushButton("CMD")
        self.cmd_button.setToolTip(tr("在当前标签页路径打开 cmd"))
        self.cmd_button.setFixedSize(btn_size, btn_size)
        self.cmd_button.setStyleSheet(_tool_btn_style("#1a1a1a"))
        self.cmd_button.clicked.connect(self.open_cmd_current_tab)
        titlebar_layout.addWidget(self.cmd_button)

        self.powershell_button = QPushButton("PS")
        self.powershell_button.setToolTip(tr("在当前标签页路径打开 PowerShell"))
        self.powershell_button.setFixedSize(btn_size, btn_size)
        self.powershell_button.setStyleSheet(_tool_btn_style("#2979ff"))
        self.powershell_button.clicked.connect(self.open_powershell_current_tab)
        titlebar_layout.addWidget(self.powershell_button)

        self.calculator_button = QPushButton("CAL")
        self.calculator_button.setToolTip(tr("打开计算器"))
        self.calculator_button.setFixedSize(btn_size, btn_size)
        self.calculator_button.setStyleSheet(_tool_btn_style("#c43e1c"))
        self.calculator_button.clicked.connect(self.open_calculator)
        titlebar_layout.addWidget(self.calculator_button)
        
        # 分隔线（可选）
        separator = QFrame()
        separator.setFrameShape(QFrame.VLine)
        separator.setFrameShadow(QFrame.Plain)
        separator.setStyleSheet("background-color: #d0d0d0; max-width: 1px;")
        separator.setFixedHeight(int(20 * getattr(self, 'dpi_scale', 1.0)))
        titlebar_layout.addWidget(separator)
        
        # 标签栏导航按钮（从标签栏移到这里）
        # 后退按钮
        self.back_button = QPushButton("◀")
        self.back_button.setToolTip(tr("后退 (Alt+←)"))
        self.back_button.setFixedSize(btn_size, btn_size)
        self.back_button.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                font-weight: bold;
                color: #202020;
            }}
            QPushButton:hover:!disabled {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:pressed:!disabled {{
                background: #d5d5d5;
                color: #000000;
            }}
            QPushButton:disabled {{
                color: #c0c0c0;
            }}
        """)
        self.back_button.clicked.connect(self.go_back_current_tab)
        self.back_button.setEnabled(False)
        titlebar_layout.addWidget(self.back_button)
        
        # 前进按钮
        self.forward_button = QPushButton("▶")
        self.forward_button.setToolTip(tr("前进 (Alt+→)"))
        self.forward_button.setFixedSize(btn_size, btn_size)
        self.forward_button.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                font-weight: bold;
                color: #202020;
            }}
            QPushButton:hover:!disabled {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:pressed:!disabled {{
                background: #d5d5d5;
                color: #000000;
            }}
            QPushButton:disabled {{
                color: #c0c0c0;
            }}
        """)
        self.forward_button.clicked.connect(self.go_forward_current_tab)
        self.forward_button.setEnabled(False)
        titlebar_layout.addWidget(self.forward_button)
        
        # 新建标签页按钮
        self.add_tab_button = QPushButton("+")
        self.add_tab_button.setToolTip(tr("新建标签页 (Ctrl+T)"))
        self.add_tab_button.setFixedSize(btn_size, btn_size)
        self.add_tab_button.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                color: #202020;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:pressed {{
                background: #d5d5d5;
                color: #000000;
            }}
        """)
        self.add_tab_button.clicked.connect(self.add_new_tab)
        titlebar_layout.addWidget(self.add_tab_button)
        
        # 恢复标签页按钮
        self.reopen_tab_button = QPushButton("↺")
        self.reopen_tab_button.setToolTip(tr("恢复关闭的标签页 (Ctrl+Shift+T)"))
        self.reopen_tab_button.setFixedSize(btn_size, btn_size)
        self.reopen_tab_button.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                font-weight: bold;
                color: #202020;
            }}
            QPushButton:hover:!disabled {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:pressed:!disabled {{
                background: #d5d5d5;
                color: #000000;
            }}
            QPushButton:disabled {{
                color: #c0c0c0;
            }}
        """)
        self.reopen_tab_button.clicked.connect(self.reopen_closed_tab)
        self.reopen_tab_button.setEnabled(False)
        titlebar_layout.addWidget(self.reopen_tab_button)
        
        # 搜索按钮
        self.search_button = QPushButton("⌕")
        self.search_button.setToolTip(tr("搜索当前文件夹 (Ctrl+F)"))
        self.search_button.setFixedSize(btn_size, btn_size)
        search_icon_size = int(13 * getattr(self, 'dpi_scale', 1.0))
        self.search_button.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {search_icon_size}pt;
                padding: 2px;
                color: #202020;
            }}
            QPushButton:hover {{
                background: #e5e5e5;
                border: 1px solid #d0d0d0;
                color: #000000;
            }}
            QPushButton:pressed {{
                background: #d5d5d5;
                border: 1px solid #c0c0c0;
                color: #000000;
            }}
        """)
        self.search_button.clicked.connect(self.show_search_dialog)
        titlebar_layout.addWidget(self.search_button)
        
        # 分屏对比按钮（切换右侧第二个独立浏览面板）
        self.split_view_btn = QPushButton("◫")
        self.split_view_btn.setToolTip(tr("分屏对比 (F3)"))
        self.split_view_btn.setFixedSize(btn_size, btn_size)
        self.split_view_btn.setCheckable(True)
        self.split_view_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                color: #202020;
            }}
            QPushButton:hover {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:pressed {{
                background: #d5d5d5;
                color: #000000;
            }}
            QPushButton:checked {{
                background: #BBDEFB;
                color: #1565C0;
            }}
        """)
        self.split_view_btn.clicked.connect(self.toggle_split_view)
        titlebar_layout.addWidget(self.split_view_btn)
        
        # 书签管理按钮
        bookmark_btn = QPushButton("★")
        bookmark_btn.setToolTip(tr("书签管理"))
        bookmark_btn_width = int(40 * getattr(self, 'dpi_scale', 1.0))
        bookmark_btn.setFixedSize(bookmark_btn_width, titlebar_height)
        bookmark_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                color: #202020;
            }}
            QPushButton:hover {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:pressed {{
                background: #d5d5d5;
                color: #000000;
            }}
        """)
        bookmark_btn.clicked.connect(self.show_bookmark_manager_dialog)
        titlebar_layout.addWidget(bookmark_btn)
        
        # 设置按钮
        settings_btn = QPushButton("⚙")
        settings_btn.setToolTip(tr("设置"))
        settings_btn.setFixedSize(bookmark_btn_width, titlebar_height)
        settings_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                color: #202020;
            }}
            QPushButton:hover {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:pressed {{
                background: #d5d5d5;
                color: #000000;
            }}
        """)
        settings_btn.clicked.connect(self.show_settings_menu)
        titlebar_layout.addWidget(settings_btn)

        # AI 助手面板切换按钮
        self.ai_chat_btn = QPushButton("🤖")
        self.ai_chat_btn.setToolTip(tr("AI 助手面板 (Ctrl+Shift+A)"))
        self.ai_chat_btn.setFixedSize(bookmark_btn_width, titlebar_height)
        self.ai_chat_btn.setCheckable(True)
        self.ai_chat_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: {btn_radius}px;
                font-size: {btn_font_size}pt;
                color: #202020;
            }}
            QPushButton:hover {{
                background: #e5e5e5;
                color: #000000;
            }}
            QPushButton:checked {{
                background: #BBDEFB;
                color: #1565C0;
            }}
            QPushButton:pressed {{
                background: #d5d5d5;
                color: #000000;
            }}
        """)
        self.ai_chat_btn.clicked.connect(self.toggle_chat_panel)
        # 初始化时同步按钮状态
        chat_config = self.config.get("ai_chat", {})
        panel_visible = chat_config.get("panel_visible", False)
        ai_enabled = chat_config.get("enabled", True)
        self.ai_chat_btn.setChecked(panel_visible)
        self.ai_chat_btn.setVisible(ai_enabled)
        titlebar_layout.addWidget(self.ai_chat_btn)
        
        # 系统原生标题栏已提供最小化/最大化/关闭按钮，无需自定义
        
        main_layout.addWidget(titlebar)
    
    def toggle_maximize(self):
        """切换最大化/还原窗口"""
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()
    
    def mousePressEvent(self, event):
        """鼠标按下事件"""
        # 处理菜单栏的右键点击
        if event.button() == Qt.RightButton:
            menubar = self.menu_bar
            menubar_rect = menubar.geometry()
            
            # 检查点击是否在菜单栏区域
            if menubar_rect.contains(event.pos()):
                # 转换为menubar的局部坐标
                local_pos = menubar.mapFrom(self, event.pos())
                action = menubar.actionAt(local_pos)
                
                debug_print(f"[DEBUG] MenuBar right click at {local_pos}, action: {action}")
                
                if action and hasattr(self, 'bookmark_actions') and action in self.bookmark_actions:
                    node = self.bookmark_actions[action]
                    bookmark_id = node.get('id')
                    bookmark_name = node.get('name', '')
                    
                    debug_print(f"[DEBUG] Found bookmark: {bookmark_name} (ID: {bookmark_id})")
                    
                    # 检查是否是特殊书签（不允许删除）
                    special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
                    is_special = any(bookmark_name.startswith(icon) for icon in special_icons)
                    
                    debug_print(f"[DEBUG] Is special bookmark: {is_special}")
                    
                    if not is_special:
                        global_pos = event.globalPos()
                        debug_print(f"[DEBUG] Showing context menu at: {global_pos}")
                        self.show_bookmark_context_menu(global_pos, bookmark_id, bookmark_name)
                        event.accept()
                        return
                else:
                    debug_print(f"[DEBUG] No bookmark action found")
        
        super().mousePressEvent(event)
    

    
    def nativeEvent(self, eventType, message):
        """处理 Windows 原生事件"""
        try:
            if eventType in (b"windows_generic_MSG", "windows_generic_MSG", b"windows_dispatcher_MSG", "windows_dispatcher_MSG"):
                from ctypes import wintypes, cast, POINTER
                import ctypes
                
                msg = cast(int(message), POINTER(wintypes.MSG)).contents

                # WM_SYSCOMMAND：从最小化恢复时设置 restore guard（抑制 IEB 虚假导航）
                if msg.message == 0x0112:  # WM_SYSCOMMAND
                    command = msg.wParam & 0xFFF0
                    SC_RESTORE = 0xF120
                    if command == SC_RESTORE and self.isMinimized():
                        self._set_restore_nav_guard()
        except Exception:
            pass
        
        return super().nativeEvent(eventType, message)
    
    def changeEvent(self, event):
        """窗口状态变化事件"""
        if event.type() == event.WindowStateChange:
            was_minimized = bool(event.oldState() & Qt.WindowMinimized)
            now_minimized = bool(self.windowState() & Qt.WindowMinimized)

            # 从最小化恢复：重新武装当前标签的刷新与路径同步（安全网）
            if was_minimized and not now_minimized:
                QTimer.singleShot(0, lambda: self._reactivate_current_tab_refresh(rebuild_pathbar=True))
        
        super().changeEvent(event)




    def event(self, event):
        """处理窗口事件"""
        if event.type() == event.WindowActivate:
            # 安全网：窗口重新激活时，若当前标签刷新被瞬态停掉则重新武装（健康时为空操作）
            self._reactivate_current_tab_refresh(rebuild_pathbar=False)
        
        return super().event(event)
    

    def _reactivate_current_tab_refresh(self, rebuild_pathbar=False):
        """窗口恢复/重新激活时的刷新自愈安全网。

        根因：set_refresh_active(True)（武装目录轮询/保活路径同步/消费待刷新）仅在
        on_tab_changed 与 close_tab 被调用，没有任何路径在窗口恢复/激活时重新武装当前
        标签。若可见标签的刷新定时器被某个瞬态停掉而标签索引未变化，就会出现“路径栏不
        更新、文件夹不刷新，需手动 resize/切标签才恢复”。此方法把该手动恢复自动化。

        健康时（刷新仍在运行）为空操作；仅在检测到确实失活时才重新武装并补一次刷新。
        """
        try:
            tab = self.get_current_tab_widget()
            if not tab:
                return
            current_path = getattr(tab, 'current_path', '') or ''
            is_slow = False
            try:
                is_slow = bool(tab._is_slow_path(current_path)) if current_path else False
            except Exception:
                is_slow = False
            poll = getattr(tab, 'dir_poll_timer', None)
            # 仅在真正失活时才重新武装：标签自认后台，或普通路径的目录轮询已停。
            # 慢速路径（OneDrive/网络）本就不启动目录轮询，不应误判为失活。
            needs_rearm = (
                not getattr(tab, '_refresh_active', False)
                or (poll is not None and not poll.isActive() and not is_slow)
            )
            if needs_rearm and hasattr(tab, 'set_refresh_active'):
                tab.set_refresh_active(True)
                if hasattr(tab, '_request_refresh'):
                    tab._request_refresh(reason="window_reactivate")
                debug_print("[Reactivate] Re-armed current tab refresh after window activation")

            if rebuild_pathbar:
                pb = getattr(tab, 'path_bar', None)
                if pb and current_path:
                    try:
                        if not getattr(pb, '_in_edit', False):
                            pb.set_path(current_path)
                    except Exception:
                        pass
        except Exception as e:
            debug_print(f"[Reactivate] failed: {e}")

    def _manual_statusbar_reflow_refresh(self):
        """底部状态栏双击触发：模拟一次 resize 级别的 UI 重排，并重新武装当前标签刷新。

        这个入口用于手动兜底“路径栏偶发不更新，但手动调整窗口大小后恢复”的情况。
        它不做导航，只强制当前标签路径栏重建、重排布局，并用 1px 临时 resize 触发 Qt/
        shell 宿主的几何刷新；最大化状态下不改窗口大小，只做布局刷新。
        """
        try:
            self._reactivate_current_tab_refresh(rebuild_pathbar=True)
            tab = self.get_current_tab_widget()
            if tab and hasattr(tab, '_manual_pathbar_rebuild'):
                tab._manual_pathbar_rebuild()

            self.updateGeometry()
            self.update()
            if not self.isMaximized() and not self.isMinimized():
                old_size = self.size()
                self.resize(old_size.width() + 1, old_size.height())
                QTimer.singleShot(0, lambda s=old_size: self.resize(s))
            debug_print("[StatusBar] Manual reflow refresh triggered")
        except Exception as e:
            debug_print(f"[StatusBar] manual reflow failed: {e}")
    
    def setup_shortcuts(self):
        """设置全局快捷键（现在使用轮询方式，不再使用QShortcut）"""
        # QShortcut 被 QAxWidget 拦截，所以现在使用定时器轮询方式
        # 保留此方法以便将来扩展或备用
        self.shortcuts = []
    
    def refresh_current_tab(self):
        """刷新当前标签页"""
        current_tab = self.get_active_pane()
        if hasattr(current_tab, 'current_path'):
            current_tab.navigate_to(current_tab.current_path, 
                                  is_shell=current_tab.current_path.startswith('shell:'))
    
    def add_current_tab_bookmark(self):
        """添加当前标签页到书签"""
        current_tab = self.get_active_pane()
        if current_tab:
            self.add_tab_bookmark(current_tab)

    def copy_selected_filename(self, mode="filename"):
        """
        复制当前选中文件名或路径+文件名到剪贴板并提示。
        mode: "filename" 只拷贝文件名，"path" 拷贝全路径+文件名。
        """
        current_tab = self.get_active_pane()
        names = []
        if current_tab and hasattr(current_tab, 'get_selected_filenames'):
            names = current_tab.get_selected_filenames()
        from PyQt5.QtWidgets import QApplication
        if names:
            if mode == "path":
                # 拷贝全路径+文件名，使用设置中定义的分隔符
                separator = self.config.get("breadcrumb_copy_separator", "/")
                # 获取当前路径，并用设置的分隔符替换所有反斜杠和正斜杠
                current_path = current_tab.current_path.replace("\\", "/").replace("/", separator)
                full_paths = [f"{current_path}{separator}{name}" for name in names]
                text = ", ".join(full_paths)
                QApplication.clipboard().setText(text)
                if len(full_paths) == 1:
                    show_toast(self, tr("复制成功"), tr("路径: {}").format(full_paths[0]), level="info")
                else:
                    show_toast(self, tr("复制成功"), f"已复制 {len(full_paths)} 个路径", level="info")
            else:
                # 只拷贝文件名
                filenames_text = ", ".join(names)
                QApplication.clipboard().setText(filenames_text)
                if len(names) == 1:
                    show_toast(self, tr("复制成功"), tr("文件名: {}").format(names[0]), level="info")
                else:
                    show_toast(self, tr("复制成功"), f"已复制 {len(names)} 个文件名", level="info")
        else:
            # 未选中文件时，复制路径栏地址，按设置分隔符
            separator = self.config.get("breadcrumb_copy_separator", "/")
            # 获取当前tab的path_bar
            path_bar = None
            if current_tab and hasattr(current_tab, 'path_bar'):
                path_bar = current_tab.path_bar
            elif hasattr(self, 'path_bar'):
                path_bar = self.path_bar
            if path_bar and hasattr(path_bar, 'get_path_for_copy'):
                path_text = path_bar.get_path_for_copy(separator)
                QApplication.clipboard().setText(path_text)
                show_toast(self, tr("复制成功"), tr("路径: {}").format(path_text), level="info")
            else:
                show_toast(self, tr("提示"), tr("未选中文件，也无法获取路径栏地址"), level="warning")

    def quick_find_in_current_directory(self):
        """通过关键字快速检索当前目录下的文件或文件夹名，并在当前目录中选中目标。"""
        current_tab = self.get_active_pane()
        if not current_tab or not hasattr(current_tab, 'current_path'):
            show_toast(self, tr("提示"), tr("当前没有可用的标签页路径"), level="warning")
            return

        search_root = current_tab.current_path
        if not isinstance(search_root, str) or not search_root or search_root.startswith('shell:') or '::' in search_root:
            show_toast(self, tr("提示"), tr("当前路径不支持快捷检索"), level="warning")
            return
        if not os.path.isdir(search_root):
            show_toast(self, tr("提示"), tr("当前目录无效"), level="warning")
            return

        keyword, ok = QInputDialog.getText(self, tr("快捷定位"), tr("请输入要检索的文件或文件夹关键字："))
        self._guard_shortcuts_after_modal()
        if not ok:
            return

        keyword = keyword.strip()
        if not keyword:
            show_toast(self, tr("提示"), tr("请输入搜索关键词"), level="warning")
            return

        keyword_lower = keyword.lower()
        matched_paths = []
        max_matches = 200

        try:
            with os.scandir(search_root) as entries:
                sorted_entries = sorted(entries, key=lambda entry: entry.name.lower())
                for entry in sorted_entries:
                    if keyword_lower in entry.name.lower():
                        matched_paths.append(entry.path)
                        if len(matched_paths) >= max_matches:
                            break
        except Exception as e:
            show_toast(self, tr("错误"), tr("快捷检索失败: {}").format(e), level="error")
            return

        if not matched_paths:
            show_toast(self, tr("提示"), tr("当前目录下未找到包含“{}”的文件或文件夹名").format(keyword), level="warning")
            return

        matched_paths.sort(key=lambda item: os.path.basename(item).lower())
        selected_path = matched_paths[0]
        if len(matched_paths) > 1:
            picker = QuickFindResultsDialog(matched_paths, self)
            picker.setWindowTitle(f"选择匹配项（共 {len(matched_paths)} 项）")
            ok = picker.exec_()
            self._guard_shortcuts_after_modal()
            if not ok or not picker.selected_path:
                return
            selected_path = picker.selected_path

        selected_name = os.path.basename(selected_path)
        current_tab.select_file_in_explorer(selected_name)
        item_type = tr("文件夹") if os.path.isdir(selected_path) else tr("文件")
        show_toast(self, tr("快捷定位"), tr("已在当前目录选中{}: {}").format(item_type, selected_name), level="info")
    
    def keyPressEvent(self, event):
        """处理快捷键（备用方案，主要使用QShortcut）"""
        # 保留此方法以防QShortcut在某些情况下不工作
        super().keyPressEvent(event)

    def _guard_shortcuts_after_modal(self, cooldown_ms=250):
        """模态输入框关闭后，短暂屏蔽轮询快捷键，避免把输入过程误判为组合键。"""
        self._last_keys_state.clear()
        # 立即消耗所有非修饰键的 GetAsyncKeyState 粘性位（bit0）。
        # 场景：用户在弹窗内输入了包含快捷键字符的关键词（如 "em_rtc" 含 't'），
        # 弹窗关闭后该粘性位残留，下次用户恰好按住 Ctrl 时会产生幽灵 Ctrl+T 触发。
        try:
            _drain = ctypes.windll.user32.GetAsyncKeyState
            for _vk in (0x5A, 0x58, 0x4C, 0x54, 0x41, 0x57, 0x46, 0x47,
                        0x44, 0x09, 0x25, 0x27, 0x26, 0x28, 0x74):
                _drain(_vk)
        except Exception:
            pass
        self._shortcut_modal_guard_until = time.monotonic() + max(0, cooldown_ms) / 1000.0
        self._shortcut_wait_for_modifier_release = True
    
    def eventFilter(self, obj, event):
        """应用级别的事件过滤器（暂时不使用，因为被QAxWidget拦截）"""
        # 由于QAxWidget在底层拦截事件，eventFilter接收不到事件
        # 现在使用定时器轮询方式处理快捷键
        return super().eventFilter(obj, event)
    
    def _check_shortcuts(self):
        """定时检查快捷键状态（用于检测被QAxWidget拦截的快捷键）"""
        try:
            # 当焦点在文本输入控件时，避免全局快捷键轮询干扰输入
            from PyQt5.QtWidgets import QApplication, QLineEdit, QTextEdit, QPlainTextEdit
            fw = QApplication.focusWidget()
            if isinstance(fw, (QLineEdit, QTextEdit, QPlainTextEdit)):
                self._last_keys_state.clear()
                if self._shortcut_timer.interval() != SHORTCUT_POLL_INACTIVE_MS:
                    self._shortcut_timer.setInterval(SHORTCUT_POLL_INACTIVE_MS)
                return

            # 窗口激活检查：以 Windows API 进程归属为唯一可靠依据。
            # 原因：QApplication.activeWindow() 和 isActiveWindow() 在嵌入的
            # Shell.Explorer ActiveX 控件持有 Win32 焦点时会返回 None/False，
            # 导致误判为"窗口不活跃"→ 降频至 500ms → 快捷键漏检或延迟。
            import ctypes
            foreground_hwnd = ctypes.windll.user32.GetForegroundWindow()
            fg_same_process = False
            if foreground_hwnd:
                fg_pid = ctypes.c_ulong(0)
                ctypes.windll.user32.GetWindowThreadProcessId(foreground_hwnd, ctypes.byref(fg_pid))
                fg_same_process = (fg_pid.value == os.getpid())

            if not fg_same_process:
                # 前台窗口属于其他进程 → 降频并消耗所有非修饰键的粘性位，
                # 防止切换回来时因积累的 GetAsyncKeyState bit0 产生幽灵触发。
                self._last_keys_state.clear()
                if self._shortcut_timer.interval() != SHORTCUT_POLL_INACTIVE_MS:
                    self._shortcut_timer.setInterval(SHORTCUT_POLL_INACTIVE_MS)
                for _vk in (0x5A, 0x58, 0x4C, 0x54, 0x41, 0x57, 0x46, 0x47, 0x44, 0x09, 0x25, 0x27, 0x26, 0x74):
                    ctypes.windll.user32.GetAsyncKeyState(_vk)
                return

            # 前台属于本进程。若是本进程的其他 Qt 顶层窗口（如非模态搜索对话框）
            # 持有焦点，也应暂停快捷键响应，避免后台误触。
            active_win = QApplication.activeWindow()
            if active_win is not None and active_win is not self:
                self._last_keys_state.clear()
                if self._shortcut_timer.interval() != SHORTCUT_POLL_INACTIVE_MS:
                    self._shortcut_timer.setInterval(SHORTCUT_POLL_INACTIVE_MS)
                return

            # 本进程主窗口或 Shell.Explorer 子窗口处于前台 → 启用高频轮询
            if self._shortcut_timer.interval() != SHORTCUT_POLL_ACTIVE_MS:
                self._shortcut_timer.setInterval(SHORTCUT_POLL_ACTIVE_MS)
            
            # Windows虚拟键码
            VK_CONTROL = 0x11
            VK_SHIFT = 0x10
            VK_MENU = 0x12  # Alt键
            VK_F5 = 0x74
            
            # 获取键盘状态
            # 组合键必须“同时物理按住”才触发：只信任 GetAsyncKeyState 的高位
            # (0x8000，表示当前正被按住)，不再使用粘滞位 (0x0001，表示自上次查询
            # 以来按过一次)。粘滞位会导致“先按松某键后再按修饰键”被误判为组合键。
            # require_down 参数保留以兼容调用点，但行为统一为“必须当前按住”。
            _GetAsyncKeyState = ctypes.windll.user32.GetAsyncKeyState

            def is_key_pressed(vk_code, require_down=False):
                state = _GetAsyncKeyState(vk_code)
                return (state & 0x8000) != 0

            modifiers_down = (
                is_key_pressed(VK_CONTROL, require_down=True)
                or is_key_pressed(VK_SHIFT, require_down=True)
                or is_key_pressed(VK_MENU, require_down=True)
            )

            if getattr(self, '_shortcut_wait_for_modifier_release', False):
                if modifiers_down:
                    return
                self._shortcut_wait_for_modifier_release = False

            if time.monotonic() < getattr(self, '_shortcut_modal_guard_until', 0):
                # 守卫期间每轮都消耗一次非修饰键粘性位，防止守卫窗口内新产生的按键残留。
                for _vk in (0x5A, 0x58, 0x4C, 0x54, 0x41, 0x57, 0x46, 0x47,
                            0x44, 0x09, 0x25, 0x27, 0x26, 0x28, 0x74):
                    _GetAsyncKeyState(_vk)
                return
            
            hotkeys = self.config.get("hotkeys", {})
            
            # Alt+Z - 拷贝文件名
            if is_key_pressed(VK_MENU, require_down=True) and is_key_pressed(0x5A) and hotkeys.get("copy_filename", True):
                key_combo = "Alt+Z"
                if not self._last_keys_state.get(key_combo, False):
                    debug_print("[Shortcut Poll] Detected Alt+Z")
                    self.copy_selected_filename(mode="filename")
                    self._last_keys_state[key_combo] = True
                return
            else:
                self._last_keys_state["Alt+Z"] = False
            
            # Alt+X - 拷贝路径+文件名
            if is_key_pressed(VK_MENU, require_down=True) and is_key_pressed(0x58) and hotkeys.get("copy_filepath", True):
                key_combo = "Alt+X"
                if not self._last_keys_state.get(key_combo, False):
                    debug_print("[Shortcut Poll] Detected Alt+X")
                    self.copy_selected_filename(mode="path")
                    self._last_keys_state[key_combo] = True
                return
            else:
                self._last_keys_state["Alt+X"] = False
            
            # 检查Ctrl组合键
            if is_key_pressed(VK_CONTROL, require_down=True):
                # Ctrl+L (0x4C) - 聚焦并全选路径栏，便于直接输入地址
                if is_key_pressed(0x4C):
                    key_combo = "Ctrl+L"
                    if not self._last_keys_state.get(key_combo, False):
                        debug_print("[Shortcut Poll] Detected Ctrl+L")
                        current_tab = self.get_active_pane()
                        if current_tab and hasattr(current_tab, 'path_bar') and current_tab.path_bar:
                            current_tab.path_bar.enter_edit_mode()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+L"] = False

                # Ctrl+Shift+T (0x54) - 恢复关闭的标签页（必须在Ctrl+T之前检测）
                if is_key_pressed(VK_SHIFT, require_down=True) and is_key_pressed(0x54) and hotkeys.get("reopen_tab", True):
                    key_combo = "Ctrl+Shift+T"
                    if not self._last_keys_state.get(key_combo, False):
                        debug_print("[Shortcut Poll] Detected Ctrl+Shift+T")
                        self.reopen_closed_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+Shift+T"] = False

                    # Ctrl+Shift+A (0x41) - 切换 AI 聊天面板
                    if is_key_pressed(VK_SHIFT, require_down=True) and is_key_pressed(0x41):
                        key_combo = "Ctrl+Shift+A"
                        if not self._last_keys_state.get(key_combo, False):
                            debug_print("[Shortcut Poll] Detected Ctrl+Shift+A")
                            self.toggle_chat_panel()
                            self._last_keys_state[key_combo] = True
                        return
                    else:
                        self._last_keys_state["Ctrl+Shift+A"] = False
                
                # Ctrl+T (0x54) - 新建标签页（不包含Shift）
                if is_key_pressed(0x54) and not is_key_pressed(VK_SHIFT, require_down=True) and hotkeys.get("new_tab", True):
                    key_combo = "Ctrl+T"
                    if not self._last_keys_state.get(key_combo, False):
                        debug_print("[Shortcut Poll] Detected Ctrl+T")
                        self.add_new_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+T"] = False
                
                # Ctrl+W (0x57)
                if is_key_pressed(0x57) and hotkeys.get("close_tab", True):
                    key_combo = "Ctrl+W"
                    if not self._last_keys_state.get(key_combo, False):
                        debug_print("[Shortcut Poll] Detected Ctrl+W")
                        self.close_current_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+W"] = False
                
                # Ctrl+F (0x46)
                if is_key_pressed(0x46) and hotkeys.get("search", True):
                    key_combo = "Ctrl+F"
                    if not self._last_keys_state.get(key_combo, False):
                        debug_print("[Shortcut Poll] Detected Ctrl+F")
                        self.show_search_dialog()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+F"] = False

                # Ctrl+G (0x47) - 快速检索当前目录中的文件/文件夹名
                if is_key_pressed(0x47) and hotkeys.get("quick_find_current_dir", True):
                    key_combo = "Ctrl+G"
                    if not self._last_keys_state.get(key_combo, False):
                        debug_print("[Shortcut Poll] Detected Ctrl+G")
                        self.quick_find_in_current_directory()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+G"] = False
                
                # Ctrl+D (0x44)
                if is_key_pressed(0x44) and hotkeys.get("add_bookmark", True):
                    key_combo = "Ctrl+D"
                    if not self._last_keys_state.get(key_combo, False):
                        debug_print("[Shortcut Poll] Detected Ctrl+D")
                        self.add_current_tab_bookmark()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+D"] = False
                
                # Ctrl+Tab (0x09)
                if is_key_pressed(0x09) and hotkeys.get("switch_tab", True):
                    if is_key_pressed(VK_SHIFT, require_down=True):
                        # Ctrl+Shift+Tab
                        key_combo = "Ctrl+Shift+Tab"
                        if not self._last_keys_state.get(key_combo, False):
                            debug_print("[Shortcut Poll] Detected Ctrl+Shift+Tab")
                            self.tab_widget.setCurrentIndex(
                                (self.tab_widget.currentIndex() - 1) % self.tab_widget.count())
                            self._last_keys_state[key_combo] = True
                        return
                    else:
                        # Ctrl+Tab
                        key_combo = "Ctrl+Tab"
                        if not self._last_keys_state.get(key_combo, False):
                            debug_print("[Shortcut Poll] Detected Ctrl+Tab")
                            self.tab_widget.setCurrentIndex(
                                (self.tab_widget.currentIndex() + 1) % self.tab_widget.count())
                            self._last_keys_state[key_combo] = True
                        return
                else:
                    self._last_keys_state["Ctrl+Tab"] = False
                    self._last_keys_state["Ctrl+Shift+Tab"] = False
            
            # 检查Alt组合键
            if is_key_pressed(VK_MENU, require_down=True):
                # Alt+Left (0x25)
                if is_key_pressed(0x25) and hotkeys.get("navigate", True):
                    key_combo = "Alt+Left"
                    if not self._last_keys_state.get(key_combo, False):
                        self.go_back_current_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Alt+Left"] = False
                
                # Alt+Right (0x27)
                if is_key_pressed(0x27) and hotkeys.get("navigate", True):
                    key_combo = "Alt+Right"
                    if not self._last_keys_state.get(key_combo, False):
                        self.go_forward_current_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Alt+Right"] = False
                
                # Alt+Up (0x26)
                if is_key_pressed(0x26) and hotkeys.get("go_up", True):
                    key_combo = "Alt+Up"
                    if not self._last_keys_state.get(key_combo, False):
                        self.go_up_current_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Alt+Up"] = False
            
            # 检查F5
            if is_key_pressed(VK_F5) and hotkeys.get("refresh", True):
                key_combo = "F5"
                if not self._last_keys_state.get(key_combo, False):
                    self.refresh_current_tab()
                    self._last_keys_state[key_combo] = True
                return
            else:
                self._last_keys_state["F5"] = False

            # F3 - 左右分屏对比
            if is_key_pressed(0x72) and hotkeys.get("split_view", True):
                key_combo = "F3"
                if not self._last_keys_state.get(key_combo, False):
                    self.toggle_split_view()
                    self._last_keys_state[key_combo] = True
                return
            else:
                self._last_keys_state["F3"] = False

        except Exception as e:
            # 如果轮询出错，不影响程序运行
            pass
    
    def show_settings_menu(self):
        """显示设置对话框"""
        self.settings_dialog = SettingsDialog(self.config, self)
        dlg = self.settings_dialog
        result = dlg.exec_()
        if result:
            # 获取新配置
            old_monitor = self.config.get("enable_explorer_monitor", True)
            old_interval = self.config.get("explorer_monitor_interval", 2.0)

            new_monitor = dlg.monitor_cb.isChecked()
            new_interval = dlg.interval_spinbox.value()

            # 更新配置
            self.config["enable_explorer_monitor"] = new_monitor
            self.config["debug_mode"] = dlg.debug_mode_cb.isChecked()
            self.config["explorer_monitor_interval"] = new_interval
            self.config["enable_cache_tabs"] = dlg.cache_tabs_cb.isChecked()
            self.config["enable_tortoisegit_buttons"] = dlg.tortoisegit_buttons_cb.isChecked()
            self.config["preferred_terminal_tool"] = normalize_terminal_tool_name(dlg.preferred_terminal_combo.currentData())
            self.config["enable_title_shortcuts"] = dlg.title_shortcuts_cb.isChecked()
            self.config["enable_mouse_gestures"] = dlg.mouse_gestures_cb.isChecked()

            # 更新全局调试开关
            set_debug_mode(self.config["debug_mode"])

            # 更新快捷键配置
            if "hotkeys" not in self.config:
                self.config["hotkeys"] = {}
            self.config["hotkeys"]["new_tab"] = dlg.hotkey_new_tab.isChecked()
            self.config["hotkeys"]["close_tab"] = dlg.hotkey_close_tab.isChecked()
            self.settings_dialog = None
            self.config["hotkeys"]["reopen_tab"] = dlg.hotkey_reopen_tab.isChecked()
            self.config["hotkeys"]["switch_tab"] = dlg.hotkey_switch_tab.isChecked()
            self.config["hotkeys"]["search"] = dlg.hotkey_search.isChecked()
            self.config["hotkeys"]["navigate"] = dlg.hotkey_navigate.isChecked()
            self.config["hotkeys"]["go_up"] = dlg.hotkey_go_up.isChecked()
            self.config["hotkeys"]["refresh"] = dlg.hotkey_refresh.isChecked()
            self.config["hotkeys"]["add_bookmark"] = dlg.hotkey_add_bookmark.isChecked()
            self.config["hotkeys"]["copy_filename"] = dlg.hotkey_copy_filename.isChecked()
            self.config["hotkeys"]["quick_find_current_dir"] = dlg.hotkey_quick_find_current_dir.isChecked()
            self.config["hotkeys"]["split_view"] = dlg.hotkey_split_view.isChecked()
            
            self.save_config()

            # 同步标题栏按钮可见性
            self.apply_tortoisegit_buttons_config()
            self.apply_title_shortcuts_config()
            
            # 重新设置快捷键
            # 清除旧的快捷键
            for shortcut in getattr(self, 'shortcuts', []):
                shortcut.setEnabled(False)
                shortcut.deleteLater()
            self.shortcuts = []
            # 重新创建快捷键
            self.setup_shortcuts()
            
            # 如果监听状态或间隔改变，重启监听
            if old_monitor != new_monitor or (new_monitor and old_interval != new_interval):
                if old_monitor:
                    self.stop_explorer_monitor()
                if new_monitor:
                    self.monitor_interval = new_interval
                    self.start_explorer_monitor()
        
    def show_bookmark_dialog(self):
        dlg = BookmarkDialog(self.bookmark_manager, self)
        dlg.exec_()
    
    def check_for_updates(self):
        """打开 GitHub Releases 页面检查更新"""
        import webbrowser
        try:
            webbrowser.open("https://github.com/caojinyuan/TabEx/releases")
            show_toast(self, tr("检查更新"), tr("已在浏览器中打开更新页面"), level="info")
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开浏览器: {}").format(e), level="error")
    
    def show_search_dialog(self):
        """显示搜索对话框（非模态）"""
        current_tab = self.get_active_pane()
        if not current_tab or not hasattr(current_tab, 'current_path'):
            show_toast(self, tr("提示"), tr("请先打开一个文件夹"), level="warning")
            self.setFocus()
            return
        
        search_path = current_tab.current_path
        
        # 不支持搜索特殊路径
        if search_path.startswith('shell:'):
            show_toast(self, tr("提示"), tr("不支持搜索特殊路径（shell:）"), level="warning")
            self.setFocus()
            return
        
        if not os.path.exists(search_path):
            show_toast(self, tr("提示"), tr("路径不存在: {}").format(search_path), level="warning")
            self.setFocus()
            return
        
        # 创建非模态对话框，传入搜索历史
        dlg = SearchDialog(search_path, self, self.search_history)
        # 恢复上次的大小和位置
        dlg_geo = self.config.get("search_dialog_geometry")
        if dlg_geo and isinstance(dlg_geo, dict):
            try:
                dlg.resize(dlg_geo.get("w", 800), dlg_geo.get("h", 500))
                x, y = dlg_geo.get("x"), dlg_geo.get("y")
                if x is not None and y is not None:
                    dlg.move(x, y)
            except Exception:
                pass
        # 关闭时保存大小和位置
        def _save_geo():
            geo = dlg.geometry()
            self.config["search_dialog_geometry"] = {
                "x": geo.x(), "y": geo.y(),
                "w": geo.width(), "h": geo.height()
            }
            self.save_config()
        dlg.finished.connect(lambda _: _save_geo())
        # 保存对话框引用，防止被垃圾回收
        if not hasattr(self, 'search_dialogs'):
            self.search_dialogs = []
        self.search_dialogs.append(dlg)
        
        # 对话框关闭时从列表中移除（注意: finished信号带int参数，需要兼容lambada）
        dlg.finished.connect(lambda result: self.search_dialogs.remove(dlg) if dlg in self.search_dialogs else None)
        
        # 非模态显示，不阻塞主窗口
        dlg.show()
    
    def add_search_history(self, keyword):
        """添加搜索关键词到历史记录（使用配置的最大值）"""
        if not keyword or not keyword.strip():
            return
        
        keyword = keyword.strip()
        
        # 如果已存在，先移除（避免重复）
        if keyword in self.search_history:
            self.search_history.remove(keyword)
        
        # 添加到列表开头（最新的在前面）
        self.search_history.insert(0, keyword)
        
        # 限制最多保留配置的数量（内存优化）
        if len(self.search_history) > self.max_search_history:
            self.search_history = self.search_history[:self.max_search_history]
        
        # 持久化搜索历史到 config.json
        self.config["search_history"] = self.search_history
        self.save_config()

    def tab_context_menu(self, pos, target_tabwidget=None):
        tw, cs, is_right = self._resolve_group(target_tabwidget)
        tab_index = tw.tabBar().tabAt(pos)
        if tab_index < 0:
            return
        tab = cs.widget(tab_index) if cs is not None else None
        is_pinned = hasattr(tab, 'is_pinned') and tab.is_pinned
        menu = QMenu()
        # 图标可用emoji或标准QIcon
        if is_pinned:
            pin_action = QAction(tr("🔨 取消固定"), self)
            pin_action.triggered.connect(lambda: self.unpin_tab(tab_index, tw))
            menu.addAction(pin_action)
        else:
            pin_action = QAction(tr("📌 固定"), self)
            pin_action.triggered.connect(lambda: self.pin_tab(tab_index, tw))
            menu.addAction(pin_action)

        # 添加“添加书签”菜单项，使用书签emoji
        add_bm_action = QAction(tr("📑 添加书签"), self)
        add_bm_action.triggered.connect(lambda: self.add_tab_bookmark(tab))
        menu.addAction(add_bm_action)

        if hasattr(tab, 'set_auto_refresh_frozen'):
            if tab.is_auto_refresh_frozen():
                refresh_action = QAction(tr("▶ 恢复自动刷新"), self)
                refresh_action.triggered.connect(lambda: self.toggle_tab_auto_refresh(tab_index, False, tw))
            else:
                refresh_action = QAction(tr("⏸ 冻结自动刷新"), self)
                refresh_action.triggered.connect(lambda: self.toggle_tab_auto_refresh(tab_index, True, tw))
            menu.addAction(refresh_action)

        menu.exec_(tw.tabBar().mapToGlobal(pos))

    def toggle_tab_auto_refresh(self, tab_index, frozen, target_tabwidget=None):
        _tw, cs, _is_right = self._resolve_group(target_tabwidget)
        tab = cs.widget(tab_index) if cs is not None else None
        if not tab or not hasattr(tab, 'set_auto_refresh_frozen'):
            return
        is_frozen = tab.set_auto_refresh_frozen(frozen)
        show_toast(
            self,
            tr("自动刷新"),
            tr("当前标签自动刷新已冻结") if is_frozen else tr("当前标签自动刷新已恢复"),
            level="info",
            duration=1800,
        )

    def add_tab_bookmark(self, tab):
        # 选择父文件夹
        bm = self.bookmark_manager
        tree = bm.get_tree()
        folder_list = []
        def collect_folders(node):
            if isinstance(node, dict):
                if node.get('type') == 'folder':
                    folder_list.append((node.get('id'), node.get('name')))
                    for child in node.get('children', []):
                        collect_folders(child)
            elif isinstance(node, list):
                for item in node:
                    collect_folders(item)
        for root in tree.values():
            collect_folders(root)
        if not folder_list:
            show_toast(self, tr("无可用书签文件夹"), tr("请先在 bookmarks.json 中添加至少一个文件夹。"), level="warning")
            return
        # 选择父文件夹
        folder_names = [f"{name} (id:{fid})" for fid, name in folder_list]
        from PyQt5.QtWidgets import QInputDialog
        idx, ok = QInputDialog.getItem(self, tr("选择书签文件夹"), tr("请选择父文件夹："), folder_names, 0, False)
        # 对话框关闭后，强制将焦点设回主窗口，防止QAxWidget拦截快捷键
        self.setFocus()
        if not ok:
            return
        folder_id = folder_list[folder_names.index(idx)][0]
        # 输入书签名称
        name, ok = QInputDialog.getText(self, tr("书签名称"), tr("请输入书签名称："), text=os.path.basename(tab.current_path))
        # 对话框关闭后，强制将焦点设回主窗口
        self.setFocus()
        if not ok or not name:
            return
        # 保存到 bookmarks.json
        url = "file:///" + tab.current_path.replace("\\", "/")
        if bm.add_bookmark(folder_id, name, url):
            self.populate_bookmark_bar_menu()
        else:
            show_toast(self, tr("添加失败"), tr("未能添加书签，请检查父文件夹。"), level="warning")

    def pin_tab(self, tab_index, target_tabwidget=None):
        tw, cs, _is_right = self._resolve_group(target_tabwidget)
        tab = cs.widget(tab_index) if cs is not None else None
        if tab is None:
            return
        tab.is_pinned = True
        # 重新排序：所有固定的在最左侧（仅作用于该标签组）
        self.sort_tabs_by_pinned(tw)
        self.save_pinned_tabs()

    def unpin_tab(self, tab_index, target_tabwidget=None):
        tw, cs, _is_right = self._resolve_group(target_tabwidget)
        tab = cs.widget(tab_index) if cs is not None else None
        if tab is None:
            return
        tab.is_pinned = False
        self.sort_tabs_by_pinned(tw)
        self.save_pinned_tabs()

    def sort_tabs_by_pinned(self, target_tabwidget=None):
        tw, cs, _is_right = self._resolve_group(target_tabwidget)
        if cs is None:
            return
        pinned = []
        unpinned = []
        # 记录当前tab对象
        current_index = tw.currentIndex()
        current_tab = cs.widget(current_index) if current_index >= 0 else None
        for i in range(tw.count()):
            tab = cs.widget(i)
            if hasattr(tab, 'is_pinned') and tab.is_pinned:
                pinned.append(tab)
            else:
                unpinned.append(tab)
        tw.clear()
        # 清空该组 content_stack
        while cs.count() > 0:
            widget = cs.widget(0)
            cs.removeWidget(widget)
        new_tabs = pinned + unpinned
        for tab in new_tabs:
            # 先添加标签页（临时标题）- 占位widget
            tw.addTab(QWidget(), "")
            # 将实际内容添加到 content_stack
            cs.addWidget(tab)
            # 然后调用update_tab_title更新标题（会考虑shell路径映射和图标）
            tab.update_tab_title()
        # 恢复原先的tab焦点
        if current_tab is not None:
            for i, tab in enumerate(new_tabs):
                if tab is current_tab:
                    tw.setCurrentIndex(i)
                    break

    def save_pinned_tabs(self):
        """保存固定标签页到config.json（扫描左右两个标签组）"""
        pinned_paths = []
        for _tw, cs in self._all_groups():
            if cs is None:
                continue
            for i in range(cs.count()):
                tab = cs.widget(i)
                if tab and getattr(tab, 'is_pinned', False) and hasattr(tab, 'current_path'):
                    pinned_paths.append(tab.current_path)

        # 更新config并保存
        self.config["pinned_tabs"] = pinned_paths
        self.save_config()
        self._schedule_session_snapshot()
        
        print(f"[Config] Saved {len(pinned_paths)} pinned tabs to config.json")

    def load_pinned_tabs(self):
        """从config.json加载固定标签页"""
        has_pinned = False
        
        # 从config.json读取
        pinned_paths = self.config.get("pinned_tabs", [])
        
        if pinned_paths:
            print(f"[Config] Loading {len(pinned_paths)} pinned tabs from config.json")
            for path in pinned_paths:
                if os.path.exists(path) or path.startswith('shell:'):
                    try:
                        is_shell = path.startswith('shell:')
                        # 懒加载：固定标签也延迟首次导航，避免启动瞬间多个 Shell 视图同时创建
                        tab = FileExplorerTab(self, path, is_shell=is_shell, defer_nav=True)
                        tab.is_pinned = True
                        short = path[-12:] if len(path) > 12 else path
                        pin_prefix = "📌"
                        title = pin_prefix + short
                        # 同时添加到 tab_widget 和 content_stack
                        self.tab_widget.addTab(QWidget(), title)  # tab_widget 只显示标签，内容用占位widget
                        self.content_stack.addWidget(tab)  # 实际内容添加到 content_stack
                        has_pinned = True
                        print(f"[Config] ✓ Loaded pinned tab: {path}")
                    except Exception as e:
                        print(f"[Config] ✗ Failed to load pinned tab {path}: {e}")
                else:
                    print(f"[Config] ⚠ Skipping non-existent path: {path}")
        else:
            print("[Config] No pinned tabs found in config.json")
        
        return has_pinned

    def __init__(self, parent=None):
        super().__init__(parent)

        self.server_socket = None
        self.server_thread = None
        self.monitor_thread = None
        self.server_running = False
        self.explorer_monitoring = False
        self.known_explorer_windows = set()
        self.last_check_time = 0
        
        # 启用主窗口拖拽支持
        self.setAcceptDrops(True)
        
        # 加载配置
        self.config = self.load_config()
        
        # 初始化全局调试开关
        set_debug_mode(self.config.get("debug_mode", False))
        set_explorer_monitor_debug(self.config.get("explorer_monitor_debug", False))
        
        # 初始化书签管理器
        self.bookmark_manager = BookmarkManager()
        # 检查并自动添加常用书签
        self.ensure_default_bookmarks()
        
        
        # 搜索历史（持久化到config.json）- 使用常量限制大小
        self.search_history = list(self.config.get("search_history", []))[:MAX_SEARCH_HISTORY]
        self.max_search_history = MAX_SEARCH_HISTORY
        
        # 关闭标签页历史 - 使用常量限制大小
        self.closed_tabs_history = []  # 每项格式: {'path': str, 'title': str, 'is_shell': bool}
        self.max_closed_tabs_history = MAX_CLOSED_TABS_HISTORY

        # 窗口恢复状态标记（用于在标题上显示“窗口恢复中”）
        self.is_restoring = False
        self._restore_title_suffix = tr(" - 窗口恢复中")
        
        # 性能优化：延迟初始化UI（先显示基本界面）
        self.init_ui()
        
        # 根据配置显示/隐藏 TortoiseGit 按钮
        self.apply_tortoisegit_buttons_config()
        # 根据配置显示/隐藏标题栏快捷方式区域
        self.apply_title_shortcuts_config()
        
        # 设置快捷键（在init_ui之后，确保所有组件已创建）
        self.setup_shortcuts()
        
        # 安装应用级别的事件过滤器，确保快捷键始终有效
        from PyQt5.QtWidgets import QApplication
        from PyQt5.QtCore import QTimer
        QApplication.instance().installEventFilter(self)
        
        # 启动快捷键轮询定时器（用于检测被QAxWidget拦截的快捷键）
        self._last_keys_state = {}
        self._shortcut_timer = QTimer(self)
        self._shortcut_timer.timeout.connect(self._check_shortcuts)
        self._shortcut_timer.start(SHORTCUT_POLL_ACTIVE_MS)

        # 运行期间持续保存标签会话，异常退出后仍可恢复最近窗口列表。
        self._session_snapshot_debounce_timer = QTimer(self)
        self._session_snapshot_debounce_timer.setSingleShot(True)
        self._session_snapshot_debounce_timer.timeout.connect(self.save_session_snapshot)
        self._session_snapshot_timer = QTimer(self)
        self._session_snapshot_timer.timeout.connect(self.save_session_snapshot)
        self._session_snapshot_timer.start(SESSION_SNAPSHOT_INTERVAL_MS)
        self._housekeeping_runs = 0
        self._housekeeping_timer = QTimer(self)
        self._housekeeping_timer.timeout.connect(self._run_housekeeping)
        self._housekeeping_timer.start(self._get_housekeeping_interval_ms())

        # 状态栏右侧 CPU/内存占用显示（默认关闭，可在设置中开启）
        self._resource_usage_timer = QTimer(self)
        self._resource_usage_timer.timeout.connect(self._update_resource_usage_display)
        self.apply_resource_usage_config()

        
        # 性能优化：延迟加载非关键功能（100ms后加载）
        QTimer.singleShot(100, self._delayed_initialization)

        # 生成并设置 TE 窗口图标：延迟到事件循环启动后执行，
        # 避免 9 张抗锯齿图标（256→16px）的渲染阻塞首帧显示，加快启动感知速度。
        QTimer.singleShot(0, self._setup_window_icon)

    def _setup_window_icon(self):
        """生成并设置 TE 窗口图标（延迟执行，不阻塞启动首帧）。"""
        try:
            from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QIcon
            
            te_icon = QIcon()
            # 按 256px 基准等比例生成各尺寸图标
            for size in [256, 128, 96, 64, 48, 32, 24, 18, 16]:
                pix = QPixmap(size, size)
                pix.fill(Qt.transparent)
                p = QPainter(pix)
                p.setRenderHint(QPainter.Antialiasing)
                blue = QColor("#2196F3")
                white = QColor("white")
                # 外层蓝色圆角背景
                outer_radius = max(2, size * 40 // 256)
                p.setBrush(blue)
                p.setPen(Qt.NoPen)
                p.drawRoundedRect(0, 0, size, size, outer_radius, outer_radius)
                # 内层白色圆角容器（形成蓝色边框效果）
                margin = max(2, size * 28 // 256)
                inner_radius = max(2, size * 24 // 256)
                p.setBrush(white)
                p.drawRoundedRect(margin, margin, size - 2*margin, size - 2*margin, inner_radius, inner_radius)
                # 中央蓝色 TE 文字
                p.setPen(blue)
                f = QFont()
                f.setBold(True)
                f.setPointSize(max(5, size * 130 // 256))
                f.setStretch(70)  # 压窄字体，使其看起来更高
                p.setFont(f)
                p.drawText(pix.rect(), Qt.AlignCenter, "TE")
                p.end()
                te_icon.addPixmap(pix, QIcon.Normal)
            
            self.setWindowIcon(te_icon)
            # 同时设置到 QApplication，使任务栏也生效
            from PyQt5.QtWidgets import QApplication
            QApplication.instance().setWindowIcon(te_icon)
            # 更新自定义标题栏中的图标 Label
            if hasattr(self, '_title_icon_label'):
                self._title_icon_label.setPixmap(te_icon.pixmap(18, 18))
            print("[Icon] ✓ TE icon set on window and application")
        except Exception as e:
            print(f"[Icon] ✗ Failed to set TE icon: {e}")
    
    def load_config(self):
        """加载配置文件"""
        default_config = {
            "enable_explorer_monitor": True,  # 默认启用Explorer监听
            "debug_mode": False,  # 默认关闭调试输出
            "explorer_monitor_debug": False,  # 默认关闭Explorer Monitor调试输出
            "resource_snapshot_logging": False,  # 默认关闭运行资源快照日志
            "resource_snapshot_interval_ms": HOUSEKEEPING_INTERVAL_MS,
            "show_resource_usage_in_statusbar": False,  # 默认关闭状态栏右侧 CPU/内存占用显示
            "pinned_tabs": [],  # 默认没有固定标签页
            "enable_cache_tabs": True,  # 默认启用缓存标签功能
            "cached_tabs": [],  # 缓存的非固定标签页
            "last_active_tab_path": "",  # 最近一次激活的标签页路径
            "split_session": {"active": False, "tabs": [], "active_index": 0},  # 右侧分屏组会话（用于重启/崩溃恢复分屏）
            "enable_tortoisegit_buttons": False,  # 默认关闭TortoiseGit按钮
            "preferred_terminal_tool": "cmd",  # 默认终端类型
            "enable_title_shortcuts": True,  # 默认启用标题栏快捷方式区域
            "title_shortcuts": [],  # 标题栏快捷方式（.lnk/.exe/.bat/.cmd/.ps1 路径）
            "enable_mouse_gestures": True,  # 默认启用鼠标手势（右键画线导航）
            # 快捷键配置
            "hotkeys": {
                "new_tab": True,           # Ctrl+T
                "close_tab": True,         # Ctrl+W
                "reopen_tab": True,        # Ctrl+Shift+T
                "switch_tab": True,        # Ctrl+Tab / Ctrl+Shift+Tab
                "search": True,            # Ctrl+F
                "quick_find_current_dir": True,  # Ctrl+G
                "navigate": True,          # Alt+Left/Right
                "go_up": True,             # Alt+Up
                "refresh": True,           # F5
                "add_bookmark": True,      # Ctrl+D
                "copy_filename": True,     # Alt+Z - 复制选中文件名
                "copy_filepath": True,     # Alt+X - 复制文件路径\文件名
                "split_view": True         # F3 - 左右分屏对比
            },
            "language": "zh",              # 界面语言：zh / en
        }
        default_config["ai_chat"] = {
            "enabled": False,
            "api_url": "",       # OpenAI 兼容 API 基础地址，如 http://your-company/v1
            "api_key": "",       # API 密钥（可留空）
            "model": "gpt-3.5-turbo",
            "system_prompt": "",  # 留空则使用内置默认提示词
            "panel_width": 360,  # 面板宽度（像素）
            "panel_visible": False,  # AI面板是否可见
        }
        default_config["performance"] = {
            "content_search_chunk_size": CONTENT_SEARCH_CHUNK_SIZE,
            "content_search_max_bytes_per_file": CONTENT_SEARCH_MAX_BYTES_PER_FILE,
            "content_search_in_memory_threshold": CONTENT_SEARCH_IN_MEMORY_THRESHOLD,
            "search_result_queue_maxsize": SEARCH_RESULT_QUEUE_MAXSIZE,
            "search_result_batch_base": SEARCH_RESULT_BATCH_BASE,
            "search_result_batch_min": SEARCH_RESULT_BATCH_MIN,
            "search_result_batch_max": SEARCH_RESULT_BATCH_MAX,
            "search_metadata_degrade_enabled": SEARCH_METADATA_DEGRADE_ENABLED,
            "search_metadata_degrade_queue_ratio": SEARCH_METADATA_DEGRADE_QUEUE_RATIO,
        }
        
        try:
            # 首先尝试加载主配置文件（使用程序所在目录的绝对路径）
            config_path = get_app_data_path("config.json")
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    # 合并默认配置
                    for key, value in default_config.items():
                        if key not in config:
                            config[key] = value
                    # 确保hotkeys存在所有键
                    if "hotkeys" in config:
                        for key, value in default_config["hotkeys"].items():
                            if key not in config["hotkeys"]:
                                config["hotkeys"][key] = value
                    else:
                        config["hotkeys"] = default_config["hotkeys"]

                    # 确保 ai_chat 存在所有键
                    if "ai_chat" in config and isinstance(config["ai_chat"], dict):
                        for key, value in default_config["ai_chat"].items():
                            if key not in config["ai_chat"]:
                                config["ai_chat"][key] = value
                    else:
                        config["ai_chat"] = default_config["ai_chat"]

                    # 确保 performance 存在所有键
                    if "performance" in config and isinstance(config["performance"], dict):
                        for key, value in default_config["performance"].items():
                            if key not in config["performance"]:
                                config["performance"][key] = value
                    else:
                        config["performance"] = default_config["performance"]

                    apply_runtime_performance_config(config.get("performance"))
                    _set_app_language(config.get("language", "zh"))
                    return config
            else:
                print("No config file found, starting with default config")
                apply_runtime_performance_config(default_config.get("performance"))
                return default_config
        except Exception as e:
            print(f"Failed to load config: {e}")
            apply_runtime_performance_config(default_config.get("performance"))
            return default_config
    
    def save_config(self, immediate=False):
        """保存配置文件（带防抖：500ms无操作后写盘，避免高频config更改频繁I/O）"""
        if immediate:
            self._flush_config_to_disk()
        else:
            if not hasattr(self, '_config_save_timer') or self._config_save_timer is None:
                from PyQt5.QtCore import QTimer
                self._config_save_timer = QTimer(self)
                self._config_save_timer.setSingleShot(True)
                self._config_save_timer.timeout.connect(self._flush_config_to_disk)
            self._config_save_timer.start(500)

    def _flush_config_to_disk(self):
        """实际写入config.json（原子写入：先写临时文件再重命名；内容无变化时跳过写盘）"""
        config_path = get_app_data_path("config.json")
        tmp_path = config_path + ".tmp"
        try:
            new_content = json.dumps(self.config, ensure_ascii=False, indent=2)
            # 若文件已存在且内容相同，则跳过写盘
            if os.path.isfile(config_path):
                try:
                    with open(config_path, "r", encoding="utf-8") as f:
                        existing_content = f.read()
                    if existing_content == new_content:
                        return
                except Exception:
                    pass
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(new_content)
            os.replace(tmp_path, config_path)
        except Exception as e:
            print(f"Failed to save config: {e}")
            try:
                os.remove(tmp_path)
            except Exception:
                pass

    def _collect_cached_tabs(self):
        cached_tabs = []
        seen_paths = set()
        pinned_norm = {
            self._normalize_path_for_compare(p)
            for p in self.config.get("pinned_tabs", []) if p
        }

        if not hasattr(self, 'tab_widget'):
            return cached_tabs

        # 仅收集左侧主标签组的非固定标签（右侧分屏组由 split_session 单独持久化，避免重复恢复）
        cs = getattr(self, 'content_stack', None)
        if cs is None:
            return cached_tabs
        for i in range(cs.count()):
            tab = cs.widget(i)
            if not tab or not hasattr(tab, 'current_path'):
                continue
            if getattr(tab, 'is_pinned', False):
                continue

            current_path = getattr(tab, 'current_path', '')
            if not current_path:
                continue

            norm = self._normalize_path_for_compare(current_path)
            if norm in pinned_norm or norm in seen_paths:
                continue

            seen_paths.add(norm)
            cached_tabs.append({
                'path': current_path,
                'is_shell': current_path.startswith('shell:'),
            })

        return cached_tabs

    def _collect_split_session(self):
        """收集右侧分屏组会话状态用于持久化恢复。

        仅记录右侧组的非固定标签（固定标签统一在左侧恢复）。返回结构：
        {"active": bool, "tabs": [{"path", "is_shell"}], "active_index": int}。"""
        state = {"active": False, "tabs": [], "active_index": 0}
        if not getattr(self, '_split_active', False):
            return state
        scs = getattr(self, 'split_content_stack', None)
        stw = getattr(self, 'split_tab_widget', None)
        if scs is None or stw is None or stw.count() == 0:
            return state
        pinned_norm = {
            self._normalize_path_for_compare(p)
            for p in self.config.get("pinned_tabs", []) if p
        }
        seen = set()
        tabs = []
        for i in range(scs.count()):
            tab = scs.widget(i)
            if not tab or not hasattr(tab, 'current_path'):
                continue
            if getattr(tab, 'is_pinned', False):
                continue
            current_path = getattr(tab, 'current_path', '')
            if not current_path:
                continue
            norm = self._normalize_path_for_compare(current_path)
            if norm in pinned_norm or norm in seen:
                continue
            seen.add(norm)
            tabs.append({
                'path': current_path,
                'is_shell': current_path.startswith('shell:'),
            })
        if not tabs:
            return state
        state["active"] = True
        state["tabs"] = tabs
        idx = stw.currentIndex()
        state["active_index"] = max(0, min(idx, len(tabs) - 1))
        return state

    def _get_last_active_tab_path(self):
        try:
            current_tab = self.get_current_tab_widget()
            current_path = getattr(current_tab, 'current_path', '') if current_tab else ''
            return current_path or ""
        except Exception:
            return ""

    def _schedule_session_snapshot(self, delay_ms=SESSION_SNAPSHOT_DEBOUNCE_MS):
        if not hasattr(self, '_session_snapshot_debounce_timer') or self._session_snapshot_debounce_timer is None:
            return
        # 节流：距上次实际写入不足 SESSION_SNAPSHOT_MIN_INTERVAL_MS 时，
        # 把延迟延长到凑满最小间隔，防止 DirPoll/FileWatcher 高频触发写盘。
        import time
        now_ms = int(time.monotonic() * 1000)
        last_save_ms = getattr(self, '_last_snapshot_save_time_ms', 0)
        elapsed = now_ms - last_save_ms
        if elapsed < SESSION_SNAPSHOT_MIN_INTERVAL_MS:
            effective_delay = max(int(delay_ms), SESSION_SNAPSHOT_MIN_INTERVAL_MS - elapsed + int(delay_ms))
        else:
            effective_delay = int(delay_ms)
        self._session_snapshot_debounce_timer.start(max(0, effective_delay))

    def _get_housekeeping_interval_ms(self):
        try:
            value = int(self.config.get("resource_snapshot_interval_ms", HOUSEKEEPING_INTERVAL_MS))
        except Exception:
            value = HOUSEKEEPING_INTERVAL_MS
        return max(60 * 1000, min(60 * 60 * 1000, value))

    def _append_resource_snapshot_log(self, reason="periodic"):
        if not self.config.get("resource_snapshot_logging", False):
            return
        try:
            from datetime import datetime

            rss_mb = get_process_memory_usage_mb()
            search_dialogs = len(getattr(self, 'search_dialogs', []) or [])
            tabs = self.tab_widget.count() if hasattr(self, 'tab_widget') else 0
            chat_worker_running = 0
            if self.chat_panel is not None:
                worker = getattr(self.chat_panel, 'worker', None)
                if worker and worker.isRunning():
                    chat_worker_running = 1

            thread_count = threading.active_count()
            # 当线程数异常增多时，记录线程名称以诊断泄漏
            thread_detail = ""
            if thread_count > 10:
                try:
                    names = [t.name for t in threading.enumerate()]
                    from collections import Counter
                    name_counts = Counter(names)
                    # 只记录出现超过1次的线程名，或全部（若总数<=20）
                    if thread_count <= 20:
                        thread_detail = " thread_names=[" + ",".join(names) + "]"
                    else:
                        repeated = {k: v for k, v in name_counts.items() if v > 1}
                        if repeated:
                            thread_detail = " thread_repeats=" + str(dict(repeated))
                        else:
                            thread_detail = " thread_names=[" + ",".join(names[:20]) + "...]"
                except Exception:
                    pass

            line = (
                f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                f" reason={reason}"
                f" rss_mb={rss_mb if rss_mb is not None else 'n/a'}"
                f" threads={thread_count}"
                f" tabs={tabs}"
                f" search_dialogs={search_dialogs}"
                f" toasts={len(_active_toasts)}"
                f" chat_worker={chat_worker_running}"
                f" shortcuts_tracked={len(getattr(self, '_last_keys_state', {}) or {})}"
                f"{thread_detail}"
            )
            log_path = get_app_data_path('runtime_health.log')
            try:
                if os.path.exists(log_path) and os.path.getsize(log_path) > MAX_HEALTH_LOG_BYTES:
                    backup = log_path + '.1'
                    try:
                        if os.path.exists(backup):
                            os.remove(backup)
                    except Exception:
                        pass
                    os.replace(log_path, backup)
            except Exception:
                pass
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(line + "\n")
        except Exception as e:
            debug_print(f"[Housekeeping] resource snapshot failed: {e}")

    def _prune_search_dialog_refs(self):
        dialogs = getattr(self, 'search_dialogs', None)
        if dialogs is None:
            return 0
        alive_dialogs = []
        removed = 0
        for dlg in dialogs:
            try:
                dlg.isVisible()
                alive_dialogs.append(dlg)
            except RuntimeError:
                removed += 1
            except Exception:
                alive_dialogs.append(dlg)
        self.search_dialogs = alive_dialogs
        return removed

    def _prune_toast_refs(self):
        global _active_toasts
        alive_toasts = []
        removed = 0
        for toast in list(_active_toasts):
            try:
                if toast.isVisible():
                    alive_toasts.append(toast)
                else:
                    removed += 1
            except RuntimeError:
                removed += 1
            except Exception:
                alive_toasts.append(toast)
        _active_toasts = alive_toasts[:MAX_ACTIVE_TOASTS]
        return removed

    @staticmethod
    def _cleanup_dead_dummy_threads():
        """清除 threading._active 中死亡 QThread 遗留的 _DummyThread 条目。

        PyQt5 的 QThread 子类在运行 Python run() 方法时，会在 threading._active 中
        登记一个 _DummyThread 占位对象。由于 _active 持有强引用，Python 3.9 的
        _DummyThread.__del__ 无法被 GC 触发，导致计数永久增长。
        用 sys._current_frames() 判断底层 OS 线程是否仍存活，对死亡条目执行手动清除。
        """
        import sys
        try:
            live_idents = set(sys._current_frames().keys())
            cleaned = 0
            with threading._active_limbo_lock:
                dead = [
                    ident for ident, t in list(threading._active.items())
                    if isinstance(t, threading._DummyThread) and ident not in live_idents
                ]
                for ident in dead:
                    del threading._active[ident]
                    cleaned += 1
            return cleaned
        except Exception:
            return 0

    def _run_housekeeping(self):
        self._housekeeping_runs += 1
        removed_dialogs = self._prune_search_dialog_refs()
        removed_toasts = self._prune_toast_refs()
        if not self.isActiveWindow():
            self._last_keys_state.clear()

        self._append_resource_snapshot_log(reason="periodic")

        gc_collected = None
        dummy_cleaned = 0
        if self._housekeeping_runs % HOUSEKEEPING_GC_EVERY_N == 0:
            try:
                import gc
                gc_collected = gc.collect()
            except Exception as e:
                debug_print(f"[Housekeeping] gc.collect failed: {e}")
            dummy_cleaned = self._cleanup_dead_dummy_threads()

        if removed_dialogs or removed_toasts or gc_collected is not None or dummy_cleaned:
            debug_print(
                f"[Housekeeping] dialogs={removed_dialogs} toasts={removed_toasts}"
                f" gc={gc_collected} dummy_threads_cleaned={dummy_cleaned}"
            )

    def apply_resource_usage_config(self):
        """根据配置开启/关闭状态栏右侧 CPU/内存占用显示。"""
        timer = getattr(self, '_resource_usage_timer', None)
        if timer is None:
            return
        enabled = self.config.get("show_resource_usage_in_statusbar", False)
        if enabled:
            if not timer.isActive():
                timer.start(2000)  # 每2秒刷新一次，足够直观又低开销
            self._update_resource_usage_display()
        else:
            if timer.isActive():
                timer.stop()
            # 隐藏所有标签上的资源标签
            for i in range(self.tab_widget.count()):
                tab = self.get_tab_widget(i)
                lbl = getattr(tab, 'resource_label', None) if tab else None
                if lbl:
                    lbl.hide()
                    lbl.setText("")

    def _update_resource_usage_display(self):
        """显示整机 CPU/内存占用到当前活动标签的资源标签上，占用过高时变色预警。"""
        if not self.config.get("show_resource_usage_in_statusbar", False):
            return

        def _color(pct):
            if pct >= RESOURCE_CRIT_PERCENT:
                return "#d32f2f"  # 红：危急
            if pct >= RESOURCE_WARN_PERCENT:
                return "#e67700"  # 橙：偏高
            return "#666"          # 常态

        cpu = get_system_cpu_percent()
        mem = get_system_memory_status()
        cpu_html = (f"<span style='color:{_color(cpu)}'>CPU {cpu:.0f}%</span>"
                    if cpu is not None else "<span style='color:#666'>CPU --</span>")
        if mem is not None:
            used_mb, total_mb, pct = mem
            mem_html = (f"<span style='color:{_color(pct)}'>{tr('内存')} "
                        f"{used_mb/1024:.1f}/{total_mb/1024:.1f} GB ({pct}%)</span>")
        else:
            mem_html = ""
        text = f"{cpu_html}&nbsp;&nbsp;&nbsp;{mem_html}".strip()
        tab = self.get_current_tab_widget()
        lbl = getattr(tab, 'resource_label', None) if tab else None
        if lbl:
            lbl.setText(text)
            lbl.show()

    def _restore_last_active_tab(self):
        last_active_path = self.config.get("last_active_tab_path", "")
        if not last_active_path:
            return False

        tab_index = self.find_tab_index_by_path(last_active_path)
        if tab_index < 0:
            return False

        self.tab_widget.setCurrentIndex(tab_index)
        return True

    def _restore_split_session(self):
        """根据持久化的 split_session 恢复右侧分屏组（启动/崩溃恢复时调用）。

        左侧组至少要有一个标签，右侧分屏才独立成立。返回 True 表示已恢复分屏。"""
        if not self.config.get("enable_cache_tabs", True):
            return False
        state = self.config.get("split_session", {}) or {}
        if not state.get("active"):
            return False
        tabs = state.get("tabs", []) or []
        if not tabs:
            return False
        # 左侧主组必须至少保留一个标签，否则不进入分屏
        if not hasattr(self, 'tab_widget') or self.tab_widget.count() == 0:
            return False
        # 显示右侧分屏 UI 布局，然后逐个在右侧组创建标签
        self._activate_split_layout()
        added = 0
        for tab_info in tabs:
            path = tab_info.get('path', '')
            if not path:
                continue
            try:
                self.add_new_tab(
                    path,
                    is_shell=tab_info.get('is_shell', False),
                    target_tabwidget=self.split_tab_widget,
                    activate=False,
                )
                added += 1
            except Exception as e:
                debug_print(f"[App] 恢复右侧分屏标签失败: {path} -> {e}")
        if added == 0:
            # 一个都没成功 → 收起分屏，回到单组
            self._teardown_split_group()
            return False
        active_index = state.get("active_index", 0)
        if 0 <= active_index < self.split_tab_widget.count():
            self.split_tab_widget.setCurrentIndex(active_index)
        return True

    def save_session_snapshot(self, immediate=False):
        if not hasattr(self, 'config') or not hasattr(self, 'tab_widget'):
            return

        try:
            import time
            self._last_snapshot_save_time_ms = int(time.monotonic() * 1000)
            cached_tabs = []
            split_session = {"active": False, "tabs": [], "active_index": 0}
            if self.config.get("enable_cache_tabs", True):
                cached_tabs = self._collect_cached_tabs()
                split_session = self._collect_split_session()

            new_active = self._get_last_active_tab_path()

            # 快照内容无变化时跳过写盘，避免无意义IO和日志刷屏
            # 使用 JSON 签名比较，避免 list/dict 对象引用差异导致误判
            import json as _json
            _sig = (_json.dumps(cached_tabs, sort_keys=True, ensure_ascii=False)
                    + '|' + new_active
                    + '|' + _json.dumps(split_session, sort_keys=True, ensure_ascii=False))
            # immediate=True 用于关闭/初始化等关键时机，必须确保落盘，因此绕过签名去重
            if not immediate and _sig == getattr(self, '_last_snapshot_sig', ''):
                return
            self._last_snapshot_sig = _sig

            self.config["cached_tabs"] = cached_tabs
            self.config["last_active_tab_path"] = new_active
            self.config["split_session"] = split_session
            self.save_config(immediate=immediate)
            debug_print(
                f"[App] 会话快照已更新: tabs={len(cached_tabs)}, active='{new_active}', "
                f"split={'on' if split_session.get('active') else 'off'}({len(split_session.get('tabs', []))})"
            )
        except Exception as e:
            print(f"Error saving session snapshot: {e}")

    def ensure_default_bookmarks(self):
        bm = self.bookmark_manager
        tree = bm.get_tree()
        # 只在bookmark_bar存在且children为空时添加
        if 'bookmark_bar' not in tree:
            # 兼容空书签文件，自动创建bookmark_bar
            import time
            bar_id = str(int(time.time() * 1000000))
            tree['bookmark_bar'] = {
                "date_added": bar_id,
                "id": bar_id,
                "name": tr("书签栏"),
                "type": "folder",
                "children": []
            }
        bar = tree['bookmark_bar']
        # 去重：同一 shell 特殊文件夹的中英文重复书签，按当前语言只保留一种
        if self._dedup_shell_bookmarks(bar):
            bm.save_bookmarks()
        if 'children' not in bar or not bar['children']:
            # 添加常用项目
            import time
            now = int(time.time() * 1000000)
            def make_bm(name, url, icon):
                nonlocal now
                now += 1
                return {
                    "date_added": str(now),
                    "id": str(now),
                    "name": f"{icon} {name}",
                    "type": "url",
                    "url": url
                }
            bar['children'] = [
                make_bm(tr("此电脑"), "shell:MyComputerFolder", "🖥️"),
                make_bm(tr("桌面"), "shell:Desktop", "🗔"),
                make_bm(tr("回收站"), "shell:RecycleBinFolder", "🗑️"),
            ]
            bm.save_bookmarks()

    def _dedup_shell_bookmarks(self, bar):
        """去除书签栏中指向同一 shell 特殊文件夹的中英文重复项。

        对已知的 shell 特殊文件夹（此电脑/回收站/桌面），若存在多条指向同一
        URL 的书签，则按当前语言设置只保留一种命名（语言匹配优先，否则保留首项）。
        返回 True 表示发生了修改。
        """
        children = bar.get('children')
        if not isinstance(children, list):
            return False
        # shell 特殊文件夹 URL（小写）-> 中文基础名（英文名由 tr() 推导）
        special = {
            'shell:mycomputerfolder': '此电脑',
            'shell:recyclebinfolder': '回收站',
            'shell:desktop': '桌面',
        }
        prefer_en = (_app_language == 'en')

        def _matches_lang(node, zh):
            name = str(node.get('name', ''))
            en = tr(zh)
            if prefer_en and en != zh:
                return en in name
            return zh in name

        # 第一遍：为每个重复 URL 选出要保留的书签
        keepers = {}        # key -> node
        keeper_matched = {}  # key -> bool（保留项是否语言匹配）
        for node in children:
            if not (isinstance(node, dict) and node.get('type') == 'url'):
                continue
            key = str(node.get('url', '')).lower()
            zh = special.get(key)
            if zh is None:
                continue
            matched = _matches_lang(node, zh)
            if key not in keepers:
                keepers[key] = node
                keeper_matched[key] = matched
            elif matched and not keeper_matched[key]:
                keepers[key] = node
                keeper_matched[key] = True

        # 第二遍：重建列表，每个特殊 URL 只保留选中的那一条
        new_children = []
        emitted = set()
        changed = False
        for node in children:
            if isinstance(node, dict) and node.get('type') == 'url':
                key = str(node.get('url', '')).lower()
                if key in special:
                    if key in emitted:
                        changed = True
                        continue
                    emitted.add(key)
                    if keepers[key] is not node:
                        changed = True
                    new_children.append(keepers[key])
                    continue
            new_children.append(node)
        if changed:
            bar['children'] = new_children
        return changed



    def init_ui(self):
        # 获取DPI缩放因子
        from PyQt5.QtWidgets import QApplication
        screen = QApplication.primaryScreen()
        dpi = screen.logicalDotsPerInch()
        self.dpi_scale = dpi / 96.0
        debug_print(f"[MainWindow] DPI scale factor: {self.dpi_scale:.2f}")
        
        # 设置窗口最小尺寸，允许窗口缩小到很小
        min_width = int(400 * self.dpi_scale)
        min_height = int(300 * self.dpi_scale)
        self.setMinimumSize(min_width, min_height)
        
        
        # 使用系统原生标题栏（彻底修复无边框窗口最小化恢复后 backing store 停摆问题）
        self.setWindowFlags(Qt.Window)

        # 填充窗口背景，避免边框与内容之间出现半透明缝隙
        self.setAutoFillBackground(True)
        from PyQt5.QtGui import QPalette, QColor
        pal = self.palette()
        pal.setColor(QPalette.Window, QColor(255, 255, 255))
        self.setPalette(pal)
        # 再次用样式表确保非客户区也以白色填充
        self.setStyleSheet("QMainWindow { background: white; }")
        
        
        # 创建主容器，无边距，纯白填充
        main_container = QWidget()
        main_container.setAutoFillBackground(True)
        pal_container = main_container.palette()
        pal_container.setColor(QPalette.Window, QColor(255, 255, 255))
        main_container.setPalette(pal_container)
        main_container.setStyleSheet("QWidget { background: white; margin: 0px; padding: 0px; border: none; }")
        container_layout = QVBoxLayout(main_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)
        
        # 保存主容器引用，用于应用阴影效果
        self._main_container = main_container
        
        # 创建内容布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 创建自定义标题栏
        self.create_custom_titlebar(main_layout)

        # 创建标签页控件（支持拖放）
        self.tab_widget = DragDropTabWidget(self)
        self.tab_widget.setTabsClosable(False)  # 禁用默认关闭按钮，使用自定义悬停关闭按钮
        self.tab_widget.currentChanged.connect(self.on_tab_changed)

        # 使用自定义 TabBar 支持双击空白区域打开新标签页
        custom_tabbar = CustomTabBar()
        custom_tabbar.main_window = self
        custom_tabbar.owner_tabwidget = self.tab_widget
        self.tab_widget.setTabBar(custom_tabbar)

        # 设置选中标签页背景色为淡黄色
        tabbar = self.tab_widget.tabBar()
        tabbar.setAcceptDrops(True)
        
        # 根据DPI缩放标签栏尺寸
        tab_height = int(24 * self.dpi_scale)
        tab_width = int(120 * self.dpi_scale)
        tab_padding_v = int(2 * self.dpi_scale)
        tab_padding_h = int(4 * self.dpi_scale)
        tab_radius = int(6 * self.dpi_scale)
        tab_font_size = int(12 * self.dpi_scale)
        tab_margin = int(2 * self.dpi_scale)
        
        tabbar.setStyleSheet(f"""
            QTabBar::tab {{
                background: #f5f5f5;
                border: 1px solid #d0d0d0;
                border-bottom: none;
                border-top-left-radius: {tab_radius}px;
                border-top-right-radius: {tab_radius}px;
                padding: {tab_padding_v}px {tab_padding_h}px;
                height: {tab_height}px;
                width: {tab_width}px;
                min-width: {tab_width}px;
                max-width: {tab_width}px;
                font-family: 'Microsoft YaHei UI', 'Segoe UI', Arial, sans-serif;
                font-size: {tab_font_size}px;
                margin-top: {tab_margin}px;
                margin-right: 0px;
                text-align: center;
                color: #505050;
            }}
            QTabBar::tab:hover:!selected {{
                background: #e8e8e8;
                border-color: #c0c0c0;
            }}
            QTabBar::tab:selected {{
                background: #FFF9CC;
                border: 1px solid #c0c0c0;
                border-bottom: none;
                margin-top: 0px;
                padding-top: {tab_padding_v + 1}px;
                color: #000000;
            }}
            QTabBar::tab:!selected {{
                font-weight: normal;
                margin-top: {tab_margin + 1}px;
            }}
        """)
        # 设置标签文本省略模式 - 左边省略，保留右侧文件/文件夹名称
        tabbar.setElideMode(Qt.ElideLeft)

        # 创建标签栏容器（只显示标签和按钮，不显示内容）
        tab_bar_container = QWidget()
        tab_bar_height = int(32 * getattr(self, 'dpi_scale', 1.0))
        tab_bar_container.setFixedHeight(tab_bar_height)  # 固定高度，只显示标签栏
        tab_bar_container.setStyleSheet("background-color: #f3f3f3;")
        tab_bar_layout = QHBoxLayout(tab_bar_container)
        tab_bar_layout.setContentsMargins(0, 0, 0, 0)
        tab_bar_layout.setSpacing(0)
        
        # 将 tab_widget 添加到标签栏容器（只显示标签栏部分）
        self.tab_widget.setMaximumHeight(tab_bar_height)  # 限制最大高度
        tab_bar_layout.addWidget(self.tab_widget)

        # 右侧分屏标签组（默认隐藏，F3 时显示）：与左侧标签栏并排于书签栏上方，
        # 拥有自己的标签栏，支持双击新建、关闭、拖拽等原生标签功能。
        self.split_tab_widget = DragDropTabWidget(self)
        self.split_tab_widget.setTabsClosable(False)
        self.split_tab_widget.currentChanged.connect(self._on_split_tab_changed)
        split_tabbar = CustomTabBar()
        split_tabbar.main_window = self
        split_tabbar.owner_tabwidget = self.split_tab_widget
        self.split_tab_widget.setTabBar(split_tabbar)
        split_tabbar.setAcceptDrops(True)
        split_tabbar.setStyleSheet(tabbar.styleSheet())
        split_tabbar.setElideMode(Qt.ElideLeft)
        # 右侧分屏标签栏也支持右键菜单（固定/取消固定/书签等），与左侧一致
        split_tabbar.setContextMenuPolicy(Qt.CustomContextMenu)
        split_tabbar.customContextMenuRequested.connect(
            lambda pos: self.tab_context_menu(pos, self.split_tab_widget))
        self.split_tab_widget.setMaximumHeight(tab_bar_height)
        self.split_tab_widget.setVisible(False)
        self._split_active = False
        tab_bar_layout.addWidget(self.split_tab_widget)
        
        # 将标签栏容器添加到主布局
        main_layout.addWidget(tab_bar_container)
        
        # 右键标签页支持固定/取消固定
        tabbar.setContextMenuPolicy(Qt.CustomContextMenu)
        tabbar.customContextMenuRequested.connect(
            lambda pos: self.tab_context_menu(pos, self.tab_widget))

        # 书签栏（使用自定义菜单栏）
        self.menu_bar = CustomMenuBar(self)
        menu_bar_height = int(28 * getattr(self, 'dpi_scale', 1.0))
        self.menu_bar.setFixedHeight(menu_bar_height)  # 设置菜单栏高度
        # 设置菜单栏的大小策略，允许它被压缩
        self.menu_bar.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self.menu_bar.setStyleSheet("""
            QMenuBar {
                background-color: #f3f3f3;
                border-top: 1px solid #d0d0d0;
                border-bottom: 1px solid #e0e0e0;
                padding: 2px;
            }
            QMenuBar::item {
                padding: 4px 10px;
                background: transparent;
                border-radius: 4px;
                min-width: 0px;
                color: #303030;
            }
            QMenuBar::item:selected {
                background: #e5e5e5;
                color: #000000;
            }
            QMenuBar::item:pressed {
                background: #d5d5d5;
                color: #000000;
            }
            QMenu {
                background-color: #ffffff;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                padding: 4px;
                color: #000000;
            }
            QMenu::item {
                padding: 6px 24px 6px 12px;
                background: transparent;
                border-radius: 4px;
                color: #303030;
                margin: 2px 4px;
            }
            QMenu::item:selected {
                background: #e3f2fd;
                color: #000000;
            }
            QMenu::item:pressed {
                background: #bbdefb;
                color: #000000;
            }
            QMenu::separator {
                height: 1px;
                background: #e5e5e5;
                margin: 4px 8px;
            }
        """)
        self.populate_bookmark_bar_menu()
        # 将菜单栏添加到主布局
        main_layout.addWidget(self.menu_bar)

        # 主分割器，左树右标签
        self.splitter = ResizableSplitter()
        self.splitter.setOrientation(Qt.Horizontal)
        # 加宽分割条，便于鼠标识别与抓取（自定义手柄会显示左右调整光标与抓取条）
        self.splitter.setHandleWidth(8)
        # content_stack 占据剩余全部空间
        self.splitter.setStretchFactor(0, 1)

        # 右侧标签页内容区域（使用 StackedWidget 独立显示，不依赖 tab_widget）
        from PyQt5.QtWidgets import QStackedWidget
        self.content_stack = QStackedWidget()
        self.content_stack.setStyleSheet("background: white;")
        self.content_stack.setAutoFillBackground(True)
        # 允许右侧内容在窗口缩小时被压缩，避免阻止左侧目录树向左拖动
        self.content_stack.setMinimumWidth(0)
        
        self.splitter.addWidget(self.content_stack)
        
        self.splitter.setCollapsible(0, False)  # content_stack 不允许折叠

        # 右侧分屏内容栈（默认不加入分割器，F3 时插入到索引 1）
        self.split_content_stack = QStackedWidget()
        self.split_content_stack.setStyleSheet("background: white;")
        self.split_content_stack.setAutoFillBackground(True)
        self.split_content_stack.setMinimumWidth(0)
        self.split_content_stack.setVisible(False)
        # 分割器拖动时让右侧标签栏宽度跟随右侧内容宽度对齐
        self.splitter.splitterMoved.connect(lambda *a: self._sync_split_tabbar_width())

        # 右侧 AI 聊天面板（延迟加载，点击 🤖 按钮后才创建）
        self.chat_panel = None  # 延迟创建
        # 先设置 splitter 只有 content_stack
        right_width = int(1200 * self.dpi_scale)
        self.splitter.setSizes([right_width, 0])
    
        # 将分割器添加到主容器
        main_layout.addWidget(self.splitter)
        
        # 将内容布局添加到主容器
        container_layout.addLayout(main_layout)
        
        # 设置主容器为中心部件
        self.setCentralWidget(main_container)

        # 性能优化：延迟加载固定标签页（移到 _delayed_initialization）
        # 先检查是否有固定标签或缓存标签，如果没有才添加默认标签页
        has_content = bool(
            self.config.get("pinned_tabs", []) or 
            self.config.get("cached_tabs", [])
        )
        if not has_content:
            # 没有固定标签也没有缓存标签，才添加默认的主目录标签
            self.add_new_tab(QDir.homePath())
        
        # 连接信号
        self.open_path_signal.connect(self.handle_open_path_from_instance)
        
        # 性能优化：单实例服务器和Explorer监听移到延迟初始化
    
    def resizeEvent(self, event):
        """窗口大小改变事件"""
        super().resizeEvent(event)

    def _bring_to_front(self):
        """强制将窗口置顶并获取焦点（兼容 Windows 防偷焦机制）"""
        # 记录当前最大化状态：Win32 激活调用有时会意外取消最大化
        was_maximized = self.isMaximized()
        # 最小化时先恢复
        if self.isMinimized():
            self.showNormal()
        try:
            import ctypes
            user32 = ctypes.windll.user32
            hwnd = int(self.winId())
            # AttachThreadInput 技巧：临时附加到前台线程，绕过 Windows 偷焦保护
            fg_hwnd = user32.GetForegroundWindow()
            fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, None)
            my_tid = ctypes.windll.kernel32.GetCurrentThreadId()
            attached = False
            if fg_tid and fg_tid != my_tid:
                user32.AttachThreadInput(fg_tid, my_tid, True)
                attached = True
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            if attached:
                user32.AttachThreadInput(fg_tid, my_tid, False)
        except Exception:
            pass
        # Qt 层兜底
        self.activateWindow()
        self.raise_()
        # Win32 激活调用有时会使最大化窗口还原；50ms 后检查并恢复
        if was_maximized:
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(50, lambda: self.showMaximized() if not self.isMaximized() else None)

    def handle_open_path_from_instance(self, path):
        """处理从其他实例接收到的路径（在主线程中）"""
        existing_index = self.find_tab_index_by_path(path)
        if existing_index >= 0:
            print(f"[MainWindow] Path already open, focus tab: {path}")
            self.tab_widget.setCurrentIndex(existing_index)
        else:
            print(f"[MainWindow] Opening path in new tab: {path}")
            self.add_new_tab(path)
        # 强制置顶窗口（使用 Win32 API 绕过 Windows 偷焦保护）
        self._bring_to_front()
    
    def _delayed_initialization(self):
        """延迟初始化非关键功能（性能优化）"""
        debug_print("[Performance] Starting delayed initialization...")
        # 进入恢复状态：在标题上显示提示
        self.is_restoring = True
        # 初次更新标题（无路径）
        self._update_window_title()
        
        debug_print(f"[App] 启动时标签页数: {self.tab_widget.count()}")
        
        # 延迟加载固定标签页
        try:
            has_pinned = self.load_pinned_tabs()
            debug_print(f"[App] 加载固定标签后标签页数: {self.tab_widget.count()}")
            if has_pinned:
                debug_print(tr("[App] 已加载固定标签页"))
        except Exception as e:
            debug_print(f"[Performance] Failed to load pinned tabs: {e}")
        
        # 恢复缓存的标签页
        try:
            if self.config.get("enable_cache_tabs", True):
                cached_tabs = self.config.get("cached_tabs", [])
                debug_print(f"[App] 待恢复的缓存标签页数: {len(cached_tabs)}")
                if cached_tabs:
                    pinned_norm = {
                        self._normalize_path_for_compare(p)
                        for p in self.config.get("pinned_tabs", []) if p
                    }
                    debug_print(f"[App] 恢复 {len(cached_tabs)} 个缓存标签页")
                    for tab_info in cached_tabs:
                        path = tab_info.get('path', '')
                        if path:
                            norm = self._normalize_path_for_compare(path)
                            if norm in pinned_norm:
                                debug_print(tr("[App] 跳过缓存标签（已固定）: {}").format(path))
                                continue
                            if self.is_path_open(path):
                                debug_print(tr("[App] 跳过缓存标签（已打开）: {}").format(path))
                                continue
                            # 懒加载：恢复的缓存标签不逐个激活，后台标签首次可见时才导航
                            self.add_new_tab(path, activate=False)
                    debug_print(f"[App] 恢复缓存标签后标签页数: {self.tab_widget.count()}")
                else:
                    # 没有缓存标签且没有固定标签，现在添加默认标签
                    if self.tab_widget.count() == 0:
                        debug_print(tr("[App] 没有缓存和固定标签，添加默认主目录标签"))
                        self.add_new_tab(QDir.homePath())
        except Exception as e:
            debug_print(f"[Performance] Failed to restore cached tabs: {e}")
            # 如果恢复失败且当前没有标签，添加默认标签
            if self.tab_widget.count() == 0:
                self.add_new_tab(QDir.homePath())

        # 恢复右侧分屏组（若上次退出/崩溃时处于分屏状态）
        try:
            if self._restore_split_session():
                debug_print(tr("[App] 已恢复右侧分屏组"))
        except Exception as e:
            debug_print(f"[Performance] Failed to restore split session: {e}")

        # 延迟启动实例服务器
        try:
            self.start_instance_server()
        except Exception as e:
            debug_print(f"[Performance] Failed to start instance server: {e}")
        
        # 延迟启动Explorer监听（如果启用）
        try:
            if self.config.get("enable_explorer_monitor", True):
                from PyQt5.QtCore import QTimer
                # 再延迟500ms启动Explorer监听，避免影响启动速度
                QTimer.singleShot(500, self.start_explorer_monitor)
        except Exception as e:
            debug_print(f"[Performance] Failed to start explorer monitor: {e}")

        restored_active = self._restore_last_active_tab()
        if restored_active:
            debug_print(tr("[App] 已恢复上次激活的标签页"))
        # 兜底：懒加载下所有恢复标签均未激活时，_restore_last_active_tab 可能未选中任何标签，
        # 导致左侧当前标签仍处于延迟态（界面空白）。这里强制激活一次左侧当前标签，
        # 触发其 showEvent → 首次导航，保证启动后左侧有一个已加载的可见标签。
        try:
            if self.tab_widget.count() > 0:
                cur = self.tab_widget.currentIndex()
                if cur < 0:
                    cur = 0
                    self.tab_widget.setCurrentIndex(0)
                cur_tab = self.get_tab_widget(cur)
                if cur_tab is not None and getattr(cur_tab, '_deferred_nav', None) is not None:
                    # 已是当前项但 showEvent 可能未触发（内容栈未切换）：显式同步并激活
                    self.on_tab_changed(cur)
                    if cur_tab.isVisible():
                        # 直接消费延迟导航，避免依赖 showEvent 时序
                        deferred = cur_tab._deferred_nav
                        cur_tab._deferred_nav = None
                        p, ish = deferred
                        cur_tab.navigate_to(p, is_shell=ish)
        except Exception as e:
            debug_print(f"[App] 兜底激活当前标签失败: {e}")
        self.save_session_snapshot(immediate=True)
        
        # 恢复 AI 聊天面板的显示状态（如果配置为显示）
        try:
            chat_config = self.config.get("ai_chat", {})
            panel_visible = chat_config.get("panel_visible", False)
            if panel_visible:
                # 用户上次关闭时 AI 面板是可见的，现在创建并显示它
                self._ensure_chat_panel_created()
                self.chat_panel.setVisible(True)
                if hasattr(self, 'ai_chat_btn'):
                    self.ai_chat_btn.setChecked(True)
                # 设置面板宽度
                ai_panel_width = int(chat_config.get("panel_width", 360) * self.dpi_scale)
                sizes = self.splitter.sizes()
                total = sum(sizes)
                self.splitter.setSizes([total - ai_panel_width, ai_panel_width])
        except Exception as e:
            debug_print(f"[Performance] Failed to restore chat panel: {e}")
        
        # 恢复完成：取消恢复状态并更新标题
        self.is_restoring = False
        # 使用当前标签路径刷新标题
        try:
            current_tab = self.get_current_tab_widget()
            current_path = getattr(current_tab, 'current_path', None)
        except Exception:
            current_path = None
        self._update_window_title(current_path)
        debug_print("[Performance] Delayed initialization completed")
    
    def start_instance_server(self):
        """启动本地服务器监听其他实例的请求"""
        if self.server_thread and self.server_thread.is_alive():
            debug_print("[Server] Instance server already running")
            return

        def server_thread():
            try:
                server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                server.bind(('127.0.0.1', 58923))  # 使用固定端口
                server.listen(5)
                server.settimeout(1.0)  # 设置超时，使线程可以退出
                self.server_socket = server
                debug_print("[Server] Instance server started on port 58923")
                
                while getattr(self, 'server_running', True):
                    try:
                        conn, addr = server.accept()
                        data = conn.recv(4096).decode('utf-8')
                        conn.close()
                        
                        if data:
                            debug_print(f"[Server] Received path: {data}")
                            # 使用信号在主线程中打开新标签页
                            self.open_path_signal.emit(data)
                    except socket.timeout:
                        continue
                    except Exception as e:
                        debug_print(f"[Server] Connection error: {e}")
                        continue
            except Exception as e:
                debug_print(f"[Server] Failed to start server: {e}")
            finally:
                self.server_socket = None
                self.server_running = False
        
        self.server_running = True
        server_thread_obj = threading.Thread(target=server_thread, daemon=True)
        self.server_thread = server_thread_obj
        server_thread_obj.start()
        # 服务器在 daemon 线程中独立 bind/listen，主线程无需等待其就绪；
        # 移除原先的 time.sleep(0.2)，避免在延迟初始化阶段无谓阻塞 UI 线程约 200ms。
    
    def start_explorer_monitor(self):
        """启动Explorer窗口监听（优化版）"""
        # 检查配置是否启用
        if not self.config.get("enable_explorer_monitor", True):
            debug_print("[Explorer Monitor] Monitoring disabled in config")
            return
        
        if not HAS_PYWIN:
            debug_print("[Explorer Monitor] Windows API not available, monitoring disabled")
            return
        
        # 获取监听间隔配置（默认2秒，更轻量）
        self.monitor_interval = self.config.get("explorer_monitor_interval", 2.0)
        debug_print(f"[Explorer Monitor] Will start monitoring in 3 seconds (interval: {self.monitor_interval}s)...")
        self.explorer_monitoring = False
        self.known_explorer_windows = set()  # 记录已知的Explorer窗口
        self.last_check_time = 0  # 上次检查时间
        
        # 延迟启动监听，确保主窗口完全初始化
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, self._start_monitor_thread)
    
    def _start_monitor_thread(self):
        """实际启动监听线程（延迟调用）"""
        if self.monitor_thread and self.monitor_thread.is_alive():
            debug_print("[Explorer Monitor] Monitor thread already running")
            return

        try:
            self.monitor_our_window = int(self.winId())  # 记录我们自己的窗口句柄
            self.explorer_monitoring = True
            debug_print("[Explorer Monitor] Starting Explorer window monitoring...")
            
            # 启动监听线程
            monitor_thread = threading.Thread(target=self._explorer_monitor_loop, daemon=True)
            self.monitor_thread = monitor_thread
            monitor_thread.start()
        except Exception as e:
            debug_print(f"[Explorer Monitor] Failed to start: {e}")
    
    def stop_explorer_monitor(self):
        """停止Explorer窗口监听"""
        self.explorer_monitoring = False
        self.known_explorer_windows.clear()
        debug_print("[Explorer Monitor] Stopped")
    
    def _explorer_monitor_loop(self):
        """Explorer窗口监听循环（优化版 - 降低CPU占用）"""
        pythoncom_initialized = False
        try:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                pythoncom_initialized = True
            except Exception as e:
                debug_print(f"[Explorer Monitor] COM init failed: {e}")

            # 首先记录所有已存在的Explorer窗口
            def enum_windows_callback(hwnd, _):
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    # CabinetWClass: 标准Explorer窗口
                    # ExploreWClass: 另一种Explorer窗口类型（如通过"打开文件夹"打开的）
                    if class_name in ("CabinetWClass", "ExploreWClass"):
                        if win32gui.IsWindowVisible(hwnd):
                            self.known_explorer_windows.add(hwnd)
                except Exception:
                    pass
                return True
            
            win32gui.EnumWindows(enum_windows_callback, None)
            debug_print(f"[Explorer Monitor] Found {len(self.known_explorer_windows)} existing Explorer windows")
            debug_print(f"[Explorer Monitor] Monitor interval: {self.monitor_interval}s (optimized for low CPU usage)")
            
            # 定期检查新的Explorer窗口（优化：使用更长的间隔）
            while self.explorer_monitoring:
                time.sleep(self.monitor_interval)  # 默认2秒检查一次（降低CPU占用）
                
                current_time = time.time()
                # 防抖：如果距离上次检查太近，跳过
                if current_time - self.last_check_time < self.monitor_interval * 0.8:
                    continue
                
                self.last_check_time = current_time
                current_explorer_windows = set()
                
                def check_windows_callback(hwnd, _):
                    try:
                        class_name = win32gui.GetClassName(hwnd)
                        # CabinetWClass: 标准Explorer窗口
                        # ExploreWClass: 另一种Explorer窗口类型
                        if class_name in ("CabinetWClass", "ExploreWClass"):
                            if win32gui.IsWindowVisible(hwnd):
                                current_explorer_windows.add(hwnd)
                    except Exception:
                        pass
                    return True
                
                win32gui.EnumWindows(check_windows_callback, None)
                
                # 找出新增的窗口
                new_windows = current_explorer_windows - self.known_explorer_windows
                
                # 优化：如果没有新窗口，直接跳过处理
                if not new_windows:
                    self.known_explorer_windows = current_explorer_windows
                    continue
                
                debug_print(f"[Explorer Monitor] Detected {len(new_windows)} new Explorer window(s)")
                
                for hwnd in new_windows:
                    # 检查是否是我们自己的窗口（避免误捕获嵌入的Explorer控件）
                    try:
                        # 获取窗口标题
                        title = win32gui.GetWindowText(hwnd)
                        
                        debug_print(f"[Explorer Monitor] Checking window: {hwnd} - {title}")
                        
                        # 检查窗口标题是否为控制面板或其子项
                        if (title in [tr('控制面板'), 'Control Panel'] or 
                            title.startswith(tr('控制面板\\')) or title.startswith(tr('控制面板 - ')) or 
                            title.startswith('Control Panel\\') or title.startswith('Control Panel - ') or
                            tr('\\控制面板\\') in title):
                            debug_print(f"[Explorer Monitor] Control Panel detected by title, keeping original window")
                            continue
                        
                        # 只排除明确是我们应用的主窗口，不要误排除路径中包含TabEx的Explorer窗口
                        # 检查是否以"TabExplorer"开头（软件主窗口）或者窗口句柄是我们的主窗口
                        if title.startswith("TabExplorer"):
                            debug_print(f"[Explorer Monitor] Skipping our main window: {title}")
                            continue
                        
                        # 获取窗口的父窗口，如果父窗口是我们的应用，则跳过
                        try:
                            parent = win32gui.GetParent(hwnd)
                            if parent == self.monitor_our_window:
                                debug_print(f"[Explorer Monitor] Skipping child window")
                                continue
                        except Exception:
                            pass
                        
                        debug_print(f"[Explorer Monitor] New Explorer window detected: {hwnd} - {title}")
                        
                        # 尝试获取Explorer窗口的当前路径
                        path = self._get_explorer_path(hwnd)
                        
                        if path:
                            debug_print(f"[Explorer Monitor] ✓ Path: {path}")
                            
                            # 控制面板及其子目录直接在原窗口打开，不拦截
                            if self._is_control_panel_path_for_monitor(path):
                                debug_print(f"[Explorer Monitor] Control Panel detected, keeping original window")
                                # 不关闭原窗口，让控制面板在原生Explorer中打开
                                continue
                            
                            # 非控制面板路径，发送信号并关闭原窗口
                            self.open_path_signal.emit(path)
                            # 优化：减少等待时间（从500ms到200ms）
                            time.sleep(0.2)
                            # 关闭原Explorer窗口
                            try:
                                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                                debug_print(f"[Explorer Monitor] ✓ Closed original Explorer (hwnd={hwnd})")
                            except Exception as e:
                                debug_print(f"[Explorer Monitor] ✗ Failed to close: {e}")
                        else:
                            debug_print(f"[Explorer Monitor] ✗ Could not get path from {hwnd}")
                    
                    except Exception as e:
                        debug_print(f"[Explorer Monitor] Error processing window {hwnd}: {e}")
                
                # 更新已知窗口列表
                self.known_explorer_windows = current_explorer_windows
                
        except Exception as e:
            debug_print(f"[Explorer Monitor] Monitor loop error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.explorer_monitoring = False
            self.known_explorer_windows.clear()
            self.monitor_thread = None
            if pythoncom_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    
    def _get_explorer_path(self, hwnd):
        """通过COM接口获取Explorer窗口的当前路径"""
        try:
            # 使用Shell.Application COM对象
            import win32com.client
            
            # 多次尝试获取路径（有时窗口刚打开时COM对象还没准备好）
            for attempt in range(3):
                try:
                    shell = win32com.client.Dispatch("Shell.Application")
                    
                    # 遍历所有打开的Explorer窗口
                    for window in shell.Windows():
                        try:
                            # 获取窗口句柄
                            window_hwnd = window.HWND
                            
                            if window_hwnd == hwnd:
                                # 先尝试获取 LocationName，用于识别控制面板
                                location_name = None
                                try:
                                    location_name = window.LocationName
                                    debug_print(f"[Explorer Monitor] LocationName: {location_name}")
                                except Exception:
                                    pass
                                
                                # 获取当前路径
                                location = window.LocationURL
                                
                                debug_print(f"[Explorer Monitor] LocationURL: {location}")
                                
                                # 检查是否为控制面板
                                if location_name and location_name in [tr('控制面板'), 'Control Panel']:
                                    debug_print(f"[Explorer Monitor] Control Panel detected by LocationName")
                                    return 'shell:ControlPanelFolder'
                                
                                if location:
                                    # 转换file:///格式的URL为本地路径
                                    if location.startswith('file:///'):
                                        from urllib.parse import unquote
                                        path = unquote(location[8:])  # 移除 'file:///'
                                        # Windows路径处理
                                        if path.startswith('/'):
                                            path = path[1:]
                                        path = path.replace('/', '\\')
                                        return path
                                    elif '::' in location:
                                        # CLSID 格式的特殊路径（如"此电脑"）
                                        # 例如：::{20D04FE0-3AEA-1069-A2D8-08002B30309D}
                                        debug_print(f"[Explorer Monitor] Special shell path detected: {location}")
                                        
                                        # 常见的 CLSID 映射
                                        clsid_map = {
                                            '{20D04FE0-3AEA-1069-A2D8-08002B30309D}': 'shell:MyComputerFolder',  # 此电脑
                                            '{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}': 'shell:NetworkPlacesFolder',  # 网络
                                            '{031E4825-7B94-4DC3-B131-E946B44C8DD5}': 'shell:Libraries',  # 库
                                        }
                                        
                                        for clsid, shell_path in clsid_map.items():
                                            if clsid in location:
                                                return shell_path
                                        
                                        # 如果是未知的特殊路径，返回默认位置
                                        debug_print(f"[Explorer Monitor] Unknown CLSID, using default home path")
                                        return QDir.homePath()
                                    else:
                                        # 其他格式的路径
                                        return location
                                else:
                                    debug_print(f"[Explorer Monitor] LocationURL is empty, trying alternative methods...")
                                    
                                    # 尝试获取 LocationName
                                    try:
                                        location_name = window.LocationName
                                        debug_print(f"[Explorer Monitor] LocationName: {location_name}")
                                        
                                        # 根据位置名称推断路径
                                        if location_name in [tr('此电脑'), 'This PC', 'My Computer']:
                                            return 'shell:MyComputerFolder'
                                        elif location_name in [tr('网络'), 'Network']:
                                            return 'shell:NetworkPlacesFolder'
                                        elif location_name in [tr('回收站'), 'Recycle Bin']:
                                            return 'shell:RecycleBinFolder'
                                        # 检查是否为控制面板相关项
                                        elif location_name in [tr('控制面板'), 'Control Panel', tr('用户帐户'), 'User Accounts', 
                                                              tr('程序和功能'), 'Programs and Features', tr('系统'), 'System',
                                                              tr('设备管理器'), 'Device Manager', tr('网络和共享中心'), 'Network and Sharing Center']:
                                            debug_print(f"[Explorer Monitor] Control Panel item detected by LocationName")
                                            return 'shell:ControlPanelFolder'
                                    except Exception:
                                        pass
                                    
                                    # 如果都失败了，返回用户主目录
                                    return QDir.homePath()
                        except Exception as e:
                            debug_print(f"[Explorer Monitor] Error accessing window properties: {e}")
                            continue
                    
                    # 如果第一次没找到，等待一下再试
                    if attempt < 2:
                        time.sleep(0.2)
                        
                except Exception as e:
                    debug_print(f"[Explorer Monitor] Attempt {attempt + 1} failed: {e}")
                    if attempt < 2:
                        time.sleep(0.2)
            
            return None
            
        except Exception as e:
            debug_print(f"[Explorer Monitor] Error getting path: {e}")
            import traceback
            traceback.print_exc()
            return None

    def closeEvent(self, event):
        """窗口关闭时停止服务器和监听"""
        try:
            app = QApplication.instance()
            if app:
                app.removeEventFilter(self)
        except Exception as e:
            print(f"Error removing app event filter: {e}")

        if hasattr(self, 'search_dialogs'):
            for dlg in list(self.search_dialogs):
                try:
                    dlg.close()
                except Exception as e:
                    print(f"Error closing search dialog: {e}")
            self.search_dialogs = []

        if self.chat_panel is not None:
            try:
                self.chat_panel.cleanup()
            except Exception as e:
                print(f"Error cleaning chat panel: {e}")

        self._active_pane = None

        self._append_resource_snapshot_log(reason="close")

        # 先保存会话快照（此时分屏仍处于激活态，确保 split_session 被正确持久化）；
        # 若先合并分屏再保存，_split_active 会被清为 False 导致分屏状态丢失、重启无法恢复
        try:
            self.save_session_snapshot(immediate=True)
        except Exception as e:
            print(f"Error caching tabs: {e}")

        # 保存快照后再合并分屏回左侧，使右侧标签随主流程正常清理
        if getattr(self, '_split_active', False):
            try:
                self._merge_split_back()
            except Exception as e:
                print(f"Error merging split back: {e}")

        # 清除搜索缓存
        try:
            global _search_cache
            _search_cache.clear()
            debug_print(tr("[App] 程序关闭，已清除搜索缓存"))
        except Exception as e:
            print(f"Error clearing search cache: {e}")
        
        # 停止服务器
        self.server_running = False
        if hasattr(self, 'server_socket'):
            try:
                self.server_socket.close()
            except Exception as e:
                print(f"Error closing server socket: {e}")

        if self.server_thread and self.server_thread.is_alive():
            try:
                self.server_thread.join(timeout=1.5)
            except Exception as e:
                print(f"Error waiting for server thread: {e}")
        self.server_thread = None
        
        # 停止Explorer监听
        try:
            self.stop_explorer_monitor()
        except Exception as e:
            print(f"Error stopping explorer monitor: {e}")

        if self.monitor_thread and self.monitor_thread.is_alive():
            try:
                self.monitor_thread.join(timeout=max(1.5, float(getattr(self, 'monitor_interval', 2.0)) + 0.5))
            except Exception as e:
                print(f"Error waiting for monitor thread: {e}")
        self.monitor_thread = None

        for timer_name in (
            '_shortcut_timer',
            '_session_snapshot_debounce_timer',
            '_session_snapshot_timer',
            '_housekeeping_timer',
            '_resource_usage_timer',
            '_config_save_timer',
        ):
            timer = getattr(self, timer_name, None)
            if timer:
                try:
                    timer.stop()
                except Exception:
                    pass
        
        # 停止所有标签页中的定时器和COM对象
        try:
            for i in range(self.tab_widget.count()):
                tab = self.get_tab_widget(i)
                if hasattr(tab, 'cleanup'):
                    try:
                        tab.cleanup()
                    except Exception as cleanup_error:
                        print(f"Error cleaning tab resources: {cleanup_error}")
                if hasattr(tab, '_path_sync_timer') and tab._path_sync_timer:
                    tab._path_sync_timer.stop()
                    tab._path_sync_timer.deleteLater()
                # 清理COM对象
                if hasattr(tab, 'explorer'):
                    try:
                        tab.explorer.clear()
                    except Exception:
                        pass
        except Exception as e:
            print(f"Error stopping timers: {e}")
        
        super().closeEvent(event)



    def open_bookmark_url(self, url):
        # 支持 file:///、file://、shell: 路径和本地绝对路径
        from urllib.parse import unquote
        if url.startswith('file:'):
            # 处理各种file URL格式
            if url.startswith('file://///'):
                # UNC路径: file://///server/share/... -> \\server\share\...
                local_path = '\\\\' + unquote(url[10:]).replace('/', '\\')
            elif url.startswith('file:////'):
                # UNC路径: file:////server/share/... -> \\server\share\...
                local_path = '\\\\' + unquote(url[9:]).replace('/', '\\')
            elif url.startswith('file:///'):
                # 本地路径: file:///C:/... -> C:\...
                local_path = unquote(url[8:])
                if os.name == 'nt' and local_path.startswith('/'):
                    local_path = local_path[1:]
                local_path = local_path.replace('/', '\\')
            else:
                # file://server/share/... -> \\server\share\...
                local_path = '\\\\' + unquote(url[7:]).replace('/', '\\')
            
            # 检查是否是 shell: 路径
            if local_path.startswith('shell:'):
                self.add_new_tab(local_path, is_shell=True)
            elif os.path.exists(local_path):
                self.add_new_tab(local_path)
            else:
                show_toast(self, tr("路径错误"), tr("路径不存在: {}").format(local_path), level="warning")
        elif url.startswith('shell:'):
            # shell:OneDrive 解析为真实路径（避免Shell.Explorer无法正确显示内容）
            if url.lower() == 'shell:onedrive':
                onedrive_path = os.environ.get('OneDrive', '')
                if onedrive_path and os.path.exists(onedrive_path):
                    self.add_new_tab(onedrive_path)
                else:
                    show_toast(self, tr("路径错误"), tr("未找到 OneDrive 文件夹"), level="warning")
            else:
                self.add_new_tab(url, is_shell=True)
        elif os.path.isabs(url) and os.path.exists(url):
            self.add_new_tab(url)
        else:
            show_toast(self, tr("不支持的书签"), tr("暂不支持打开此类型书签: {}").format(url), level="warning")

    def delete_bookmark_by_id(self, bookmark_id):
        """根据ID删除书签"""
        bm = self.bookmark_manager
        tree = bm.get_tree()
        
        def remove_node(parent_node):
            if 'children' in parent_node:
                parent_node['children'] = [
                    child for child in parent_node['children'] 
                    if child.get('id') != bookmark_id
                ]
                # 递归处理子节点
                for child in parent_node['children']:
                    if child.get('type') == 'folder':
                        remove_node(child)
        
        # 在所有根节点中查找并删除
        for root_key, root_node in tree.items():
            remove_node(root_node)
        
        bm.save_bookmarks()
        # 清除现有菜单并重新填充
        self.menu_bar.clear()
        self.populate_bookmark_bar_menu()
    
    def show_bookmark_context_menu(self, pos, bookmark_id, bookmark_name):
        """显示书签右键菜单"""
        debug_print(f"[DEBUG] show_bookmark_context_menu called: pos={pos}, id={bookmark_id}, name={bookmark_name}")
        menu = QMenu(self)
        
        delete_action = menu.addAction(tr("🗑️ 删除书签"))
        delete_action.triggered.connect(lambda: self.confirm_delete_bookmark(bookmark_id, bookmark_name))
        
        debug_print(f"[DEBUG] Showing menu...")
        menu.exec_(pos)
        debug_print(f"[DEBUG] Menu closed")
    
    def confirm_delete_bookmark(self, bookmark_id, bookmark_name):
        """直接删除书签并给出轻量提示"""
        self.delete_bookmark_by_id(bookmark_id)
        show_toast(self, tr("已删除"), tr("书签 '{}' 已删除").format(bookmark_name), level="info")

    def populate_bookmark_bar_menu(self):
        self.ensure_default_icons_on_bookmark_bar()
        self.menu_bar.clear()
        
        bm = self.bookmark_manager
        tree = bm.get_tree()
        bookmark_bar = tree.get('bookmark_bar')
        if not bookmark_bar or 'children' not in bookmark_bar:
            return
        
        # 存储action/menu到节点的映射
        self.bookmark_actions = {}
        self.bookmark_menus = {}  # 存储QMenu到节点的映射
        
        def add_menu_items(parent_menu, node):
            if node.get('type') == 'folder':
                menu = parent_menu.addMenu(f"📁 {node.get('name', '')}")
                # 存储QMenu和节点的映射
                self.bookmark_menus[menu] = node
                # 也为QMenu的menuAction存储映射（用于事件过滤）
                self.bookmark_actions[menu.menuAction()] = node
                # 为子菜单安装事件过滤器
                menu.installEventFilter(self)
                for child in node.get('children', []):
                    add_menu_items(menu, child)
            elif node.get('type') == 'url':
                # 判断是否为四个常用项目
                special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
                name = node.get('name', '')
                is_special = any(name.startswith(icon) for icon in special_icons)
                if is_special:
                    action = parent_menu.addAction(name)
                else:
                    action = parent_menu.addAction(f"📑 {name}")
                url = node.get('url', '')
                action.triggered.connect(lambda checked, u=url: self.open_bookmark_url(u))
                # 存储action和节点的映射
                self.bookmark_actions[action] = node
        # 直接在菜单栏顶层添加
        menubar = self.menu_bar
        # 先添加所有书签和文件夹
        for child in bookmark_bar['children']:
            if child.get('type') == 'folder':
                add_menu_items(menubar, child)
            elif child.get('type') == 'url':
                # 判断是否为四个常用项目
                special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
                name = child.get('name', '')
                is_special = any(name.startswith(icon) for icon in special_icons)
                if is_special:
                    action = menubar.addAction(name)
                else:
                    action = menubar.addAction(f"📑 {name}")
                url = child.get('url', '')
                action.triggered.connect(lambda checked, u=url: self.open_bookmark_url(u))
                # 存储action和节点的映射
                self.bookmark_actions[action] = child
                # 存储action和节点的映射
                self.bookmark_actions[action] = child
        # 仅显示书签内容，不在菜单栏添加“设置”或“书签管理”入口
    
    def on_menubar_context_menu(self, pos):
        """菜单栏右键菜单处理"""
        menubar = self.menu_bar
        action = menubar.actionAt(pos)
        
        if action and hasattr(self, 'bookmark_actions') and action in self.bookmark_actions:
            node = self.bookmark_actions[action]
            bookmark_id = node.get('id')
            bookmark_name = node.get('name', '')
            
            # 检查是否是特殊书签（不允许删除）
            special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
            is_special = any(bookmark_name.startswith(icon) for icon in special_icons)
            
            if not is_special:
                global_pos = menubar.mapToGlobal(pos)
                self.show_bookmark_context_menu(global_pos, bookmark_id, bookmark_name)
    
    def eventFilter(self, obj, event):
        """事件过滤器，处理菜单栏和子菜单的右键菜单"""
        from PyQt5.QtCore import QEvent
        from PyQt5.QtWidgets import QMenu
        from PyQt5.QtGui import QMouseEvent
        

        # 处理主菜单栏的右键点击
        if obj == self.menu_bar:
            if event.type() == QEvent.MouseButtonPress:
                debug_print(f"[DEBUG] MenuBar MouseButtonPress, button: {event.button()}, Qt.RightButton: {Qt.RightButton}")
                
                if event.button() == Qt.RightButton:
                    pos = event.pos()
                    action = self.menu_bar.actionAt(pos)
                    
                    debug_print(f"[DEBUG] MenuBar right click at {pos}, action: {action}")
                    
                    if action:
                        debug_print(f"[DEBUG] Action found, has bookmark_actions: {hasattr(self, 'bookmark_actions')}")
                        if hasattr(self, 'bookmark_actions'):
                            debug_print(f"[DEBUG] bookmark_actions count: {len(self.bookmark_actions)}")
                            debug_print(f"[DEBUG] action in bookmark_actions: {action in self.bookmark_actions}")
                    
                    if action and hasattr(self, 'bookmark_actions') and action in self.bookmark_actions:
                        node = self.bookmark_actions[action]
                        bookmark_id = node.get('id')
                        bookmark_name = node.get('name', '')
                        
                        debug_print(f"[DEBUG] Found bookmark: {bookmark_name} (ID: {bookmark_id})")
                        
                        # 检查是否是特殊书签（不允许删除）
                        special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
                        is_special = any(bookmark_name.startswith(icon) for icon in special_icons)
                        
                        debug_print(f"[DEBUG] Is special bookmark: {is_special}")
                        
                        if not is_special:
                            global_pos = self.menu_bar.mapToGlobal(pos)
                            debug_print(f"[DEBUG] Showing context menu at: {global_pos}")
                            self.show_bookmark_context_menu(global_pos, bookmark_id, bookmark_name)
                            return True  # 事件已处理
                    else:
                        debug_print(f"[DEBUG] Action not in bookmark_actions or no bookmark_actions")
        
        # 处理子菜单（文件夹）的右键点击
        elif isinstance(obj, QMenu):
            if event.type() == QEvent.MouseButtonPress:
                debug_print(f"[DEBUG] QMenu MouseButtonPress, button: {event.button()}")
                
                if event.button() == Qt.RightButton:
                    pos = event.pos()
                    action = obj.actionAt(pos)
                    
                    debug_print(f"[DEBUG] QMenu right click at {pos}, action: {action}")
                    
                    if action and hasattr(self, 'bookmark_actions') and action in self.bookmark_actions:
                        node = self.bookmark_actions[action]
                        bookmark_id = node.get('id')
                        bookmark_name = node.get('name', '')
                        
                        debug_print(f"[DEBUG] Found bookmark in submenu: {bookmark_name} (ID: {bookmark_id})")
                        
                        # 检查是否是特殊书签（不允许删除）
                        special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
                        is_special = any(bookmark_name.startswith(icon) for icon in special_icons)
                        
                        debug_print(f"[DEBUG] Is special bookmark: {is_special}")
                        
                        if not is_special:
                            global_pos = obj.mapToGlobal(pos)
                            debug_print(f"[DEBUG] Showing context menu at: {global_pos}")
                            self.show_bookmark_context_menu(global_pos, bookmark_id, bookmark_name)
                            return True  # 事件已处理
                    else:
                        debug_print(f"[DEBUG] Action not in bookmark_actions")
        
        return super().eventFilter(obj, event)
    
    def toggle_explorer_monitor(self, checked):
        """切换Explorer监听功能"""
        self.config["enable_explorer_monitor"] = checked
        self.save_config()
        
        if checked:
            print("[Settings] Enabling Explorer monitoring")
            self.start_explorer_monitor()
        else:
            print("[Settings] Disabling Explorer monitoring")
            self.stop_explorer_monitor()
        
        show_toast(
            self,
            tr("设置已更新"),
            tr("Explorer窗口监听已{}\n{}").format('启用' if checked else '禁用', '新打开的文件管理器窗口将自动嵌入到标签页中' if checked else '新打开的文件管理器窗口将独立显示'),
            level="info",
        )

    def show_bookmark_manager_dialog(self):
        self.bookmark_manager_dialog = BookmarkManagerDialog(self.bookmark_manager, self)
        self.bookmark_manager_dialog.exec_()
        self.bookmark_manager_dialog = None

class SettingsDialog(QDialog):

    def __init__(self, config, parent=None):
        from PyQt5.QtWidgets import QDialogButtonBox, QLabel, QGroupBox, QComboBox, QHBoxLayout, QVBoxLayout, QCheckBox, QSpinBox
        super().__init__(parent)
        self.setWindowTitle(tr("设置"))
        # 设置为不可调边框的对话框
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        # 宽度固定，高度按内容自动计算
        self.setFixedWidth(700)
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(12)

        # 按类型分页：常规 / 手势与快捷键 / 工具集成 / AI 助手 / 高级
        from PyQt5.QtWidgets import QTabWidget, QWidget
        self.settings_tabs = QTabWidget(self)

        general_page = QWidget()
        general_layout = QVBoxLayout(general_page)
        general_layout.setSpacing(6)

        input_page = QWidget()
        input_layout = QVBoxLayout(input_page)
        input_layout.setSpacing(6)

        tools_page = QWidget()
        tools_page_layout = QVBoxLayout(tools_page)
        tools_page_layout.setSpacing(6)

        ai_page = QWidget()
        ai_page_layout = QVBoxLayout(ai_page)
        ai_page_layout.setSpacing(6)

        advanced_page = QWidget()
        advanced_layout = QVBoxLayout(advanced_page)
        advanced_layout.setSpacing(6)

        # 紧凑化所有GroupBox和控件的布局
        def compact_groupbox(groupbox):
            lay = groupbox.layout()
            if lay:
                lay.setContentsMargins(6, 6, 6, 6)
                lay.setSpacing(4)

        # 控件紧凑化工具
        def compact_widget(widget):
            if hasattr(widget, 'setStyleSheet'):
                widget.setStyleSheet("font-size: 10.5pt; padding: 2px 4px;")

        # 路径栏分隔符设置组
        pathbar_group = QGroupBox(tr("路径栏分隔符设置"))
        pathbar_layout = QHBoxLayout()
        pathbar_layout.addWidget(QLabel(tr("路径栏拷贝分隔符:")))
        self.path_separator_combo = QComboBox(self)
        self.path_separator_combo.addItem("/", "/")
        self.path_separator_combo.addItem("\\", "\\")
        sep = config.get("breadcrumb_copy_separator", "/")
        idx = 0 if sep == "/" else 1
        self.path_separator_combo.setCurrentIndex(idx)
        self.path_separator_combo.setToolTip(tr("设置从路径栏拷贝时使用的分隔符"))
        pathbar_layout.addWidget(self.path_separator_combo)
        pathbar_layout.addStretch(1)
        pathbar_group.setLayout(pathbar_layout)
        compact_groupbox(pathbar_group)
        for i in range(pathbar_layout.count()):
            w = pathbar_layout.itemAt(i).widget()
            if w: compact_widget(w)
        general_layout.addWidget(pathbar_group)

        # Explorer监听设置组
        monitor_group = QGroupBox(tr("Explorer监听设置"))
        monitor_layout = QVBoxLayout()
        self.monitor_cb = QCheckBox(tr("监听新Explorer窗口"), self)
        self.monitor_cb.setChecked(config.get("enable_explorer_monitor", True))
        self.monitor_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        monitor_layout.addWidget(self.monitor_cb)
        # 状态栏右侧 CPU/内存占用显示
        self.resource_usage_cb = QCheckBox(tr("状态栏显示 CPU/内存占用"), self)
        self.resource_usage_cb.setChecked(config.get("show_resource_usage_in_statusbar", False))
        self.resource_usage_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.resource_usage_cb.setToolTip(tr("在状态栏右侧实时显示整机 CPU 与内存占用，每2秒刷新"))
        monitor_layout.addWidget(self.resource_usage_cb)
        # 监听间隔设置
        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel(tr("监听间隔（秒）:")))
        from PyQt5.QtWidgets import QDoubleSpinBox
        self.interval_spinbox = QDoubleSpinBox()
        self.interval_spinbox.setRange(0.5, 10.0)
        self.interval_spinbox.setSingleStep(0.5)
        self.interval_spinbox.setValue(config.get("explorer_monitor_interval", 2.0))
        self.interval_spinbox.setToolTip(tr("检查新Explorer窗口的时间间隔，更长的间隔降低CPU占用"))
        interval_layout.addWidget(self.interval_spinbox)
        interval_layout.addWidget(QLabel(tr("（推荐: 2.0秒）")))
        interval_layout.addStretch(1)
        monitor_layout.addLayout(interval_layout)
        monitor_group.setLayout(monitor_layout)
        compact_groupbox(monitor_group)
        for i in range(monitor_layout.count()):
            item = monitor_layout.itemAt(i)
            if item.layout():
                for j in range(item.layout().count()):
                    w = item.layout().itemAt(j).widget()
                    if w: compact_widget(w)
            elif item.widget():
                compact_widget(item.widget())
        general_layout.addWidget(monitor_group)

        # 调试设置组
        debug_group = QGroupBox(tr("调试设置"))
        debug_layout = QVBoxLayout()
        self.debug_mode_cb = QCheckBox(tr("启用调试输出（输出到终端）"), self)
        self.debug_mode_cb.setChecked(config.get("debug_mode", False))
        self.debug_mode_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.debug_mode_cb.setToolTip(tr("启用后将在终端输出调试信息，用于开发和问题排查"))
        debug_layout.addWidget(self.debug_mode_cb)
        self.explorer_monitor_debug_cb = QCheckBox(tr("启用 Explorer Monitor 调试输出"), self)
        self.explorer_monitor_debug_cb.setChecked(config.get("explorer_monitor_debug", False))
        self.explorer_monitor_debug_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.explorer_monitor_debug_cb.setToolTip(tr("单独控制 Explorer Monitor 的日志输出（需要先启用调试输出）"))
        debug_layout.addWidget(self.explorer_monitor_debug_cb)
        self.resource_snapshot_logging_cb = QCheckBox(tr("启用资源快照日志"), self)
        self.resource_snapshot_logging_cb.setChecked(config.get("resource_snapshot_logging", False))
        self.resource_snapshot_logging_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.resource_snapshot_logging_cb.setToolTip(tr("定时写入 runtime_health.log，用于观察长期运行时的内存和线程趋势"))
        debug_layout.addWidget(self.resource_snapshot_logging_cb)

        resource_interval_layout = QHBoxLayout()
        resource_interval_layout.addWidget(QLabel(tr("资源快照间隔（分钟）:")))
        self.resource_snapshot_interval_spin = QSpinBox(self)
        self.resource_snapshot_interval_spin.setRange(1, 60)
        self.resource_snapshot_interval_spin.setSingleStep(1)
        interval_minutes = max(1, int(config.get("resource_snapshot_interval_ms", HOUSEKEEPING_INTERVAL_MS) / 60000))
        self.resource_snapshot_interval_spin.setValue(interval_minutes)
        self.resource_snapshot_interval_spin.setToolTip(tr("资源快照日志写入周期，建议 5 分钟或更长"))
        self.resource_snapshot_interval_spin.setEnabled(self.resource_snapshot_logging_cb.isChecked())
        self.resource_snapshot_logging_cb.toggled.connect(self.resource_snapshot_interval_spin.setEnabled)
        resource_interval_layout.addWidget(self.resource_snapshot_interval_spin)
        resource_interval_layout.addWidget(QLabel(tr("分钟")))
        resource_interval_layout.addStretch(1)
        debug_layout.addLayout(resource_interval_layout)

        resource_log_layout = QHBoxLayout()
        self.open_resource_log_btn = QPushButton(tr("打开资源日志"), self)
        self.open_resource_log_btn.setToolTip(tr("打开 runtime_health.log；如果日志尚未生成，则打开所在目录"))
        self.open_resource_log_btn.clicked.connect(self._open_resource_snapshot_log)
        resource_log_layout.addWidget(self.open_resource_log_btn)
        resource_log_layout.addStretch(1)
        debug_layout.addLayout(resource_log_layout)
        debug_group.setLayout(debug_layout)
        compact_groupbox(debug_group)
        for i in range(debug_layout.count()):
            item = debug_layout.itemAt(i)
            if item.layout():
                for j in range(item.layout().count()):
                    w = item.layout().itemAt(j).widget()
                    if w: compact_widget(w)
            elif item.widget():
                compact_widget(item.widget())
        advanced_layout.addWidget(debug_group)

        # 标签页设置组
        tabs_group = QGroupBox(tr("标签页设置"))
        tabs_layout = QVBoxLayout()
        self.cache_tabs_cb = QCheckBox(tr("关闭时缓存当前标签页，下次启动时恢复"), self)
        self.cache_tabs_cb.setChecked(config.get("enable_cache_tabs", True))
        self.cache_tabs_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.cache_tabs_cb.setToolTip(tr("关闭软件时保存非固定标签，下次启动时自动恢复（不包括固定标签）"))
        tabs_layout.addWidget(self.cache_tabs_cb)
        tabs_group.setLayout(tabs_layout)
        compact_groupbox(tabs_group)
        for i in range(tabs_layout.count()):
            w = tabs_layout.itemAt(i).widget()
            if w: compact_widget(w)
        general_layout.addWidget(tabs_group)

        # 鼠标手势设置组
        gesture_group = QGroupBox(tr("鼠标手势设置"))
        gesture_layout = QVBoxLayout()
        self.mouse_gestures_cb = QCheckBox(tr("启用鼠标手势（按住右键画线）"), self)
        self.mouse_gestures_cb.setChecked(config.get("enable_mouse_gestures", True))
        self.mouse_gestures_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.mouse_gestures_cb.setToolTip(tr("在文件区域按住鼠标右键画线即可触发导航操作；关闭后右键恢复为普通右键菜单"))
        gesture_layout.addWidget(self.mouse_gestures_cb)
        gesture_help = QLabel(
            tr("← 向左：后退\n→ 向右：前进\n↓ 向下：关闭当前标签页\n↑ 向上：打开新标签页\n↑↓ 上下：刷新\n↓↑ 下上：返回上级目录\n↓→ 下右：恢复关闭的标签页")
        )
        gesture_help.setStyleSheet(
            "QLabel { color: #555; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }"
        )
        gesture_layout.addWidget(gesture_help)
        gesture_group.setLayout(gesture_layout)
        compact_groupbox(gesture_group)
        for i in range(gesture_layout.count()):
            w = gesture_layout.itemAt(i).widget()
            if w: compact_widget(w)
        input_layout.addWidget(gesture_group)

        # 开机启动设置组
        startup_group = QGroupBox(tr("开机启动设置"))
        startup_layout = QVBoxLayout()
        self.auto_startup_cb = QCheckBox(tr("开机自动启动 TabExplorer"), self)
        self.auto_startup_cb.setChecked(self._is_auto_startup_enabled())
        self.auto_startup_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.auto_startup_cb.setToolTip(tr("在 Windows 启动时自动运行 TabExplorer.exe"))
        startup_layout.addWidget(self.auto_startup_cb)
        startup_group.setLayout(startup_layout)
        compact_groupbox(startup_group)
        for i in range(startup_layout.count()):
            w = startup_layout.itemAt(i).widget()
            if w: compact_widget(w)
        general_layout.addWidget(startup_group)

        # Git 工具设置组
        git_group = QGroupBox(tr("Git 工具设置"))
        git_layout = QVBoxLayout()
        self.tortoisegit_buttons_cb = QCheckBox(tr("显示 TortoiseGit 快捷按钮（标题栏）"), self)
        self.tortoisegit_buttons_cb.setChecked(config.get("enable_tortoisegit_buttons", False))
        self.tortoisegit_buttons_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.tortoisegit_buttons_cb.setToolTip(tr("在标题栏显示 Git Log 和 Git Commit 快捷按钮"))
        git_layout.addWidget(self.tortoisegit_buttons_cb)

        terminal_pref_layout = QHBoxLayout()
        terminal_pref_layout.addWidget(QLabel(tr("默认终端:")))
        self.preferred_terminal_combo = QComboBox(self)
        self.preferred_terminal_combo.addItem("CMD", "cmd")
        self.preferred_terminal_combo.addItem("PowerShell", "powershell")
        self.preferred_terminal_combo.addItem("Git Bash", "git-bash")
        preferred_terminal = normalize_terminal_tool_name(config.get("preferred_terminal_tool", "cmd"))
        terminal_idx = max(0, self.preferred_terminal_combo.findData(preferred_terminal))
        self.preferred_terminal_combo.setCurrentIndex(terminal_idx)
        self.preferred_terminal_combo.setToolTip(tr("路径栏输入 terminal 或 term 时使用的默认终端"))
        terminal_pref_layout.addWidget(self.preferred_terminal_combo)
        terminal_pref_layout.addStretch(1)
        git_layout.addLayout(terminal_pref_layout)
        git_group.setLayout(git_layout)
        compact_groupbox(git_group)
        for i in range(git_layout.count()):
            item = git_layout.itemAt(i)
            if item.layout():
                for j in range(item.layout().count()):
                    w = item.layout().itemAt(j).widget()
                    if w: compact_widget(w)
            elif item.widget():
                compact_widget(item.widget())
        tools_page_layout.addWidget(git_group)

        # 快捷方式设置组（右侧独立分组）
        shortcuts_group = QGroupBox(tr("快捷方式设置"))
        shortcuts_layout = QVBoxLayout()
        self.title_shortcuts_cb = QCheckBox(tr("启用标题栏启动区（可拖拽应用、快捷方式或脚本并点击启动）"), self)
        self.title_shortcuts_cb.setChecked(config.get("enable_title_shortcuts", True))
        self.title_shortcuts_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.title_shortcuts_cb.setToolTip(tr("拖拽应用、快捷方式或脚本到标题栏 Git 左侧区域，后续可一键启动"))
        shortcuts_layout.addWidget(self.title_shortcuts_cb)
        shortcuts_group.setLayout(shortcuts_layout)
        compact_groupbox(shortcuts_group)
        for i in range(shortcuts_layout.count()):
            w = shortcuts_layout.itemAt(i).widget()
            if w: compact_widget(w)
        tools_page_layout.addWidget(shortcuts_group)

        # 快捷键设置组
        hotkey_group = QGroupBox(tr("快捷键设置"))
        hotkey_layout = QVBoxLayout()
        hotkeys = config.get("hotkeys", {})
        self.hotkey_new_tab = QCheckBox(tr("Ctrl+T - 新建标签页"))
        self.hotkey_new_tab.setChecked(hotkeys.get("new_tab", True))
        hotkey_layout.addWidget(self.hotkey_new_tab)
        self.hotkey_close_tab = QCheckBox(tr("Ctrl+W - 关闭当前标签页"))
        self.hotkey_close_tab.setChecked(hotkeys.get("close_tab", True))
        hotkey_layout.addWidget(self.hotkey_close_tab)
        self.hotkey_reopen_tab = QCheckBox(tr("Ctrl+Shift+T - 恢复关闭的标签页"))
        self.hotkey_reopen_tab.setChecked(hotkeys.get("reopen_tab", True))
        hotkey_layout.addWidget(self.hotkey_reopen_tab)
        self.hotkey_switch_tab = QCheckBox(tr("Ctrl+Tab / Ctrl+Shift+Tab - 切换标签页"))
        self.hotkey_switch_tab.setChecked(hotkeys.get("switch_tab", True))
        hotkey_layout.addWidget(self.hotkey_switch_tab)
        self.hotkey_search = QCheckBox(tr("Ctrl+F - 打开搜索对话框"))
        self.hotkey_search.setChecked(hotkeys.get("search", True))
        hotkey_layout.addWidget(self.hotkey_search)
        self.hotkey_quick_find_current_dir = QCheckBox(tr("Ctrl+G - 检索当前目录文件夹/文件名"))
        self.hotkey_quick_find_current_dir.setChecked(hotkeys.get("quick_find_current_dir", True))
        hotkey_layout.addWidget(self.hotkey_quick_find_current_dir)
        self.hotkey_navigate = QCheckBox(tr("Alt+Left/Right - 前进/后退"))
        self.hotkey_navigate.setChecked(hotkeys.get("navigate", True))
        hotkey_layout.addWidget(self.hotkey_navigate)
        self.hotkey_go_up = QCheckBox(tr("Alt+Up - 返回上级目录"))
        self.hotkey_go_up.setChecked(hotkeys.get("go_up", True))
        hotkey_layout.addWidget(self.hotkey_go_up)
        self.hotkey_refresh = QCheckBox(tr("F5 - 刷新当前路径"))
        self.hotkey_refresh.setChecked(hotkeys.get("refresh", True))
        hotkey_layout.addWidget(self.hotkey_refresh)
        self.hotkey_add_bookmark = QCheckBox(tr("Ctrl+D - 添加当前路径到书签"))
        self.hotkey_add_bookmark.setChecked(hotkeys.get("add_bookmark", True))
        hotkey_layout.addWidget(self.hotkey_add_bookmark)
        self.hotkey_copy_filename = QCheckBox(tr("Alt+Z - 复制选中文件名（含后缀）"))
        self.hotkey_copy_filename.setChecked(hotkeys.get("copy_filename", True))
        hotkey_layout.addWidget(self.hotkey_copy_filename)
        self.hotkey_copy_filepath = QCheckBox(tr("Alt+X - 复制文件路径\\文件名"))
        self.hotkey_copy_filepath.setChecked(hotkeys.get("copy_filepath", True))
        hotkey_layout.addWidget(self.hotkey_copy_filepath)
        self.hotkey_split_view = QCheckBox(tr("F3 - 左右分屏对比"))
        self.hotkey_split_view.setChecked(hotkeys.get("split_view", True))
        hotkey_layout.addWidget(self.hotkey_split_view)
        # 提示信息（放在快捷键设置框内）
        tip_label = QLabel(tr("💡 提示：取消勾选可禁用对应的快捷键"))
        tip_label.setStyleSheet("QLabel { color: #666; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }")
        hotkey_layout.addWidget(tip_label)
        hotkey_group.setLayout(hotkey_layout)
        compact_groupbox(hotkey_group)
        for i in range(hotkey_layout.count()):
            w = hotkey_layout.itemAt(i).widget()
            if w: compact_widget(w)
        # 手势与快捷键页放快捷键设置组
        input_layout.addWidget(hotkey_group)

            # AI 助手设置组
        ai_group = QGroupBox(tr("AI 助手设置"))
        ai_layout = QVBoxLayout()
        ai_layout.setSpacing(6)
        from PyQt5.QtWidgets import QLineEdit, QComboBox as _CB2
        # 启用 AI 助手开关
        self.ai_enabled_cb = QCheckBox(tr("启用 AI 助手（显示标题栏机器人按钮🤖）"))
        self.ai_enabled_cb.setChecked(config.get("ai_chat", {}).get("enabled", True))
        ai_layout.addWidget(self.ai_enabled_cb)

        # ── 免费服务商预设 ──────────────────────────────────────────────────────
        # 格式: (显示名称, api_url, 默认model, 获取Key说明)
        _AI_PRESETS = [
            (tr("── 请选择预设服务商 ──"), "", "", ""),
            (tr("Groq（免费·极速·推荐）"),
             "https://api.groq.com/openai/v1",
             "llama-3.3-70b-versatile",
             tr("免费注册获取Key: https://console.groq.com/keys")),
            (tr("SiliconFlow 硅基流动（免费额度·国内快）"),
             "https://api.siliconflow.cn/v1",
             "Qwen/Qwen2.5-7B-Instruct",
             tr("免费注册获取Key: https://cloud.siliconflow.cn")),
            (tr("DeepSeek（注册送额度·中文强）"),
             "https://api.deepseek.com/v1",
             "deepseek-chat",
             tr("注册获取Key: https://platform.deepseek.com/api_keys")),
            (tr("Google Gemini（免费版）"),
             "https://generativelanguage.googleapis.com/v1beta/openai",
             "gemini-2.0-flash",
             tr("免费获取Key: https://aistudio.google.com/app/apikey")),
            (tr("OpenRouter（含永久免费模型）"),
             "https://openrouter.ai/api/v1",
             "meta-llama/llama-3.3-70b-instruct:free",
             tr("注册获取Key: https://openrouter.ai/keys")),
            (tr("本地 LM Studio（无需Key）"),
             "http://localhost:1234/v1",
             "local-model",
             tr("启动 LM Studio → Local Server 后使用")),
            (tr("本地 Ollama（无需Key）"),
             "http://localhost:11434/v1",
             "qwen2.5:7b",
             tr("安装 Ollama 并运行模型后使用")),
            (tr("── 自定义（手动填写下方） ──"), "", "", ""),
        ]

        preset_row = QHBoxLayout()
        preset_row.addWidget(QLabel(tr("快速选择:")))
        self.ai_preset_combo = _CB2()
        for name, url, model, tip in _AI_PRESETS:
            self.ai_preset_combo.addItem(name, (url, model, tip))
        self.ai_preset_combo.setToolTip(tr("选择预设后自动填充地址和模型名，然后只需粘贴对应的 API Key"))
        preset_row.addWidget(self.ai_preset_combo, 1)
        ai_layout.addLayout(preset_row)

        # 获取Key提示标签
        self.ai_key_tip_label = QLabel("")
        self.ai_key_tip_label.setStyleSheet(
            "color:#1565C0; font-size:8.5pt; padding:2px 4px;"
            "background:#E3F2FD; border-radius:3px;"
        )
        self.ai_key_tip_label.setWordWrap(True)
        self.ai_key_tip_label.setVisible(False)
        ai_layout.addWidget(self.ai_key_tip_label)

        def _on_preset_changed(idx):
            url, model, tip = self.ai_preset_combo.itemData(idx)
            if url:
                self.ai_api_url_edit.setText(url)
                self.ai_model_edit.setText(model)
            if tip:
                self.ai_key_tip_label.setText(f"💡 {tip}")
                self.ai_key_tip_label.setVisible(True)
            else:
                self.ai_key_tip_label.setVisible(False)

        self.ai_preset_combo.currentIndexChanged.connect(_on_preset_changed)

        # API 地址
        api_url_row = QHBoxLayout()
        api_url_row.addWidget(QLabel(tr("API 地址:")))
        self.ai_api_url_edit = QLineEdit()
        self.ai_api_url_edit.setPlaceholderText(tr("例: https://api.groq.com/openai/v1"))
        self.ai_api_url_edit.setText(config.get("ai_chat", {}).get("api_url", ""))
        self.ai_api_url_edit.setToolTip(tr("填写 OpenAI 兼容 API 的基础地址（不含 /chat/completions）"))
        api_url_row.addWidget(self.ai_api_url_edit, 1)
        ai_layout.addLayout(api_url_row)
        # API 密钥
        api_key_row = QHBoxLayout()
        api_key_row.addWidget(QLabel(tr("API 密钥:")))
        self.ai_api_key_edit = QLineEdit()
        self.ai_api_key_edit.setPlaceholderText(tr("粘贴从服务商网站获取的 Key（本地模型可留空）"))
        self.ai_api_key_edit.setEchoMode(QLineEdit.Password)
        self.ai_api_key_edit.setText(config.get("ai_chat", {}).get("api_key", ""))
        # 显示/隐藏密钥按钮
        eye_btn = QPushButton("👁")
        eye_btn.setFixedSize(26, 26)
        eye_btn.setCheckable(True)
        eye_btn.setStyleSheet(
            "QPushButton{border:none;background:transparent;font-size:11pt;}"
            "QPushButton:hover{background:#e5e5e5;border-radius:3px;}"
        )
        def _toggle_key_visibility(checked):
            self.ai_api_key_edit.setEchoMode(
                QLineEdit.Normal if checked else QLineEdit.Password
            )
        eye_btn.toggled.connect(_toggle_key_visibility)
        api_key_row.addWidget(self.ai_api_key_edit, 1)
        api_key_row.addWidget(eye_btn)
        ai_layout.addLayout(api_key_row)
        # 模型名称
        model_row = QHBoxLayout()
        model_row.addWidget(QLabel(tr("模型名称:")))
        self.ai_model_edit = QLineEdit()
        self.ai_model_edit.setPlaceholderText(tr("例: llama-3.3-70b-versatile"))
        self.ai_model_edit.setText(config.get("ai_chat", {}).get("model", ""))
        model_row.addWidget(self.ai_model_edit, 1)
        ai_layout.addLayout(model_row)
        # 系统提示词
        from PyQt5.QtWidgets import QPlainTextEdit as _PE
        sp_label = QLabel(tr("系统提示词（留空使用默认）:"))
        ai_layout.addWidget(sp_label)
        self.ai_system_prompt_edit = _PE()
        self.ai_system_prompt_edit.setFixedHeight(56)
        self.ai_system_prompt_edit.setPlaceholderText(tr("留空则使用内置提示词（支持 [OPEN_DIR:] 和 [RUN_SCRIPT:] 指令）"))
        self.ai_system_prompt_edit.setPlainText(config.get("ai_chat", {}).get("system_prompt", ""))
        ai_layout.addWidget(self.ai_system_prompt_edit)
        # 面板宽度
        panel_w_row = QHBoxLayout()
        panel_w_row.addWidget(QLabel(tr("面板宽度 (px):")))
        from PyQt5.QtWidgets import QSpinBox as _SB
        self.ai_panel_width_spin = _SB()
        self.ai_panel_width_spin.setRange(200, 800)
        self.ai_panel_width_spin.setValue(config.get("ai_chat", {}).get("panel_width", 360))
        panel_w_row.addWidget(self.ai_panel_width_spin)
        panel_w_row.addStretch(1)
        ai_layout.addLayout(panel_w_row)
        ai_group.setLayout(ai_layout)
        compact_groupbox(ai_group)
        ai_page_layout.addWidget(ai_group)

        # 各页末尾加弹性空间
        general_layout.addStretch(1)
        input_layout.addStretch(1)
        tools_page_layout.addStretch(1)
        ai_page_layout.addStretch(1)
        advanced_layout.addStretch(1)

        # 组装分页
        self.settings_tabs.addTab(general_page, tr("常规"))
        self.settings_tabs.addTab(input_page, tr("手势与快捷键"))
        self.settings_tabs.addTab(tools_page, tr("工具集成"))
        self.settings_tabs.addTab(ai_page, tr("AI 助手"))
        self.settings_tabs.addTab(advanced_page, tr("高级"))

        # 创建主垂直布局，放置内容和底部区域
        main_vertical_layout = QVBoxLayout()
        main_vertical_layout.addWidget(self.settings_tabs, 1)
        
        # 底部区域（横跨整个宽度）
        bottom_layout = QVBoxLayout()
        bottom_layout.setSpacing(8)

        # 语言 / Language 设置行
        lang_row = QHBoxLayout()
        lang_row.addWidget(QLabel(tr("语言 / Language:")))
        self.lang_combo = QComboBox(self)
        self.lang_combo.addItem("中文", "zh")
        self.lang_combo.addItem("English", "en")
        current_lang = config.get("language", "zh")
        self.lang_combo.setCurrentIndex(0 if current_lang == "zh" else 1)
        lang_row.addWidget(self.lang_combo)
        lang_row.addStretch(1)
        bottom_layout.addLayout(lang_row)
        
        # 检查更新链接
        update_link = QLabel()
        update_link.setText(tr('检查更新: <a href="https://github.com/caojinyuan/TabEx/releases">https://github.com/caojinyuan/TabEx/releases</a>'))
        update_link.setOpenExternalLinks(True)
        update_link.setStyleSheet("QLabel { padding: 10px; font-size: 10pt; }")
        update_link.setTextFormat(Qt.RichText)
        update_link.setToolTip(tr("点击链接在浏览器中打开 GitHub Releases 页面"))
        update_link.setWordWrap(True)
        bottom_layout.addWidget(update_link)
        
        # 按钮区域
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        bottom_layout.addWidget(buttons)
        
        main_vertical_layout.addLayout(bottom_layout)
        
        # 主布局拼接
        main_layout.addLayout(main_vertical_layout)

        # 高度根据内容自适应，并限制不超过屏幕可用高度
        try:
            from PyQt5.QtWidgets import QApplication
            from PyQt5.QtCore import QTimer

            def _apply_auto_height():
                self.adjustSize()
                target_h = self.sizeHint().height()
                screen_geo = QApplication.primaryScreen().availableGeometry()
                max_h = int(screen_geo.height() * 0.9)
                self.setFixedHeight(min(target_h, max_h))

            # 延迟到布局稳定后计算，避免初次 sizeHint 偏差
            QTimer.singleShot(0, _apply_auto_height)
        except Exception:
            # 兜底：给一个较合理默认高度
            self.resize(600, 620)
    
    def _is_auto_startup_enabled(self):
        """检查是否已启用开机启动"""
        try:
            startup_path = self._get_startup_shortcut_path()
            return os.path.exists(startup_path)
        except Exception:
            return False
    
    def _get_startup_shortcut_path(self):
        """获取启动项快捷方式路径"""
        # 使用环境变量获取启动文件夹，避免依赖 winshell
        startup_folder = os.path.join(
            os.environ.get('APPDATA', ''),
            'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup'
        )
        return os.path.join(startup_folder, "TabExplorer.lnk")
    
    def _set_auto_startup(self, enabled):
        """设置开机启动"""
        try:
            shortcut_path = self._get_startup_shortcut_path()
            
            if enabled:
                # 获取 TabExplorer.exe 路径
                exe_path = self._get_exe_path()
                if not exe_path or not os.path.exists(exe_path):
                    show_toast(self.parent(), tr("错误"), tr("未找到 TabExplorer.exe，请确保程序已正确安装"), level="error")
                    return False
                
                # 创建快捷方式（使用 win32com）
                from win32com.client import Dispatch
                
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = exe_path
                shortcut.WorkingDirectory = os.path.dirname(exe_path)
                shortcut.IconLocation = exe_path
                shortcut.save()
                
                show_toast(self.parent(), tr("成功"), tr("已启用开机自动启动"), level="success")
                return True
            else:
                # 删除快捷方式
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                    show_toast(self.parent(), tr("成功"), tr("已禁用开机自动启动"), level="success")
                return True
        except Exception as e:
            show_toast(self.parent(), tr("错误"), tr("设置开机启动失败: {}").format(e), level="error")
            return False
    
    def _get_exe_path(self):
        """获取 TabExplorer.exe 路径"""
        # 如果是打包的exe，使用sys.executable
        import sys
        if getattr(sys, 'frozen', False):
            return sys.executable
        
        # 如果是开发环境，尝试查找同目录下的 TabExplorer.exe
        script_dir = get_app_base_dir()
        exe_path = os.path.join(script_dir, "TabExplorer.exe")
        if os.path.exists(exe_path):
            return exe_path
        
        # 查找上级目录
        parent_dir = os.path.dirname(script_dir)
        exe_path = os.path.join(parent_dir, "TabExplorer.exe")
        if os.path.exists(exe_path):
            return exe_path
        
        return None

    def _open_resource_snapshot_log(self):
        log_path = get_app_data_path('runtime_health.log')
        log_dir = os.path.dirname(log_path)
        try:
            if os.path.exists(log_path):
                os.startfile(log_path)
            elif os.path.isdir(log_dir):
                os.startfile(log_dir)
                if self.parent():
                    show_toast(self.parent(), tr("资源日志"), tr("日志尚未生成，已打开日志目录"), level="info")
            else:
                if self.parent():
                    show_toast(self.parent(), tr("资源日志"), tr("日志目录不存在"), level="warning")
        except Exception as e:
            if self.parent():
                show_toast(self.parent(), tr("资源日志"), tr("无法打开资源日志: {}").format(e), level="error")
    
    def accept(self):
        """保存设置"""
        # 保存所有设置到 parent (MainWindow)
        if self.parent():
            # 处理开机启动设置
            auto_startup_enabled = self.auto_startup_cb.isChecked()
            current_enabled = self._is_auto_startup_enabled()
            if auto_startup_enabled != current_enabled:
                self._set_auto_startup(auto_startup_enabled)
            
            self.parent().config["enable_explorer_monitor"] = self.monitor_cb.isChecked()
            self.parent().config["explorer_monitor_interval"] = self.interval_spinbox.value()
            self.parent().config["debug_mode"] = self.debug_mode_cb.isChecked()
            self.parent().config["explorer_monitor_debug"] = self.explorer_monitor_debug_cb.isChecked()
            self.parent().config["resource_snapshot_logging"] = self.resource_snapshot_logging_cb.isChecked()
            self.parent().config["resource_snapshot_interval_ms"] = self.resource_snapshot_interval_spin.value() * 60 * 1000
            self.parent().config["show_resource_usage_in_statusbar"] = self.resource_usage_cb.isChecked()
            self.parent().config["enable_cache_tabs"] = self.cache_tabs_cb.isChecked()
            self.parent().config["enable_tortoisegit_buttons"] = self.tortoisegit_buttons_cb.isChecked()
            self.parent().config["preferred_terminal_tool"] = normalize_terminal_tool_name(self.preferred_terminal_combo.currentData())
            # 保存路径栏分隔符设置
            self.parent().config["breadcrumb_copy_separator"] = self.path_separator_combo.currentData()
            
            # 保存快捷键配置
            self.parent().config["hotkeys"] = {
                "new_tab": self.hotkey_new_tab.isChecked(),
                "close_tab": self.hotkey_close_tab.isChecked(),
                "reopen_tab": self.hotkey_reopen_tab.isChecked(),
                "switch_tab": self.hotkey_switch_tab.isChecked(),
                "search": self.hotkey_search.isChecked(),
                "quick_find_current_dir": self.hotkey_quick_find_current_dir.isChecked(),
                "navigate": self.hotkey_navigate.isChecked(),
                "go_up": self.hotkey_go_up.isChecked(),
                "refresh": self.hotkey_refresh.isChecked(),
                "add_bookmark": self.hotkey_add_bookmark.isChecked(),
                "copy_filename": self.hotkey_copy_filename.isChecked(),
                "copy_filepath": self.hotkey_copy_filepath.isChecked()
            }
            
            # 保存到文件
            self.parent().save_config()
            # 刷新所有tab的路径栏分隔符显示（立即生效）
            mainwin = self.parent()
            if hasattr(mainwin, 'tab_widget') and hasattr(mainwin, 'get_tab_widget'):
                for i in range(mainwin.tab_widget.count()):
                    tab = mainwin.get_tab_widget(i)
                    if hasattr(tab, 'path_bar') and hasattr(tab, 'current_path'):
                        # 强制重设路径，确保分隔符立即生效
                        tab.path_bar.set_path(tab.current_path)
            # 应用设置
            set_debug_mode(self.parent().config.get("debug_mode", False))
            set_explorer_monitor_debug(self.parent().config.get("explorer_monitor_debug", False))
            if hasattr(self.parent(), '_housekeeping_timer') and self.parent()._housekeeping_timer:
                self.parent()._housekeeping_timer.start(self.parent()._get_housekeeping_interval_ms())
            if hasattr(self.parent(), 'apply_resource_usage_config'):
                self.parent().apply_resource_usage_config()
            self.parent().apply_tortoisegit_buttons_config()
            # 重新设置快捷键
            self.parent().setup_shortcuts()
            # 保存 AI 助手设置
            ai_enabled = self.ai_enabled_cb.isChecked()
            self.parent().config["ai_chat"] = {
                "enabled": ai_enabled,
                "api_url": self.ai_api_url_edit.text().strip(),
                "api_key": self.ai_api_key_edit.text().strip(),
                "model": self.ai_model_edit.text().strip() or "gpt-3.5-turbo",
                "system_prompt": self.ai_system_prompt_edit.toPlainText().strip(),
                "panel_width": self.ai_panel_width_spin.value(),
            }

            # 保存语言设置并应用
            new_lang = self.lang_combo.currentData()
            old_lang = self.parent().config.get("language", "zh")
            self.parent().config["language"] = new_lang
            _set_app_language(new_lang)

            self.parent().save_config()
            # 立即更新标题栏 AI 按钮显隐
            if hasattr(self.parent(), 'ai_chat_btn'):
                self.parent().ai_chat_btn.setVisible(ai_enabled)
                if not ai_enabled and hasattr(self.parent(), 'chat_panel'):
                    self.parent().chat_panel.setVisible(False)

            # 语言切换提示（部分静态 UI 需重启生效）
            if new_lang != old_lang:
                from PyQt5.QtWidgets import QMessageBox
                if new_lang == "en":
                    QMessageBox.information(self, "Language Changed",
                        "Language set to English.\nSome UI elements will update after restart.")
                else:
                    QMessageBox.information(self, "语言已切换",
                        "界面语言已切换为中文。\n部分界面元素重启后生效。")
        
        super().accept()

# 书签管理对话框（初步框架，后续可扩展重命名/新建/删除等功能）
from PyQt5.QtWidgets import QDialog, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QPushButton, QHBoxLayout, QInputDialog, QLabel
class BookmarkManagerDialog(QDialog):
    def __init__(self, bookmark_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("书签管理器"))
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        self.setFixedSize(600, 500)

    def move_item_up(self):
        item = self.tree.currentItem()
        if not item:
            show_toast(self, tr("未选择"), tr("请先选择要上移的书签或文件夹。"), level="warning")
            return
        parent = item.parent()
        if parent:
            siblings = [parent.child(i) for i in range(parent.childCount())]
        else:
            siblings = [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())]
        idx = siblings.index(item)
        if idx <= 0:
            return  # 已经在最上面
        # 交换UI顺序
        if parent:
            parent.removeChild(item)
            parent.insertChild(idx-1, item)
        else:
            self.tree.takeTopLevelItem(idx)
            self.tree.insertTopLevelItem(idx-1, item)
        node_id = item.data(0, 1)
        self.update_bookmark_order(item, -1)
        # 重新选中移动后的项目
        self.reselect_item_by_id(node_id)
        # 刷新主界面书签栏
        self.refresh_main_window_bookmark_bar()

    def move_item_down(self):
        item = self.tree.currentItem()
        if not item:
            show_toast(self, tr("未选择"), tr("请先选择要下移的书签或文件夹。"), level="warning")
            return
        parent = item.parent()
        if parent:
            siblings = [parent.child(i) for i in range(parent.childCount())]
        else:
            siblings = [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())]
        idx = siblings.index(item)
        if idx >= len(siblings) - 1:
            return  # 已经在最下面
        # 交换UI顺序
        if parent:
            parent.removeChild(item)
            parent.insertChild(idx+1, item)
        else:
            self.tree.takeTopLevelItem(idx)
            self.tree.insertTopLevelItem(idx+1, item)
        node_id = item.data(0, 1)
        self.update_bookmark_order(item, 1)
        # 重新选中移动后的项目
        self.reselect_item_by_id(node_id)
        # 刷新主界面书签栏
        self.refresh_main_window_bookmark_bar()
    def reselect_item_by_id(self, node_id):
        # 遍历tree，找到id为node_id的item并选中
        def find_item(item):
            if item.data(0, 1) == node_id:
                return item
            for i in range(item.childCount()):
                found = find_item(item.child(i))
                if found:
                    return found
            return None
        root_count = self.tree.topLevelItemCount()
        for i in range(root_count):
            item = self.tree.topLevelItem(i)
            found = find_item(item)
            if found:
                self.tree.setCurrentItem(found)
                break

    def refresh_main_window_bookmark_bar(self):
        main_window = self.parent() if self.parent() and hasattr(self.parent(), 'populate_bookmark_bar_menu') else None
        if main_window:
            main_window.populate_bookmark_bar_menu()

    def update_bookmark_order(self, item, direction):
        # direction: -1=up, 1=down
        node_id = item.data(0, 1)
        def reorder_children(children):
            idx = None
            for i, node in enumerate(children):
                if node.get('id') == node_id:
                    idx = i
                    break
            if idx is not None:
                new_idx = idx + direction
                if 0 <= new_idx < len(children):
                    children[idx], children[new_idx] = children[new_idx], children[idx]
                    return True
            return False
        def recursive_reorder(node):
            if isinstance(node, dict) and 'children' in node:
                if reorder_children(node['children']):
                    return True
                for child in node['children']:
                    if recursive_reorder(child):
                        return True
            elif isinstance(node, list):
                if reorder_children(node):
                    return True
                for child in node:
                    if recursive_reorder(child):
                        return True
            return False
        tree = self.bookmark_manager.get_tree()
        recursive_reorder(tree.get('bookmark_bar'))
        self.bookmark_manager.save_bookmarks()
        self.populate_tree()
    
    def on_items_moved(self):
        """拖拽完成后重建书签数据结构"""
        try:
            debug_print("[BookmarkDrag] Starting to rebuild structure after drag")
            
            # 从树形控件重建书签结构
            new_structure = self._rebuild_bookmark_structure()
            
            debug_print(f"[BookmarkDrag] Rebuilt {len(new_structure)} top-level items")
            
            # 更新书签管理器
            tree = self.bookmark_manager.get_tree()
            if 'bookmark_bar' in tree:
                tree['bookmark_bar']['children'] = new_structure
                self.bookmark_manager.save_bookmarks()
                
                # 刷新主窗口书签栏
                self.refresh_main_window_bookmark_bar()
                
                debug_print("[BookmarkDrag] Bookmark structure updated and saved")
                show_toast(self, tr("已保存"), tr("书签已保存"), level="success")
        except Exception as e:
            debug_print(f"[BookmarkDrag] Error updating structure: {e}")
            import traceback
            traceback.print_exc()
            show_toast(self, tr("保存失败"), tr("拖拽保存失败: {}").format(e), level="error")
    
    def _rebuild_bookmark_structure(self):
        """从树形控件重建书签数据结构"""
        # 首先获取原始数据，以便保留date_added等字段
        original_tree = self.bookmark_manager.get_tree()
        original_nodes = {}
        
        def collect_original_nodes(node):
            if isinstance(node, dict):
                node_id = node.get('id')
                if node_id:
                    original_nodes[node_id] = node
                if 'children' in node:
                    for child in node['children']:
                        collect_original_nodes(child)
        
        if 'bookmark_bar' in original_tree:
            collect_original_nodes(original_tree['bookmark_bar'])
        
        def process_item(item):
            node_id = item.data(0, 1)
            node_type = item.text(1)
            name = item.text(0).lstrip("📁 ").lstrip("📑 ")
            
            # 尝试从原始数据中获取节点
            original = original_nodes.get(node_id, {})
            
            if node_type == tr('文件夹'):
                node = {
                    'id': node_id,
                    'name': name,
                    'type': 'folder',
                    'date_added': original.get('date_added', node_id),
                    'children': []
                }
                # 递归处理子项
                for i in range(item.childCount()):
                    child = item.child(i)
                    child_node = process_item(child)
                    if child_node:
                        node['children'].append(child_node)
                return node
            elif node_type == tr('书签'):
                url = item.text(2)
                return {
                    'id': node_id,
                    'name': name,
                    'type': 'url',
                    'url': url,
                    'date_added': original.get('date_added', node_id)
                }
            return None
        
        # 处理所有顶层项
        result = []
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            node = process_item(item)
            if node:
                result.append(node)
        
        return result
    
    def __init__(self, bookmark_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("书签管理"))
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.resize(600, 500)
        self.bookmark_manager = bookmark_manager
        layout = QVBoxLayout(self)
        
        # 使用标准树形控件（拖拽不自动保存）
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels([tr("名称"), tr("类型"), tr("路径")])
        self.tree.setColumnWidth(0, 250)  # 第一列宽一些
        
        # 启用拖拽
        self.tree.setDragEnabled(True)
        self.tree.setAcceptDrops(True)
        self.tree.setDropIndicatorShown(True)
        self.tree.setDragDropMode(QTreeWidget.InternalMove)
        self.tree.setSelectionMode(QTreeWidget.SingleSelection)
        
        layout.addWidget(self.tree)
        
        # 添加拖拽提示
        drag_hint = QLabel(tr("💡 提示：可以拖动书签和文件夹调整顺序和层级，调整后点击【保存】按钮保存更改"))
        drag_hint.setStyleSheet("QLabel { color: #666; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }")
        layout.addWidget(drag_hint)
        
        self.populate_tree()

        btn_layout = QHBoxLayout()
        self.edit_btn = QPushButton(tr("编辑"))
        self.edit_btn.clicked.connect(self.edit_item)
        btn_layout.addWidget(self.edit_btn)
        self.delete_btn = QPushButton(tr("删除"))
        self.delete_btn.clicked.connect(self.delete_item)
        btn_layout.addWidget(self.delete_btn)
        self.new_folder_btn = QPushButton(tr("新建文件夹"))
        self.new_folder_btn.clicked.connect(self.create_folder)
        btn_layout.addWidget(self.new_folder_btn)
        self.up_btn = QPushButton(tr("上移"))
        self.up_btn.clicked.connect(self.move_item_up)
        btn_layout.addWidget(self.up_btn)
        self.down_btn = QPushButton(tr("下移"))
        self.down_btn.clicked.connect(self.move_item_down)
        btn_layout.addWidget(self.down_btn)
        
        # 添加导入/导出按钮
        self.export_btn = QPushButton(tr("📤 导出"))
        self.export_btn.setToolTip(tr("导出书签到JSON文件"))
        self.export_btn.clicked.connect(self.export_bookmarks)
        btn_layout.addWidget(self.export_btn)
        
        self.import_btn = QPushButton(tr("📥 导入"))
        self.import_btn.setToolTip(tr("从JSON文件导入书签"))
        self.import_btn.clicked.connect(self.import_bookmarks)
        btn_layout.addWidget(self.import_btn)
        
        # 添加手动保存按钮
        self.save_btn = QPushButton(tr("💾 保存"))
        self.save_btn.setToolTip(tr("保存当前书签顺序和层级"))
        self.save_btn.clicked.connect(self.manual_save)
        btn_layout.addWidget(self.save_btn)
        
        close_btn = QPushButton(tr("关闭"))
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
    
    def manual_save(self):
        """手动保存书签"""
        try:
            self.on_items_moved()
        except Exception as e:
            show_toast(self, tr("保存失败"), tr("保存失败: {}").format(e), level="error")
    
    def edit_item(self):
        item = self.tree.currentItem()
        if not item:
            show_toast(self, tr("未选择"), tr("请先选择要编辑的书签或文件夹。"), level="warning")
            return
        node_type = item.text(1)
        old_name = item.text(0).lstrip("📁 ").lstrip("📑 ")
        main_window = self.parent() if self.parent() and hasattr(self.parent(), 'populate_bookmark_bar_menu') else None
        if node_type == tr('文件夹'):
            new_name, ok = QInputDialog.getText(self, tr("编辑文件夹"), tr("请输入新名称："), text=old_name)
            if ok and new_name and new_name != old_name:
                item.setText(0, f"📁 {new_name}")
                self.update_name_in_bookmark_manager(item, new_name)
                self.bookmark_manager.save_bookmarks()
                self.populate_tree()
                if main_window:
                    main_window.populate_bookmark_bar_menu()
        elif node_type == tr('书签'):
            new_name, ok1 = QInputDialog.getText(self, tr("编辑书签"), tr("请输入新名称："), text=old_name)
            old_url = item.text(2)
            new_url, ok2 = QInputDialog.getText(self, tr("编辑书签"), tr("请输入新路径："), text=old_url)
            if ok1 and new_name and (new_name != old_name or new_url != old_url) and ok2 and new_url:
                item.setText(0, f"📑 {new_name}")
                item.setText(2, new_url)
                self.update_bookmark_in_manager(item, new_name, new_url)
                self.bookmark_manager.save_bookmarks()
                self.populate_tree()
                if main_window:
                    main_window.populate_bookmark_bar_menu()

    def update_bookmark_in_manager(self, item, new_name, new_url):
        node_id = item.data(0, 1)
        def update_node(node):
            if isinstance(node, dict):
                if node.get('id') == node_id:
                    node['name'] = new_name
                    node['url'] = new_url
                    return True
                if 'children' in node:
                    for child in node['children']:
                        if update_node(child):
                            return True
            elif isinstance(node, list):
                for child in node:
                    if update_node(child):
                        return True
            return False
        tree = self.bookmark_manager.get_tree()
        update_node(tree.get('bookmark_bar'))

    def delete_item(self):
        item = self.tree.currentItem()
        if not item:
            show_toast(self, tr("未选择"), tr("请先选择要删除的书签或文件夹。"), level="warning")
            return
        node_id = item.data(0, 1)
        # 直接执行删除并给出提示，避免阻塞
        show_toast(self, tr("已删除"), tr("选中的书签/文件夹已删除"), level="info")
        def delete_node(parent, node_list):
            for i, node in enumerate(node_list):
                if isinstance(node, dict) and node.get('id') == node_id:
                    del node_list[i]
                    return True
                if isinstance(node, dict) and 'children' in node:
                    if delete_node(node, node['children']):
                        return True
            return False
        tree = self.bookmark_manager.get_tree()
        bookmark_bar = tree.get('bookmark_bar')
        if bookmark_bar and 'children' in bookmark_bar:
            delete_node(bookmark_bar, bookmark_bar['children'])
            self.bookmark_manager.save_bookmarks()
            self.populate_tree()
            main_window = self.parent() if self.parent() and hasattr(self.parent(), 'populate_bookmark_bar_menu') else None
            if main_window:
                main_window.populate_bookmark_bar_menu()

    def populate_tree(self):
        self.tree.clear()
        tree = self.bookmark_manager.get_tree()
        bookmark_bar = tree.get('bookmark_bar')
        if not bookmark_bar or 'children' not in bookmark_bar:
            return
        def add_node(parent_item, node):
            if node.get('type') == 'folder':
                item = QTreeWidgetItem([f"📁 {node.get('name', '')}", tr('文件夹'), ''])
                item.setData(0, 1, node.get('id'))
                if parent_item:
                    parent_item.addChild(item)
                else:
                    self.tree.addTopLevelItem(item)
                for child in node.get('children', []):
                    add_node(item, child)
            elif node.get('type') == 'url':
                item = QTreeWidgetItem([f"📑 {node.get('name', '')}", tr('书签'), node.get('url', '')])
                item.setData(0, 1, node.get('id'))
                if parent_item:
                    parent_item.addChild(item)
                else:
                    self.tree.addTopLevelItem(item)
        for child in bookmark_bar['children']:
            add_node(None, child)
        self.tree.expandAll()

    def rename_item(self):
        item = self.tree.currentItem()
        if not item:
            show_toast(self, tr("未选择"), tr("请先选择要重命名的书签或文件夹。"), level="warning")
            return
        old_name = item.text(0)
        new_name, ok = QInputDialog.getText(self, tr("重命名"), tr("请输入新名称："), text=old_name)
        if ok and new_name and new_name != old_name:
            item.setText(0, new_name)
            # 实际数据同步
            self.update_name_in_bookmark_manager(item, new_name)
            self.bookmark_manager.save_bookmarks()

    def update_name_in_bookmark_manager(self, item, new_name):
        # 递归查找并更新id对应的节点名称
        node_id = item.data(0, 1)
        def update_name(node):
            if isinstance(node, dict):
                if node.get('id') == node_id:
                    node['name'] = new_name
                    return True
                if 'children' in node:
                    for child in node['children']:
                        if update_name(child):
                            return True
            elif isinstance(node, list):
                for child in node:
                    if update_name(child):
                        return True
            return False
        tree = self.bookmark_manager.get_tree()
        update_name(tree.get('bookmark_bar'))

    def create_folder(self):
        item = self.tree.currentItem()
        parent_id = None
        if item and item.text(1) == tr('文件夹'):
            parent_id = item.data(0, 1)
        else:
            # 默认加到bookmark_bar根
            parent_id = self.bookmark_manager.get_tree().get('bookmark_bar', {}).get('id')
        folder_name, ok = QInputDialog.getText(self, tr("新建文件夹"), tr("请输入文件夹名称："))
        if ok and folder_name:
            import time
            new_id = str(int(time.time() * 1000000))
            folder = {
                "date_added": new_id,
                "id": new_id,
                "name": folder_name,
                "type": "folder",
                "children": []
            }
            # 插入到父节点
            def insert_folder(node):
                if isinstance(node, dict):
                    if node.get('id') == parent_id:
                        node.setdefault('children', []).append(folder)
                        return True
                    if 'children' in node:
                        for child in node['children']:
                            if insert_folder(child):
                                return True
                elif isinstance(node, list):
                    for child in node:
                        if insert_folder(child):
                            return True
                return False
            tree = self.bookmark_manager.get_tree()
            if not insert_folder(tree.get('bookmark_bar')):
                # 根节点
                tree.get('bookmark_bar', {}).setdefault('children', []).append(folder)
            self.bookmark_manager.save_bookmarks()
            self.populate_tree()

    def export_bookmarks(self):
        """导出书签到JSON文件"""
        from PyQt5.QtWidgets import QFileDialog
        import shutil
        from datetime import datetime
        
        # 生成默认文件名（包含日期时间）
        default_name = f"bookmarks_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        # 打开保存文件对话框
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            tr("导出书签"),
            default_name,
            "JSON Files (*.json);;All Files (*)"
        )
        
        if file_path:
            try:
                # 复制当前的bookmarks.json到目标位置
                bookmarks_path = get_app_data_path("bookmarks.json")
                shutil.copy2(bookmarks_path, file_path)
                show_toast(self, tr("导出成功"), tr("书签已成功导出到:\n{}").format(file_path), level="success")
                print(f"[Bookmark Export] Successfully exported to: {file_path}")
            except Exception as e:
                show_toast(self, tr("导出失败"), f"导出书签时出错:\n{str(e)}", level="error")
                print(f"[Bookmark Export] Error: {e}")
    
    def import_bookmarks(self):
        """从JSON文件导入书签"""
        from PyQt5.QtWidgets import QFileDialog
        import json
        
        # 打开文件选择对话框
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            tr("导入书签"),
            "",
            "JSON Files (*.json);;All Files (*)"
        )
        
        if not file_path:
            return
        
        try:
            # 读取导入的JSON文件
            with open(file_path, 'r', encoding='utf-8') as f:
                imported_data = json.load(f)
            
            # 处理可能有 roots 层的书签格式（兼容Chrome书签格式）
            if 'roots' in imported_data:
                imported_data = imported_data['roots']
            
            # 验证JSON格式
            if not isinstance(imported_data, dict) or 'bookmark_bar' not in imported_data:
                show_toast(self, tr("格式错误"), tr("导入的文件格式不正确，必须包含 'bookmark_bar' 节点"), level="warning")
                return
            
            # 默认选择更安全的“合并”模式，避免阻塞确认
            show_toast(self, tr("导入方式"), tr("已自动选择合并模式，将导入内容追加到现有书签。"), level="info")
            # 合并模式：将导入的书签添加到现有书签的末尾
            current_tree = self.bookmark_manager.get_tree()
            imported_bar = imported_data.get('bookmark_bar', {})
            imported_children = imported_bar.get('children', [])
            
            if imported_children:
                current_bar = current_tree.get('bookmark_bar', {})
                if 'children' not in current_bar:
                    current_bar['children'] = []
                
                # 添加到末尾
                current_bar['children'].extend(imported_children)
                self.bookmark_manager.save_bookmarks(immediate=True)  # 立即保存
                
                count = len(imported_children)
                show_toast(self, tr("导入成功"), tr("成功导入 {} 个书签项").format(count), level="success")
                print(f"[Bookmark Import] Merged {count} items from: {file_path}")
            else:
                show_toast(self, tr("提示"), tr("导入的文件中没有书签内容"), level="info")
            
            # 刷新书签管理对话框显示
            self.populate_tree()
            
            # 刷新主窗口书签栏
            main_window = self.parent()
            if main_window and hasattr(main_window, 'populate_bookmark_bar_menu'):
                print("[Bookmark Import] Refreshing main window bookmark bar")
                main_window.populate_bookmark_bar_menu()
                # 确保默认图标显示
                if hasattr(main_window, 'ensure_default_icons_on_bookmark_bar'):
                    main_window.ensure_default_icons_on_bookmark_bar()
                
        except json.JSONDecodeError:
            show_toast(self, tr("格式错误"), tr("导入的文件不是有效的JSON格式"), level="error")
            print(f"[Bookmark Import] Invalid JSON format: {file_path}")
        except Exception as e:
            show_toast(self, tr("导入失败"), f"导入书签时出错:\n{str(e)}", level="error")
            print(f"[Bookmark Import] Error: {e}")


    # BookmarkManagerDialog不再包含标签页相关方法

def try_send_to_existing_instance(path):
    """尝试将路径发送给已运行的实例"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client.settimeout(2.0)  # 增加超时时间
            client.connect(('127.0.0.1', 58923))
            client.send(path.encode('utf-8'))
            client.close()
            debug_print(f"[Client] Successfully sent path to existing instance: {path}")
            return True
        except Exception as e:
            debug_print(f"[Client] Attempt {attempt + 1}/{max_retries} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(0.1)  # 短暂等待后重试
            continue
    debug_print("[Client] No existing instance found, starting new instance")
    return False

def main():
    # 支持命令行参数：打开指定路径
    import sys
    import os

    # TortoiseOverlays IsMemberOf vtable patch is applied at module load time
    # (see early init block near _patch_tortoise_overlays). No registry changes needed.

    path_to_open = None
    if len(sys.argv) > 1:
        path = sys.argv[1]
        # 处理可能的引号
        path = path.strip('"').strip("'")
        if os.path.exists(path):
            # 如果是文件，打开其所在目录
            if os.path.isfile(path):
                path = os.path.dirname(path)
            path_to_open = path
            
            # 尝试发送给已运行的实例
            if try_send_to_existing_instance(path):
                print(f"Sent path to existing instance: {path}")
                sys.exit(0)  # 退出程序，不启动新实例
    
    # 禁用 Qt 的警告输出（在创建 QApplication 之前设置）
    os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.qpa.*=false'
    
    # 启用高DPI支持（在创建QApplication之前）
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    # 启动新实例
    app = QApplication(sys.argv)
    app.setApplicationName("TabExplorer")
    
    # 获取屏幕DPI和缩放因子
    screen = app.primaryScreen()
    dpi = screen.logicalDotsPerInch()
    scale_factor = dpi / 96.0  # 96是标准DPI
    debug_print(f"[DPI] Screen DPI: {dpi}, Scale Factor: {scale_factor:.2f}")
    
    # 根据DPI动态调整全局样式
    base_font_size = 9  # 基础字体大小 (pt)
    scaled_font_size = int(base_font_size * scale_factor)
    
    # 设置全局字体
    from PyQt5.QtGui import QFont
    app_font = QFont("Microsoft YaHei UI", scaled_font_size)
    app.setFont(app_font)
    debug_print(f"[DPI] Global font size: {scaled_font_size}pt")
    
    # 安装自定义 Qt 消息处理器，过滤 QAxBase 等警告
    from PyQt5.QtCore import qInstallMessageHandler
    qInstallMessageHandler(qt_message_handler)

    # 安装 IExplorerBrowser 键盘消息过滤器
    app.installNativeEventFilter(_ieb_keyboard_filter)
    
    # 图标将在 MainWindow.__init__ 中生成并设置，确保 Qt 完全初始化后执行
    
    # 创建窗口（图标在 MainWindow.__init__ 内部生成）
    window = MainWindow()
    
    
    # 如果有路径参数，在新窗口中打开
    # 注意：固定标签页在延迟初始化阶段加载，这里要避免和固定标签重复
    if path_to_open:
        pinned_norm = {
            window._normalize_path_for_compare(p)
            for p in window.config.get("pinned_tabs", []) if p
        }
        if window._normalize_path_for_compare(path_to_open) in pinned_norm:
            debug_print(f"[App] Skip argv path (already pinned): {path_to_open}")
        else:
            window.add_new_tab(path_to_open)
    
    # 启动时最大化显示
    window.showMaximized()
    

    # No HKCU cleanup needed – overlay fix is vtable-only (process-local).

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()


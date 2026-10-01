"""设置窗口。"""

import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QCheckBox, QDialog, QLabel, QPushButton, QVBoxLayout

from . import theme as _theme
from .paths import get_app_base_dir
from .i18n import _set_app_language, tr
from .debuglog import set_debug_mode, set_explorer_monitor_debug
from .system import normalize_terminal_tool_name
from .hotkeys import (
    _format_hotkey, _hotkey_bindings, HOTKEY_COMMANDS, _parse_hotkey, _validate_hotkey_bindings,
)
from .widgets import show_toast


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
        self.bottom_statusbar_cb = QCheckBox(tr("显示底部状态栏（状态/CPU信息）"), self)
        self.bottom_statusbar_cb.setChecked(config.get("show_bottom_statusbar", True))
        self.bottom_statusbar_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.bottom_statusbar_cb.setToolTip(tr("显示或隐藏每个标签页底部的状态栏（含状态文本、取消按钮和CPU/内存信息）"))
        monitor_layout.addWidget(self.bottom_statusbar_cb)
        # 监听间隔设置
        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel(tr("监听间隔（秒）:")))
        from PyQt5.QtWidgets import QDoubleSpinBox
        self.interval_spinbox = QDoubleSpinBox()
        self.interval_spinbox.setRange(0.5, 10.0)
        self.interval_spinbox.setSingleStep(0.5)
        self.interval_spinbox.setValue(config.get("explorer_monitor_interval", 2.0))
        self.interval_spinbox.setToolTip(tr("系统窗口事件不可用时的轮询间隔；事件可用时新窗口即时嵌入，并至少每 10 秒兜底扫描一次"))
        interval_layout.addWidget(self.interval_spinbox)
        interval_layout.addWidget(QLabel(tr("（推荐: 2.0秒）")))
        interval_layout.addStretch(1)
        monitor_layout.addLayout(interval_layout)
        hibernate_layout = QHBoxLayout()
        hibernate_layout.addWidget(QLabel(tr("后台标签休眠（分钟，0=关闭）:")))
        from PyQt5.QtWidgets import QSpinBox
        self.tab_hibernate_spin = QSpinBox()
        self.tab_hibernate_spin.setRange(0, 1440)
        self.tab_hibernate_spin.setValue(int(config.get("tab_hibernate_minutes", 30) or 0))
        self.tab_hibernate_spin.setToolTip(tr("后台标签超过该时间未访问时释放文件视图，切回时重新加载（滚动位置和选中项不保留）"))
        hibernate_layout.addWidget(self.tab_hibernate_spin)
        hibernate_layout.addStretch(1)
        monitor_layout.addLayout(hibernate_layout)
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

        file_op_workers_layout = QHBoxLayout()
        file_op_workers_layout.addWidget(QLabel(tr("文件操作并发数（0=自动）:")))
        self.file_op_workers_spin = QSpinBox(self)
        self.file_op_workers_spin.setRange(0, 16)
        self.file_op_workers_spin.setSingleStep(1)
        self.file_op_workers_spin.setSpecialValueText(tr("自动"))
        self.file_op_workers_spin.setValue(int(config.get("file_op_max_workers", 0) or 0))
        self.file_op_workers_spin.setToolTip(tr("后台复制/删除的并发文件任务数。0=自动，建议机械盘 2-4，SSD 4-8"))
        file_op_workers_layout.addWidget(self.file_op_workers_spin)
        file_op_workers_layout.addStretch(1)
        debug_layout.addLayout(file_op_workers_layout)
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
        self.show_tab_group_markers_cb = QCheckBox(tr("显示标签分组标记（颜色）"), self)
        self.show_tab_group_markers_cb.setChecked(config.get("show_tab_group_markers", True))
        self.show_tab_group_markers_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.show_tab_group_markers_cb.setToolTip(
            tr("在标签页上显示分组颜色，关闭后仅保留分组逻辑不显示颜色")
        )
        tabs_layout.addWidget(self.show_tab_group_markers_cb)
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
        _theme.bind_style(
            gesture_help,
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
        hotkey_group = self._create_hotkey_group(config, compact_widget)
        compact_groupbox(hotkey_group)
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
        _theme.bind_style(
            self.ai_key_tip_label,
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
        _theme.bind_style(
            eye_btn,
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
        self.ai_system_prompt_edit.setPlaceholderText(tr("留空则使用内置提示词（支持文件、目录、脚本与 Git 指令）"))
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
        lang_row.addSpacing(16)
        lang_row.addWidget(QLabel(tr("主题:")))
        self.theme_combo = QComboBox(self)
        self.theme_combo.addItem(tr("跟随系统"), "system")
        self.theme_combo.addItem(tr("浅色"), "light")
        self.theme_combo.addItem(tr("深色"), "dark")
        self.theme_combo.setCurrentIndex(
            max(0, self.theme_combo.findData(_theme.normalize_mode(config.get("theme", "system")))))
        self.theme_combo.setToolTip(tr("跟随系统：随 Windows “应用模式”的浅色/深色设置自动切换"))
        lang_row.addWidget(self.theme_combo)
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
        update_row = QHBoxLayout()
        update_row.addWidget(update_link, 1)
        self.check_update_button = QPushButton(tr("立即检查"), self)
        self.check_update_button.setToolTip(tr("查询 GitHub 上的最新发布版本，只提示不下载"))
        self.check_update_button.clicked.connect(self._check_updates_now)
        update_row.addWidget(self.check_update_button)
        bottom_layout.addLayout(update_row)
        self.auto_update_cb = QCheckBox(tr("自动检查新版本（每天最多一次，只提示不下载）"), self)
        self.auto_update_cb.setChecked(bool(config.get("auto_update_check", False)))
        bottom_layout.addWidget(self.auto_update_cb)
        
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

    def _create_hotkey_group(self, config, compact_widget):
        """每个命令一行：启用开关 + 按键录制框 + 清空按钮；共用开关的命令联动勾选。"""
        from PyQt5.QtWidgets import QGroupBox, QGridLayout, QKeySequenceEdit, QToolButton, QScrollArea, QWidget
        from PyQt5.QtGui import QKeySequence
        group = QGroupBox(tr("快捷键设置"))
        group_layout = QVBoxLayout(group)
        rows_widget = QWidget()
        grid = QGridLayout(rows_widget)
        grid.setContentsMargins(0, 0, 4, 0)
        grid.setColumnStretch(1, 1)
        hotkeys = config.get("hotkeys", {})
        bindings = _hotkey_bindings(config)
        self.hotkey_edits = {}
        switches = {}

        def keep_first_chord(edit, sequence):
            if sequence.count() > 1:
                edit.setKeySequence(QKeySequence(sequence[0]))

        for row, (command, enable_key, _default, label) in enumerate(HOTKEY_COMMANDS):
            switch = QCheckBox(tr(label))
            if enable_key is None:
                switch.setChecked(True)
                switch.setEnabled(False)
                switch.setToolTip(tr("始终启用，清空按键即可停用"))
            elif enable_key in switches:
                partner = switches[enable_key]
                switch.setChecked(partner.isChecked())
                switch.toggled.connect(partner.setChecked)
                partner.toggled.connect(switch.setChecked)
            else:
                switch.setChecked(hotkeys.get(enable_key, True))
                switches[enable_key] = switch
                setattr(self, 'hotkey_' + enable_key, switch)
            edit = QKeySequenceEdit(QKeySequence(bindings[command], QKeySequence.PortableText), rows_widget)
            edit.setAccessibleName(tr(label))
            edit.keySequenceChanged.connect(lambda sequence, target=edit: keep_first_chord(target, sequence))
            clear = QToolButton(rows_widget)
            clear.setText('×')
            clear.setToolTip(tr("清空按键"))
            clear.setAccessibleName(tr("清空按键"))
            clear.clicked.connect(edit.clear)
            for widget in (switch, edit):
                compact_widget(widget)
            grid.addWidget(switch, row, 0)
            grid.addWidget(edit, row, 1)
            grid.addWidget(clear, row, 2)
            self.hotkey_edits[command] = edit
        row = len(HOTKEY_COMMANDS)
        self.hotkey_switch_tab_number = QCheckBox(tr("切换到第 N 个标签（Ctrl+9 为最后一个）"))
        self.hotkey_switch_tab_number.setChecked(hotkeys.get("switch_tab_number", True))
        number_keys = QLabel("Ctrl+1…9")
        for widget in (self.hotkey_switch_tab_number, number_keys):
            compact_widget(widget)
        grid.addWidget(self.hotkey_switch_tab_number, row, 0)
        grid.addWidget(number_keys, row, 1)
        # 行数较多，放进滚动区，避免设置窗口在 1080p 屏幕上超高
        scroll = QScrollArea(group)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setWidget(rows_widget)
        row_height = max(edit.sizeHint().height(), switch.sizeHint().height()) + max(grid.verticalSpacing(), 0)
        scroll.setFixedHeight(row_height * 9)
        group_layout.addWidget(scroll)
        tip_label = QLabel(tr("💡 点击按键框后直接按下新组合；需包含 Ctrl 或 Alt（F1–F24 可单独使用）。取消勾选或清空按键可停用。"))
        tip_label.setWordWrap(True)
        _theme.bind_style(
            tip_label,
            "QLabel { color: #666; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }"
        )
        group_layout.addWidget(tip_label)
        reset_button = QPushButton(tr("恢复默认按键"), group)
        reset_button.clicked.connect(self._reset_hotkey_bindings)
        group_layout.addWidget(reset_button, 0, Qt.AlignRight)
        self._hotkey_group = group
        return group

    def _reset_hotkey_bindings(self):
        from PyQt5.QtGui import QKeySequence
        for command, _enable_key, default, _label in HOTKEY_COMMANDS:
            self.hotkey_edits[command].setKeySequence(QKeySequence(default, QKeySequence.PortableText))

    def _collect_hotkey_bindings(self):
        """读取按键框，返回 (规范化后的绑定, 错误说明列表)。"""
        from PyQt5.QtGui import QKeySequence
        bindings = {}
        for command, edit in self.hotkey_edits.items():
            text = edit.keySequence().toString(QKeySequence.PortableText).split(', ')[0]
            parsed = _parse_hotkey(text)
            bindings[command] = _format_hotkey(*parsed) if parsed else text
        return bindings, _validate_hotkey_bindings(bindings, self.hotkey_switch_tab_number.isChecked())

    def _check_updates_now(self):
        owner = self.parent()
        if owner is not None and hasattr(owner, 'check_for_updates'):
            owner.check_for_updates(manual=True)

    def accept(self):
        """保存设置"""
        hotkey_bindings, hotkey_errors = self._collect_hotkey_bindings()
        if hotkey_errors:
            from PyQt5.QtWidgets import QMessageBox
            self.settings_tabs.setCurrentWidget(self._hotkey_group.parentWidget())
            QMessageBox.warning(self, tr("快捷键冲突"), '\n'.join(hotkey_errors))
            return
        # 保存所有设置到 parent (MainWindow)
        if self.parent():
            # 处理开机启动设置
            auto_startup_enabled = self.auto_startup_cb.isChecked()
            current_enabled = self._is_auto_startup_enabled()
            if auto_startup_enabled != current_enabled:
                self._set_auto_startup(auto_startup_enabled)
            
            self.parent().config["enable_explorer_monitor"] = self.monitor_cb.isChecked()
            self.parent().config["explorer_monitor_interval"] = self.interval_spinbox.value()
            self.parent().config["tab_hibernate_minutes"] = self.tab_hibernate_spin.value()
            self.parent().config["debug_mode"] = self.debug_mode_cb.isChecked()
            self.parent().config["explorer_monitor_debug"] = self.explorer_monitor_debug_cb.isChecked()
            self.parent().config["file_op_max_workers"] = self.file_op_workers_spin.value()
            self.parent().config["show_bottom_statusbar"] = self.bottom_statusbar_cb.isChecked()
            self.parent().config["show_resource_usage_in_statusbar"] = self.resource_usage_cb.isChecked()
            self.parent().config["show_tab_group_markers"] = self.show_tab_group_markers_cb.isChecked()
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
                "switch_tab_number": self.hotkey_switch_tab_number.isChecked(),
                "search": self.hotkey_search.isChecked(),
                "quick_find_current_dir": self.hotkey_quick_find_current_dir.isChecked(),
                "navigate": self.hotkey_navigate.isChecked(),
                "go_up": self.hotkey_go_up.isChecked(),
                "refresh": self.hotkey_refresh.isChecked(),
                "add_bookmark": self.hotkey_add_bookmark.isChecked(),
                "quick_copy": self.hotkey_quick_copy.isChecked(),
                "quick_paste": self.hotkey_quick_paste.isChecked(),
                "quick_delete": self.hotkey_quick_delete.isChecked(),
                "cancel_file_op": self.hotkey_cancel_file_op.isChecked(),
                "copy_filename": self.hotkey_copy_filename.isChecked(),
                "copy_filepath": self.hotkey_copy_filepath.isChecked()
            }
            self.parent().config["hotkey_bindings"] = hotkey_bindings
            self.parent().config["auto_update_check"] = self.auto_update_cb.isChecked()
            self.parent().config["theme"] = self.theme_combo.currentData()
            
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
            if hasattr(self.parent(), 'apply_bottom_statusbar_config'):
                self.parent().apply_bottom_statusbar_config()
            if hasattr(self.parent(), 'apply_resource_usage_config'):
                self.parent().apply_resource_usage_config()
            if hasattr(self.parent(), 'apply_tab_group_markers_config'):
                self.parent().apply_tab_group_markers_config()
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
                if not ai_enabled and getattr(self.parent(), 'chat_panel', None) is not None:
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

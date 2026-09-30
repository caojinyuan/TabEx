"""Configuration storage and session capture/restore ownership."""

import json
import os
import time
from . import search as _search
from .paths import get_app_data_path
from .i18n import _set_app_language
from .debuglog import debug_print
from .search import apply_runtime_performance_config
from .bookmarks import _preserve_unreadable_file

class ConfigStore:
    def __init__(self, path):
        self.path = os.fspath(path)
        self.last_written = None
        self.recovered_backup = None
        self.write_blocked = False
        self.last_error = ''

    @staticmethod
    def _merge_defaults(config, defaults):
        for key, default in defaults.items():
            value = config.get(key)
            if not isinstance(value, type(default)) or (isinstance(value, bool) and not isinstance(default, bool)):
                config[key] = default
            elif isinstance(default, dict):
                ConfigStore._merge_defaults(value, default)

    def load(self):
        """加载配置文件"""
        default_config = {
            "enable_explorer_monitor": True,  # 默认启用Explorer监听
            "debug_mode": False,  # 默认关闭调试输出
            "explorer_monitor_debug": False,  # 默认关闭Explorer Monitor调试输出
            "file_op_max_workers": 0,  # 后台文件操作并发数：0=自动
            "show_bottom_statusbar": True,  # 默认显示底部状态栏（状态/CPU信息区域）
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
            "show_tab_group_markers": True,  # 默认显示标签分组颜色
            "tab_hibernate_minutes": 30,  # 后台标签闲置多久后释放 Shell 视图：0=关闭
            "hotkey_bindings": {},  # 自定义按键 {命令: "Ctrl+T"}；未列出的命令使用默认按键
            "auto_update_check": False,  # 默认关闭；开启后每天最多查询一次 GitHub 最新发布，只提示不下载
            "last_update_check": 0,
            "update_notified_version": "",
            # 快捷键配置
            "hotkeys": {
                "new_tab": True,           # Ctrl+T
                "close_tab": True,         # Ctrl+W
                "reopen_tab": True,        # Ctrl+Shift+T
                "switch_tab": True,        # Ctrl+Tab / Ctrl+Shift+Tab
                "switch_tab_number": True, # Ctrl+1..9
                "search": True,            # Ctrl+F
                "quick_find_current_dir": True,  # Ctrl+G
                "navigate": True,          # Alt+Left/Right
                "go_up": True,             # Alt+Up
                "refresh": True,           # F5
                "add_bookmark": True,      # Ctrl+D
                "quick_copy": True,        # Alt+C - 快速复制选中项
                "quick_paste": True,       # Alt+V - 快速粘贴到当前目录
                "quick_delete": True,      # Alt+Delete - 快速删除
                "cancel_file_op": True,    # Alt+Q - 取消后台复制/删除
                "copy_filename": True,     # Alt+Z - 复制选中文件名
                "copy_filepath": True,     # Alt+X - 复制文件路径\文件名
                "split_view": True,        # F3 - 左右分屏对比
                "insert_group_bookmark": True  # F4 - 插入标签分组（保留旧键名兼容已有配置）
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
            "content_search_chunk_size": _search.CONTENT_SEARCH_CHUNK_SIZE,
            "content_search_max_bytes_per_file": _search.CONTENT_SEARCH_MAX_BYTES_PER_FILE,
            "content_search_in_memory_threshold": _search.CONTENT_SEARCH_IN_MEMORY_THRESHOLD,
            "search_result_queue_maxsize": _search.SEARCH_RESULT_QUEUE_MAXSIZE,
            "search_result_batch_base": _search.SEARCH_RESULT_BATCH_BASE,
            "search_result_batch_min": _search.SEARCH_RESULT_BATCH_MIN,
            "search_result_batch_max": _search.SEARCH_RESULT_BATCH_MAX,
            "search_metadata_degrade_enabled": _search.SEARCH_METADATA_DEGRADE_ENABLED,
            "search_metadata_degrade_queue_ratio": _search.SEARCH_METADATA_DEGRADE_QUEUE_RATIO,
        }

        try:
            # 首先尝试加载主配置文件（使用程序所在目录的绝对路径）
            config_path = self.path
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    if not isinstance(config, dict):
                        raise ValueError("config root is not an object")
                    self._merge_defaults(config, default_config)
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

                    # v3.76 曾把默认开启的自动检查写入配置，改为默认关闭后丢弃该旧键
                    config.pop("auto_check_updates", None)
                    # 资源快照日志已并入崩溃诊断，旧开关不再使用
                    config.pop("resource_snapshot_logging", None)
                    config.pop("resource_snapshot_interval_ms", None)

                    apply_runtime_performance_config(config.get("performance"))
                    _set_app_language(config.get("language", "zh"))
                    return config
            else:
                print("No config file found, starting with default config")
                apply_runtime_performance_config(default_config.get("performance"))
                return default_config
        except Exception as e:
            print(f"Failed to load config: {e}")
            if isinstance(e, ValueError):
                self.recovered_backup = _preserve_unreadable_file(self.path)
            self.write_blocked = os.path.exists(self.path) and not self.recovered_backup
            self.last_error = str(e)
            apply_runtime_performance_config(default_config.get("performance"))
            return default_config

    def save(self, config):
        """实际写入config.json（原子写入：先写临时文件再重命名；内容无变化时跳过写盘）"""
        config_path = self.path
        tmp_path = config_path + ".tmp"
        if self.write_blocked:
            return False
        try:
            new_content = json.dumps(config, ensure_ascii=False, indent=2)
            last_written = getattr(self, 'last_written', None)
            if last_written is None and os.path.isfile(config_path):
                # 仅首次落盘读取磁盘内容做比较，之后与内存中的上次写入内容比较
                try:
                    with open(config_path, "r", encoding="utf-8") as f:
                        last_written = f.read()
                except Exception:
                    last_written = None
            if last_written == new_content:
                self.last_written = new_content
                return True
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(new_content)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, config_path)
            self.last_written = new_content
            self.last_error = ''
            return True
        except Exception as e:
            print(f"Failed to save config: {e}")
            self.last_error = str(e)
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return False


class SessionController:
    def __init__(self, window):
        self.window = window

    def _capture_named_workspace(self):
        window = self.window
        groups = []
        for tab_widget, content_stack in window._all_groups():
            tabs = []
            for index in range(content_stack.count()):
                tab = content_stack.widget(index)
                if not getattr(tab, 'current_path', ''):
                    continue
                tabs.append({
                    'path': tab.current_path,
                    'is_shell': tab.current_path.startswith('shell:'),
                    'is_pinned': bool(getattr(tab, 'is_pinned', False)),
                    'bookmark_group_color': str(getattr(tab, 'bookmark_group_color', '') or ''),
                    'tab_group_separator_after': bool(getattr(tab, 'tab_group_separator_after', False)),
                    'tab_group_separator_color': str(getattr(tab, 'tab_group_separator_color', '') or ''),
                    'tab_group_separator_name': str(getattr(tab, 'tab_group_separator_name', '') or ''),
                })
            if tabs:
                groups.append({'tabs': tabs, 'active_index': tab_widget.currentIndex()})
        return {'version': 1, 'groups': groups}

    def _get_pinned_paths_from_config(self):
        """兼容 pinned_tabs 的新旧格式，统一提取路径列表。"""
        window = self.window
        paths = []
        for entry in window.config.get("pinned_tabs", []) or []:
            if isinstance(entry, dict):
                p = entry.get('path', '')
            else:
                p = entry
            if p:
                paths.append(p)
        return paths

    def _collect_cached_tabs(self):
        window = self.window
        cached_tabs = []
        pinned_paths = window._get_pinned_paths_from_config()
        pinned_norm = {
            window._normalize_path_for_compare(p)
            for p in pinned_paths if p
        }

        if not hasattr(window, 'tab_widget'):
            return cached_tabs

        # 仅收集左侧主标签组的非固定标签（右侧分屏组由 split_session 单独持久化，避免重复恢复）
        cs = getattr(window, 'content_stack', None)
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

            norm = window._normalize_path_for_compare(current_path)
            if norm in pinned_norm:
                continue
            cached_tabs.append({
                'path': current_path,
                'is_shell': current_path.startswith('shell:'),
                'bookmark_group_color': str(getattr(tab, 'bookmark_group_color', '') or ''),
                'tab_group_separator_after': bool(getattr(tab, 'tab_group_separator_after', False)),
                'tab_group_separator_color': str(getattr(tab, 'tab_group_separator_color', '') or ''),
                'tab_group_separator_name': str(getattr(tab, 'tab_group_separator_name', '') or ''),
            })

        return cached_tabs

    def _collect_split_session(self):
        """收集右侧分屏组会话状态用于持久化恢复。

        仅记录右侧组的非固定标签（固定标签统一在左侧恢复）。返回结构：
        {"active": bool, "tabs": [{"path", "is_shell"}], "active_index": int}。"""
        window = self.window
        state = {"active": False, "tabs": [], "active_index": 0}
        if not getattr(window, '_split_active', False):
            return state
        scs = getattr(window, 'split_content_stack', None)
        stw = getattr(window, 'split_tab_widget', None)
        if scs is None or stw is None or stw.count() == 0:
            return state
        pinned_paths = window._get_pinned_paths_from_config()
        pinned_norm = {
            window._normalize_path_for_compare(p)
            for p in pinned_paths if p
        }
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
            norm = window._normalize_path_for_compare(current_path)
            if norm in pinned_norm:
                continue
            tabs.append({
                'path': current_path,
                'is_shell': current_path.startswith('shell:'),
                'bookmark_group_color': str(getattr(tab, 'bookmark_group_color', '') or ''),
                'tab_group_separator_after': bool(getattr(tab, 'tab_group_separator_after', False)),
                'tab_group_separator_color': str(getattr(tab, 'tab_group_separator_color', '') or ''),
                'tab_group_separator_name': str(getattr(tab, 'tab_group_separator_name', '') or ''),
            })
        if not tabs:
            return state
        state["active"] = True
        state["tabs"] = tabs
        idx = stw.currentIndex()
        state["active_index"] = max(0, min(idx, len(tabs) - 1))
        return state

    def _get_last_active_tab_path(self):
        window = self.window
        try:
            current_tab = window.get_current_tab_widget()
            current_path = getattr(current_tab, 'current_path', '') if current_tab else ''
            return current_path or ""
        except Exception:
            return ""

    def _restore_last_active_tab(self):
        window = self.window
        last_active_path = window.config.get("last_active_tab_path", "")
        if not last_active_path:
            return False

        tab_index = window.find_tab_index_by_path(last_active_path)
        if tab_index < 0:
            return False

        window.tab_widget.setCurrentIndex(tab_index)
        return True

    def _restore_split_session(self):
        """根据持久化的 split_session 恢复右侧分屏组（启动/崩溃恢复时调用）。

        左侧组至少要有一个标签，右侧分屏才独立成立。返回 True 表示已恢复分屏。"""
        window = self.window
        if not window.config.get("enable_cache_tabs", True):
            return False
        state = window.config.get("split_session", {}) or {}
        if not state.get("active"):
            return False
        tabs = state.get("tabs", []) or []
        if not tabs:
            return False
        # 左侧主组必须至少保留一个标签，否则不进入分屏
        if not hasattr(window, 'tab_widget') or window.tab_widget.count() == 0:
            return False
        # 显示右侧分屏 UI 布局，然后逐个在右侧组创建标签
        window._activate_split_layout()
        added = 0
        for tab_info in tabs:
            path = tab_info.get('path', '')
            if not path:
                continue
            try:
                window.add_new_tab(
                    path,
                    is_shell=tab_info.get('is_shell', False),
                    target_tabwidget=window.split_tab_widget,
                    activate=False,
                    bookmark_group_color=(
                        tab_info.get('bookmark_group_color', '')
                        or tab_info.get('tab_group_separator_color', '')
                    ),
                    tab_group_separator_after=tab_info.get('tab_group_separator_after', False),
                    tab_group_separator_color=tab_info.get('tab_group_separator_color', ''),
                    tab_group_separator_name=tab_info.get('tab_group_separator_name', ''),
                )
                added += 1
            except Exception as e:
                debug_print(f"[App] 恢复右侧分屏标签失败: {path} -> {e}")
        if added == 0:
            # 一个都没成功 → 收起分屏，回到单组
            window._teardown_split_group()
            return False
        active_index = state.get("active_index", 0)
        if 0 <= active_index < window.split_tab_widget.count():
            window.split_tab_widget.setCurrentIndex(active_index)
        window._apply_tab_grouping_for_pane(window.split_tab_widget)
        return True

    def save_session_snapshot(self, immediate=False):
        window = self.window
        if not hasattr(window, 'config') or not hasattr(window, 'tab_widget'):
            return

        try:
            import time
            window._last_snapshot_save_time_ms = int(time.monotonic() * 1000)
            cached_tabs = []
            split_session = {"active": False, "tabs": [], "active_index": 0}
            if window.config.get("enable_cache_tabs", True):
                cached_tabs = window._collect_cached_tabs()
                split_session = window._collect_split_session()

            new_active = window._get_last_active_tab_path()

            # 快照内容无变化时跳过写盘，避免无意义IO和日志刷屏
            # 使用 JSON 签名比较，避免 list/dict 对象引用差异导致误判
            import json as _json
            _sig = (_json.dumps(cached_tabs, sort_keys=True, ensure_ascii=False)
                    + '|' + new_active
                    + '|' + _json.dumps(split_session, sort_keys=True, ensure_ascii=False))
            # immediate=True 用于关闭/初始化等关键时机，必须确保落盘，因此绕过签名去重
            if not immediate and _sig == getattr(window, '_last_snapshot_sig', ''):
                return
            window._last_snapshot_sig = _sig

            window.config["cached_tabs"] = cached_tabs
            window.config["last_active_tab_path"] = new_active
            window.config["split_session"] = split_session
            window.save_config(immediate=immediate)
            debug_print(
                f"[App] 会话快照已更新: tabs={len(cached_tabs)}, active='{new_active}', "
                f"split={'on' if split_session.get('active') else 'off'}({len(split_session.get('tabs', []))})"
            )
        except Exception as e:
            print(f"Error saving session snapshot: {e}")

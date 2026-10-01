"""主窗口。"""

import ctypes.wintypes
import json
import os
import socket
import threading
import time

from PyQt5.QtCore import pyqtSignal, pyqtSlot, QDir, QEvent, Qt, QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QAction, QApplication, QDialog, QFrame, QHBoxLayout, QInputDialog, QMainWindow, QMenu, QPushButton,
    QSizePolicy, QToolButton, QVBoxLayout, QWidget,
)

from . import debuglog as _debuglog
from . import i18n as _i18n
from . import search as _search
from . import theme as _theme
from .paths import get_app_data_path
from .i18n import _set_app_language, tr
from .constants import (
    APP_VERSION, HOUSEKEEPING_GC_EVERY_N, HOUSEKEEPING_INTERVAL_MS, MAX_ACTIVE_TOASTS,
    MAX_CLOSED_TABS_HISTORY, MAX_SEARCH_HISTORY, RESOURCE_CRIT_PERCENT, RESOURCE_SAMPLE_INTERVAL_S,
    RESOURCE_WARN_PERCENT, SESSION_SNAPSHOT_DEBOUNCE_MS, SESSION_SNAPSHOT_INTERVAL_MS,
    SESSION_SNAPSHOT_MIN_INTERVAL_MS,
)
from .debuglog import _DEBUG_LOG_PATH, debug_print, set_debug_mode, set_explorer_monitor_debug
from .system import (
    get_process_memory_usage_mb, get_system_cpu_percent, get_system_memory_status, HAS_PYWIN,
    is_supported_title_shortcut_path, launch_detached, launch_shell_tool, normalize_external_launch_dir,
    normalize_terminal_tool_name, _retain_thread_until_finished, _retained_background_threads,
    _thread_name_summary,
)
from .hotkeys import (
    _foreground_pid, _HOTKEY_ENABLE_KEYS, _hotkey_hint, _hotkey_reserved_combos, _hotkey_table, _hotkey_text,
    SHORTCUT_POLL_ACTIVE_MS, SHORTCUT_POLL_INACTIVE_MS, _ShortcutKeyHook, _toolbar_hotkey_tooltip,
    _TOOLBAR_HOTKEY_TOOLTIPS,
)
from .title_shortcuts import TitleShortcutBar
from .widgets import (
    _active_toasts, _build_te_icon, _pinned_tab_icon, ResizableSplitter, _set_tool_icon, show_toast,
)
from .updates import (
    _open_release_page, UPDATE_CHECK_DELAY_MS, UPDATE_CHECK_INTERVAL_S, UPDATE_RELEASES_PAGE,
    UpdateCheckWorker, _version_tuple,
)
from .workers import DiagnosticExportWorker, QuickFindWorker, OpenPathWorker
from .search import apply_runtime_performance_config, QuickFindResultsDialog, _search_cache, SearchDialog
from .fileops import (
    _confirm_file_preview, CopyConflictDialog, DirectoryCompareDialog, FileTaskPanel, _find_copy_conflicts,
    _plan_batch_rename,
)
from .bookmarks import BookmarkDialog, BookmarkManager, BookmarkManagerDialog, _preserve_unreadable_file
from .shellview import (
    _explorer_location_to_path, _ieb_keyboard_filter, _path_is_slow_for_shell, _shell_file_operation_window_open,
)
from .netstatus import is_network_path
from .explorer_tab import FileExplorerTab
from .tabbar import CustomMenuBar, CustomTabBar, DragDropTabWidget, _tab_display_labels, TabListDialog
from .chat import ChatPanel
from .settings import SettingsDialog
from .persistence import ConfigStore, SessionController

if HAS_PYWIN:
    import win32con
    import win32gui


def _tab_git_branch(pane):
    """标签当前路径对应的 Git 分支；导航后尚未重新查询时返回空。"""
    branch_path, branch = getattr(pane, '_git_branch_info', ('', ''))
    return branch if branch and branch_path == getattr(pane, 'current_path', None) else ''


def _tab_target_available(path):
    """网络/慢速位置不在 UI 线程探测（断线时会卡住界面），直接打开，由标签内提示条报告结果。"""
    return _path_is_slow_for_shell(path) or is_network_path(path) or os.path.exists(path)


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

    def _group_palette(self):
        return [
            "#E57373", "#64B5F6", "#81C784", "#FFB74D", "#BA68C8",
            "#4DB6AC", "#F06292", "#7986CB", "#AED581", "#FFD54F",
        ]

    def _is_group_separator_node(self, node):
        return bool(isinstance(node, dict) and node.get('type') == 'url' and node.get('is_group_separator'))

    def _compute_effective_group_colors(self, children):
        """按“分组在右，成员在左”的规则计算每个顶层书签的有效分组色。"""
        node_colors = {}
        if not isinstance(children, list):
            return node_colors
        current_color = None
        for node in reversed(children):
            if not isinstance(node, dict):
                continue
            node_id = node.get('id')
            if self._is_group_separator_node(node):
                current_color = node.get('group_color') or "#64B5F6"
                if node_id:
                    node_colors[node_id] = current_color
            elif current_color and node_id:
                node_colors[node_id] = current_color
        return node_colors

    def _pick_next_group_color(self, children):
        palette = self._group_palette()
        used = set()
        for node in children or []:
            if self._is_group_separator_node(node):
                c = str(node.get('group_color', '')).strip()
                if c:
                    used.add(c)
        for c in palette:
            if c not in used:
                return c
        return palette[len(used) % len(palette)]

    def _pick_next_tab_group_color(self, tab_widget):
        palette = self._group_palette()
        cs = self._content_stack_for(tab_widget)
        used = set()
        if cs is not None:
            for i in range(cs.count()):
                tab = cs.widget(i)
                if tab is None:
                    continue
                if bool(getattr(tab, 'is_pinned', False)):
                    continue
                c = str(getattr(tab, 'bookmark_group_color', '') or '').strip()
                if c:
                    used.add(c)
        for c in palette:
            if c not in used:
                return c
        return palette[len(used) % len(palette)]

    def _apply_tab_grouping_for_pane(self, tab_widget):
        """按标签自身 bookmark_group_color 刷新分组视觉（固定标签不参与分组）。"""
        if tab_widget is None:
            return
        cs = self._content_stack_for(tab_widget)
        if cs is None:
            return
        for i in range(cs.count()):
            tab = cs.widget(i)
            if tab is None:
                continue
            if bool(getattr(tab, 'is_pinned', False)):
                tab.bookmark_group_color = ""
                tab.tab_group_separator_after = False
                tab.tab_group_separator_color = ""
                tab.tab_group_separator_name = ""
                self._apply_tab_group_color(tab_widget, i, tab)
                continue
            # 新策略下不再使用边界元数据，统一清空，分组仅由连续同色决定。
            tab.tab_group_separator_after = False
            tab.tab_group_separator_color = ""
            tab.tab_group_separator_name = ""
            tab.bookmark_group_color = str(getattr(tab, 'bookmark_group_color', '') or '').strip()
            self._apply_tab_group_color(tab_widget, i, tab)

    def insert_tab_group_marker(self, side='right'):
        """F4 分组切换：创建/取消“当前标签及其左侧连续同色（或未分组）标签”的分组。"""
        tw = self.get_active_group_tabwidget()
        cs = self._content_stack_for(tw)
        if tw is None or cs is None or tw.count() <= 0:
            debug_print("[TabGroup] Insert failed: no active tab group")
            show_toast(self, tr("提示"), tr("请先选择一个标签页后再插入分组"), level="warning")
            return False

        cur = tw.currentIndex()
        if cur < 0 or cur >= cs.count():
            debug_print(f"[TabGroup] Insert failed: invalid current index={cur}, tab_count={cs.count()}")
            show_toast(self, tr("提示"), tr("请先选择一个标签页后再插入分组"), level="warning")
            return False

        current_tab = cs.widget(cur)
        if current_tab is None:
            show_toast(self, tr("提示"), tr("请先选择一个标签页后再插入分组"), level="warning")
            return False
        if bool(getattr(current_tab, 'is_pinned', False)):
            debug_print(f"[TabGroup] Insert skipped: current pinned index={cur}")
            show_toast(self, tr("提示"), tr("固定标签不参与分组"), level="warning")
            return False

        cur_color = str(getattr(current_tab, 'bookmark_group_color', '') or '').strip()
        affected_tabs = []
        if cur_color:
            # 取消分组：仅影响当前标签，不连带左侧同色标签
            t = cs.widget(cur)
            if t is not None and not bool(getattr(t, 'is_pinned', False)):
                t.bookmark_group_color = ""
                t.tab_group_separator_after = False
                t.tab_group_separator_color = ""
                t.tab_group_separator_name = ""
                affected_tabs.append(t)
            created = False
        else:
            # 创建分组：影响当前标签及其左侧连续未分组标签
            start = cur
            while start - 1 >= 0:
                left_tab = cs.widget(start - 1)
                if left_tab is None or bool(getattr(left_tab, 'is_pinned', False)):
                    break
                left_color = str(getattr(left_tab, 'bookmark_group_color', '') or '').strip()
                if left_color:
                    break
                start -= 1
            color = self._pick_next_tab_group_color(tw)
            for i in range(start, cur + 1):
                t = cs.widget(i)
                if t is None or bool(getattr(t, 'is_pinned', False)):
                    continue
                t.bookmark_group_color = color
                t.tab_group_separator_after = False
                t.tab_group_separator_color = ""
                t.tab_group_separator_name = ""
                affected_tabs.append(t)
            created = True

        self._apply_tab_grouping_for_pane(tw)
        moved_ungrouped = 0
        if not created and affected_tabs:
            moved_ungrouped = self._move_ungrouped_tabs_before_existing_ungrouped(tw, affected_tabs)
        self.save_pinned_tabs()
        self._schedule_session_snapshot()

        if not created:
            debug_print(
                f"[TabGroup] Removed group: pane={'right' if tw is getattr(self, 'split_tab_widget', None) else 'left'}, "
                f"current={cur}, color={cur_color}, moved_ungrouped={moved_ungrouped}"
            )
            show_toast(self, tr("分组已取消"), tr("当前标签已取消分组"), level="info")
        else:
            debug_print(
                f"[TabGroup] Created group: pane={'right' if tw is getattr(self, 'split_tab_widget', None) else 'left'}, "
                f"current={cur}, color={color}"
            )
            show_toast(self, tr("分组已插入"), tr("当前标签及其左侧已创建分组"), level="info")
        return True

    def _get_bookmark_bar_children(self):
        tree = self.bookmark_manager.get_tree()
        bar = tree.get('bookmark_bar') if isinstance(tree, dict) else None
        children = bar.get('children') if isinstance(bar, dict) else None
        return children if isinstance(children, list) else None

    def _find_top_level_anchor_by_descendant_id(self, bookmark_id, children):
        """根据任意层级节点 ID，找到其所在的顶层书签栏节点（用于分组插入锚点）。"""
        if not bookmark_id or not isinstance(children, list):
            return None

        def _contains_id(node, target_id):
            if not isinstance(node, dict):
                return False
            if node.get('id') == target_id:
                return True
            for child in node.get('children', []) or []:
                if _contains_id(child, target_id):
                    return True
            return False

        for top_node in children:
            if _contains_id(top_node, bookmark_id):
                return top_node
        return None

    def _find_bookmark_anchor_from_path(self, current_path, children):
        """按路径递归匹配书签，返回(顶层锚点节点, 命中的URL节点)。"""
        if not current_path or not isinstance(children, list):
            return None, None

        target_key = self._normalize_path_for_compare(current_path)
        if not target_key:
            return None, None

        def _match_url_node(node):
            if not (isinstance(node, dict) and node.get('type') == 'url'):
                return None
            if self._is_group_separator_node(node):
                return None
            url_key = self._normalize_bookmark_url_for_compare(node.get('url', ''))
            if url_key and url_key == target_key:
                return node
            return None

        def _walk(node):
            direct = _match_url_node(node)
            if direct is not None:
                return direct
            if not isinstance(node, dict):
                return None
            for child in node.get('children', []) or []:
                found = _walk(child)
                if found is not None:
                    return found
            return None

        for top_node in children:
            found = _walk(top_node)
            if found is not None:
                return top_node, found
        return None, None

    def _find_bookmark_node_by_id(self, bookmark_id):
        if not bookmark_id:
            return None
        tree = self.bookmark_manager.get_tree()

        def _walk(node):
            if not isinstance(node, dict):
                return None
            if node.get('id') == bookmark_id:
                return node
            for child in node.get('children', []) or []:
                found = _walk(child)
                if found is not None:
                    return found
            return None

        for root in (tree or {}).values():
            found = _walk(root)
            if found is not None:
                return found
        return None

    def _find_top_level_bookmark_info(self, bookmark_id):
        children = self._get_bookmark_bar_children()
        if not children:
            return None, -1, None
        for idx, node in enumerate(children):
            if isinstance(node, dict) and node.get('id') == bookmark_id:
                return children, idx, node
        return children, -1, None

    def _group_member_indices(self, children, separator_index):
        members = []
        if not isinstance(children, list) or separator_index < 0 or separator_index >= len(children):
            return members
        for i in range(separator_index - 1, -1, -1):
            node = children[i]
            if self._is_group_separator_node(node):
                break
            members.append(i)
        return members

    def _count_group_member_map(self, children):
        member_map = {}
        if not isinstance(children, list):
            return member_map
        for i, node in enumerate(children):
            if self._is_group_separator_node(node):
                member_map[node.get('id')] = len(self._group_member_indices(children, i))
        return member_map

    def _count_open_tabs_by_group_color(self):
        color_count = {}

        def _collect(cs):
            if cs is None:
                return
            for i in range(cs.count()):
                tab = cs.widget(i)
                color = str(getattr(tab, 'bookmark_group_color', '') or '').strip().lower()
                if color:
                    color_count[color] = color_count.get(color, 0) + 1

        _collect(getattr(self, 'content_stack', None))
        _collect(getattr(self, 'split_content_stack', None))
        return color_count

    def _close_tabs_matching_group_color(self, group_color, invert=False):
        """按组色批量关标签。invert=False 关闭同组；invert=True 关闭非同组。"""
        color = str(group_color or '').strip().lower()
        if not color:
            return 0

        left_indices = []
        right_indices = []
        for i in range(self.content_stack.count()):
            tab = self.content_stack.widget(i)
            tab_color = str(getattr(tab, 'bookmark_group_color', '') or '').strip().lower()
            matched = (tab_color == color)
            if (matched and not invert) or ((not matched) and invert):
                left_indices.append(i)
        if getattr(self, 'split_content_stack', None) is not None:
            for i in range(self.split_content_stack.count()):
                tab = self.split_content_stack.widget(i)
                tab_color = str(getattr(tab, 'bookmark_group_color', '') or '').strip().lower()
                matched = (tab_color == color)
                if (matched and not invert) or ((not matched) and invert):
                    right_indices.append(i)

        total = len(left_indices) + len(right_indices)
        if total <= 0:
            return 0

        # 防止把左侧主组全部关闭导致程序直接退出。
        if len(left_indices) >= self.tab_widget.count():
            self.add_new_tab()

        closed = 0
        for i in sorted(right_indices, reverse=True):
            if i < self.split_tab_widget.count():
                self.close_tab(i, target_tabwidget=self.split_tab_widget)
                closed += 1
        for i in sorted(left_indices, reverse=True):
            if i < self.tab_widget.count():
                self.close_tab(i, target_tabwidget=self.tab_widget)
                closed += 1
        return closed

    def _resolve_group_node(self, bookmark_id):
        node = self._find_bookmark_node_by_id(bookmark_id)
        if self._is_group_separator_node(node):
            return node
        return None

    def rename_group_bookmark(self, bookmark_id):
        node = self._resolve_group_node(bookmark_id)
        if node is None:
            return False
        old_name = str(node.get('name', '') or tr("插入分组"))
        text, ok = QInputDialog.getText(self, tr("分组名称"), tr("请输入分组名称："), text=old_name)
        if not ok:
            return False
        new_name = str(text or '').strip()
        if not new_name:
            return False
        node['name'] = new_name
        self.bookmark_manager.save_bookmarks()
        self.populate_bookmark_bar_menu()
        show_toast(self, tr("分组已重命名"), tr("已重命名为：{}").format(new_name), level="info")
        return True

    def toggle_group_collapsed(self, bookmark_id):
        children, idx, node = self._find_top_level_bookmark_info(bookmark_id)
        if node is None or not self._is_group_separator_node(node):
            return False
        collapsed = bool(node.get('group_collapsed', False))
        node['group_collapsed'] = not collapsed
        self.bookmark_manager.save_bookmarks()
        self.populate_bookmark_bar_menu()
        show_toast(
            self,
            tr("提示"),
            tr("该分组目前已展开") if collapsed else tr("该分组目前已折叠"),
            level="info"
        )
        return True

    def close_group_tabs(self, bookmark_id):
        children, idx, node = self._find_top_level_bookmark_info(bookmark_id)
        if node is None or not self._is_group_separator_node(node):
            return False
        group_color = str(node.get('group_color', '') or '').strip().lower()
        if not group_color:
            show_toast(self, tr("提示"), tr("没有已打开的该分组标签页"), level="info")
            return False

        closed = self._close_tabs_matching_group_color(group_color, invert=False)
        if closed <= 0:
            show_toast(self, tr("提示"), tr("没有已打开的该分组标签页"), level="info")
            return False

        show_toast(self, tr("已关闭分组标签页"), tr("已关闭 {} 个标签页").format(closed), level="info")
        return True

    def keep_only_current_group_tabs(self, bookmark_id):
        children, idx, node = self._find_top_level_bookmark_info(bookmark_id)
        if node is None or not self._is_group_separator_node(node):
            return False
        group_color = str(node.get('group_color', '') or '').strip().lower()
        if not group_color:
            show_toast(self, tr("提示"), tr("没有已打开的该分组标签页"), level="info")
            return False
        closed = self._close_tabs_matching_group_color(group_color, invert=True)
        show_toast(self, tr("已保留当前分组"), tr("已关闭其它分组 {} 个标签页").format(closed), level="info")
        return True

    def _find_bookmark_bar_anchor(self):
        """返回用于插入分组的锚点书签：优先最后一次交互书签，其次当前路径匹配书签。"""
        tree = self.bookmark_manager.get_tree()
        bar = tree.get('bookmark_bar') if isinstance(tree, dict) else None
        children = bar.get('children') if isinstance(bar, dict) else None
        if not isinstance(children, list) or not children:
            debug_print("[GroupInsert] Anchor resolve failed: bookmark_bar children empty")
            return None

        last_id = getattr(self, '_last_bookmark_node_id', None)
        if last_id:
            anchor = self._find_top_level_anchor_by_descendant_id(last_id, children)
            if anchor is not None:
                debug_print(
                    f"[GroupInsert] Anchor from last bookmark id={last_id}, top='{anchor.get('name', '')}'"
                )
                return anchor

        current_tab = self.get_active_pane()
        current_path = str(getattr(current_tab, 'current_path', '') or '') if current_tab else ''
        if current_path:
            anchor, matched = self._find_bookmark_anchor_from_path(current_path, children)
            if anchor is not None:
                matched_id = matched.get('id') if isinstance(matched, dict) else None
                if matched_id:
                    self._last_bookmark_node_id = matched_id
                debug_print(
                    f"[GroupInsert] Anchor from current tab path='{current_path}', "
                    f"matched_id={matched_id}, top='{anchor.get('name', '')}'"
                )
                return anchor

        debug_print(
            f"[GroupInsert] Anchor resolve failed: last_id={last_id}, "
            f"current_path='{current_path}'"
        )
        return None

    def _normalize_bookmark_url_for_compare(self, url):
        """将书签 URL 归一化为可与 current_path 比较的键。"""
        from urllib.parse import unquote
        u = str(url or '').strip()
        if not u:
            return ""
        try:
            if u.startswith('file:'):
                if u.startswith('file://///'):
                    local_path = '\\\\' + unquote(u[10:]).replace('/', '\\')
                elif u.startswith('file:////'):
                    local_path = '\\\\' + unquote(u[9:]).replace('/', '\\')
                elif u.startswith('file:///'):
                    local_path = unquote(u[8:])
                    if os.name == 'nt' and local_path.startswith('/'):
                        local_path = local_path[1:]
                    local_path = local_path.replace('/', '\\')
                else:
                    local_path = '\\\\' + unquote(u[7:]).replace('/', '\\')
                return self._normalize_path_for_compare(local_path)
            if u.startswith('shell:'):
                return self._normalize_path_for_compare(u)
            if os.path.isabs(u):
                return self._normalize_path_for_compare(u)
        except Exception:
            pass
        return self._normalize_path_for_compare(u)

    def insert_group_bookmark(self, side='right'):
        """兼容旧入口：改为操作当前标签分组，不再依赖书签。"""
        return self.insert_tab_group_marker()

    def show_insert_group_menu(self):
        # 入口简化：仅保留默认分组动作，不再弹出左右选项。
        self.insert_tab_group_marker()

    def _get_tab_group_icon(self, color_hex, separator=False):
        """为标签分组生成小色块图标（separator 使用不同形状）。"""
        from PyQt5.QtCore import QPoint
        from PyQt5.QtGui import QPixmap, QPainter, QColor, QIcon
        key = (str(color_hex or '').lower(), bool(separator))
        cache = getattr(self, '_tab_group_icon_cache', None)
        if cache is None:
            cache = {}
            self._tab_group_icon_cache = cache
        if key in cache:
            return cache[key]

        pix = QPixmap(10, 10)
        pix.fill(Qt.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setPen(Qt.NoPen)
        c = QColor(key[0] if key[0] else '#64B5F6')
        p.setBrush(c)
        if separator:
            # 分隔符：菱形，和普通成员点做区分
            pts = [
                QPoint(5, 1),
                QPoint(9, 5),
                QPoint(5, 9),
                QPoint(1, 5),
            ]
            p.drawPolygon(*pts)
        else:
            # 成员：圆点
            p.drawEllipse(1, 1, 8, 8)
        p.end()
        icon = QIcon(pix)
        cache[key] = icon
        return icon

    def _apply_tab_group_color(self, tab_widget, index, tab_obj=None):
        if tab_widget is None or index < 0:
            return
        tab_ref = tab_obj
        if tab_ref is None:
            try:
                cs = self._content_stack_for(tab_widget)
                if cs is not None and index < cs.count():
                    tab_ref = cs.widget(index)
            except Exception:
                tab_ref = None
        color_hex = str(getattr(tab_ref, 'bookmark_group_color', '') or '').strip() if tab_ref is not None else ''
        from PyQt5.QtGui import QColor
        show_markers = bool(getattr(self, 'config', {}).get('show_tab_group_markers', True))
        if color_hex and show_markers:
            group_color = QColor(color_hex)
            tab_widget.tabBar().setTabTextColor(index, group_color if _theme.is_dark() else group_color.darker(125))
        else:
            tab_widget.tabBar().setTabTextColor(index, QColor(_theme.fg("#505050")))
        tab_widget.setTabIcon(index, _pinned_tab_icon() if getattr(tab_ref, 'is_pinned', False) else QIcon())
        try:
            tab_widget.tabBar().update()
        except Exception:
            pass

    def apply_tab_group_markers_config(self):
        """根据配置刷新标签分组视觉标记（左右标签组）。"""
        self._apply_tab_grouping_for_pane(self.tab_widget)
        self._apply_tab_grouping_for_pane(getattr(self, 'split_tab_widget', None))

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
                    insert_idx = self._tabbar_insert_index_from_global(tw, gpos)
                    return tw, insert_idx
            except Exception:
                continue
        return None, None

    def _tabbar_insert_index_from_global(self, tab_widget, gpos):
        """根据全局坐标计算标签栏插入槽位（0..count）。

        命中标签左半返回该标签索引（插到其前），命中右半返回索引+1（插到其后）。"""
        if tab_widget is None:
            return -1
        bar = tab_widget.tabBar() if hasattr(tab_widget, 'tabBar') else None
        if bar is None:
            return -1
        count = bar.count()
        if count <= 0:
            return 0
        local = bar.mapFromGlobal(gpos)
        idx = bar.tabAt(local)
        if idx < 0:
            first_rect = bar.tabRect(0)
            last_rect = bar.tabRect(count - 1)
            if first_rect.isValid() and local.x() <= first_rect.left():
                return 0
            if last_rect.isValid() and local.x() >= last_rect.right():
                return count
            return count
        rect = bar.tabRect(idx)
        if not rect.isValid():
            return min(max(idx, 0), count)
        return idx if local.x() < rect.center().x() else idx + 1

    def _normalize_drop_insert_index(self, tab_widget, raw_index):
        if tab_widget is None:
            return -1
        count = tab_widget.count()
        try:
            idx = int(raw_index)
        except Exception:
            idx = -1
        if idx < 0 or idx > count:
            return count
        return idx

    def _update_drag_insert_indicators(self, dest_tabwidget, dest_index):
        for tw, _cs in self._all_groups():
            bar = tw.tabBar() if tw is not None else None
            if bar is None:
                continue
            if tw is dest_tabwidget and hasattr(bar, 'set_drop_insert_indicator'):
                idx = self._normalize_drop_insert_index(tw, dest_index)
                bar.set_drop_insert_indicator(idx)
            elif hasattr(bar, 'clear_drop_insert_indicator'):
                bar.clear_drop_insert_indicator()

    def _clear_drag_insert_indicators(self):
        for tw, _cs in self._all_groups():
            bar = tw.tabBar() if tw is not None else None
            if bar is not None and hasattr(bar, 'clear_drop_insert_indicator'):
                bar.clear_drop_insert_indicator()

    def _get_tab_group_ranges_for_drag(self, tab_widget):
        """返回可拖拽分组块区间列表（start, end）。

        新策略：连续同色标签为一组；未分组标签按单个块处理。"""
        cs = self._content_stack_for(tab_widget)
        if cs is None:
            return []

        ranges = []
        i = 0
        while i < cs.count():
            tab = cs.widget(i)
            if tab is None or bool(getattr(tab, 'is_pinned', False)):
                i += 1
                continue
            color = str(getattr(tab, 'bookmark_group_color', '') or '').strip()
            if not color:
                ranges.append((i, i))
                i += 1
                continue
            end = i
            while end + 1 < cs.count():
                nxt = cs.widget(end + 1)
                if nxt is None or bool(getattr(nxt, 'is_pinned', False)):
                    break
                nxt_color = str(getattr(nxt, 'bookmark_group_color', '') or '').strip()
                if nxt_color != color:
                    break
                end += 1
            ranges.append((i, end))
            i = end + 1
        return ranges

    def _get_tab_group_block_for_drag(self, tab_widget, index):
        """根据任意索引返回可拖拽区块；有组色时返回连续同色块，否则返回单标签块。"""
        cs = self._content_stack_for(tab_widget)
        if cs is None or index < 0 or index >= cs.count():
            return None
        tab = cs.widget(index)
        if tab is None or bool(getattr(tab, 'is_pinned', False)):
            return None
        color = str(getattr(tab, 'bookmark_group_color', '') or '').strip()
        if not color:
            return (index, index)
        start = index
        while start - 1 >= 0:
            left = cs.widget(start - 1)
            if left is None or bool(getattr(left, 'is_pinned', False)):
                break
            if str(getattr(left, 'bookmark_group_color', '') or '').strip() != color:
                break
            start -= 1
        end = index
        while end + 1 < cs.count():
            right = cs.widget(end + 1)
            if right is None or bool(getattr(right, 'is_pinned', False)):
                break
            if str(getattr(right, 'bookmark_group_color', '') or '').strip() != color:
                break
            end += 1
        return (start, end)

    def _apply_right_neighbor_grouping_for_moved_tabs(self, tab_widget, moved_tabs):
        """拖动落位后，以新位置右侧相邻标签为准更新被拖动标签的分组。"""
        tw, cs, _is_right = self._resolve_group(tab_widget)
        if tw is None or cs is None or not moved_tabs:
            return

        tabs = [t for t in moved_tabs if t is not None and not bool(getattr(t, 'is_pinned', False))]
        if not tabs:
            return

        first_idx = cs.indexOf(tabs[0])
        if first_idx < 0:
            return

        right_idx = first_idx + len(tabs)
        target_color = ""
        if 0 <= right_idx < cs.count():
            right_tab = cs.widget(right_idx)
            if right_tab is not None and not bool(getattr(right_tab, 'is_pinned', False)):
                target_color = str(getattr(right_tab, 'bookmark_group_color', '') or '').strip()

        for t in tabs:
            t.bookmark_group_color = target_color
            t.tab_group_separator_after = False
            t.tab_group_separator_color = ""
            t.tab_group_separator_name = ""

        self._apply_tab_grouping_for_pane(tw)

    def _split_group_color_after_insertion_if_needed(self, tab_widget, insert_start, insert_len):
        """当插入位置把同色分组切开时，将右半段改为新颜色，形成两个不同分组。"""
        tw, cs, _is_right = self._resolve_group(tab_widget)
        if tw is None or cs is None:
            return False
        if insert_start is None or insert_len is None:
            return False
        insert_start = int(insert_start)
        insert_len = int(insert_len)
        if insert_len <= 0:
            return False

        left_idx = insert_start - 1
        right_idx = insert_start + insert_len
        if left_idx < 0 or right_idx >= cs.count():
            return False

        left_tab = cs.widget(left_idx)
        right_tab = cs.widget(right_idx)
        if left_tab is None or right_tab is None:
            return False
        if bool(getattr(left_tab, 'is_pinned', False)) or bool(getattr(right_tab, 'is_pinned', False)):
            return False

        split_color = str(getattr(left_tab, 'bookmark_group_color', '') or '').strip()
        right_color = str(getattr(right_tab, 'bookmark_group_color', '') or '').strip()
        if not split_color or right_color != split_color:
            return False

        new_color = self._pick_next_tab_group_color(tw)
        if new_color == split_color:
            palette = self._group_palette()
            for cand in palette:
                if cand != split_color:
                    new_color = cand
                    break
        if not new_color or new_color == split_color:
            return False

        changed = 0
        idx = right_idx
        while idx < cs.count():
            tab = cs.widget(idx)
            if tab is None or bool(getattr(tab, 'is_pinned', False)):
                break
            color = str(getattr(tab, 'bookmark_group_color', '') or '').strip()
            if color != split_color:
                break
            tab.bookmark_group_color = new_color
            tab.tab_group_separator_after = False
            tab.tab_group_separator_color = ""
            tab.tab_group_separator_name = ""
            changed += 1
            idx += 1

        return changed > 0

    def _move_ungrouped_tabs_before_existing_ungrouped(self, tab_widget, candidate_tabs):
        """将候选中的非固定且无分组色标签移到“现有非分组标签”左侧，保持相对顺序。

        若当前顺序已满足，或不存在可对齐的现有非分组标签，则不执行移动。"""
        tw, cs, _is_right = self._resolve_group(tab_widget)
        if tw is None or cs is None or not candidate_tabs:
            return 0

        candidate_set = {tab for tab in candidate_tabs if tab is not None}
        if not candidate_set:
            return 0

        all_tabs = []
        for i in range(cs.count()):
            w = cs.widget(i)
            if w is not None:
                all_tabs.append(w)

        move_tabs = []
        for w in all_tabs:
            if (w in candidate_set and
                    not bool(getattr(w, 'is_pinned', False)) and
                    not str(getattr(w, 'bookmark_group_color', '') or '').strip()):
                move_tabs.append(w)

        if not move_tabs:
            return 0

        move_set = set(move_tabs)

        # 目标插入点：当前列表中“非候选的首个非分组标签”位置。
        first_existing_ungrouped_idx = -1
        for i, w in enumerate(all_tabs):
            if w in move_set:
                continue
            if bool(getattr(w, 'is_pinned', False)):
                continue
            if not str(getattr(w, 'bookmark_group_color', '') or '').strip():
                first_existing_ungrouped_idx = i
                break

        keep_tabs = [w for w in all_tabs if w not in move_set]
        if first_existing_ungrouped_idx < 0:
            # 尚无未分组区域：把本次取消得到的未分组块放到最右侧，形成未分组区域。
            insert_pos = len(keep_tabs)
        else:
            insert_pos = 0
            for w in all_tabs[:first_existing_ungrouped_idx]:
                if w not in move_set:
                    insert_pos += 1

        current_idx = tw.currentIndex()
        current_tab = cs.widget(current_idx) if current_idx >= 0 else None
        new_tabs = keep_tabs[:insert_pos] + move_tabs + keep_tabs[insert_pos:]

        # 顺序无变化则不动作。
        if len(new_tabs) == len(all_tabs) and all(a is b for a, b in zip(new_tabs, all_tabs)):
            return 0

        tw.clear()
        while cs.count() > 0:
            w = cs.widget(0)
            cs.removeWidget(w)

        for w in new_tabs:
            tw.addTab(QWidget(), "")
            cs.addWidget(w)
            try:
                w.update_tab_title()
            except Exception:
                pass

        if current_tab is not None:
            new_idx = cs.indexOf(current_tab)
            if new_idx >= 0:
                tw.setCurrentIndex(new_idx)

        self._apply_tab_grouping_for_pane(tw)
        return len(move_tabs)

    def _clear_tab_group_separator_metadata(self, tabs):
        """将拖动块中的分组边界降级为普通成员，落位后按新位置边界重新归组。"""
        for tab in tabs or []:
            if tab is None:
                continue
            tab.tab_group_separator_after = False
            tab.tab_group_separator_color = ""
            tab.tab_group_separator_name = ""

    def reorder_tab_group_within_group(self, tab_widget, block, dest_index, anchor_widget=None):
        """同组内整块移动标签分组（右键拖拽）。"""
        tw, cs, _is_right = self._resolve_group(tab_widget)
        if tw is None or cs is None or not block or len(block) != 2:
            return False
        start, end = int(block[0]), int(block[1])
        if start < 0 or end >= cs.count() or start > end:
            return False
        if start <= int(dest_index) <= end + 1:
            return False

        moving = []
        for i in range(start, end + 1):
            w = cs.widget(i)
            if w is None or bool(getattr(w, 'is_pinned', False)):
                return False
            moving.append(w)

        for i in range(end, start - 1, -1):
            w = cs.widget(i)
            cs.removeWidget(w)
            tw.removeTab(i)

        # 整组拖动时，原分组边界随块移动后应视为普通成员，避免在新位置重建旧边界。
        self._clear_tab_group_separator_metadata(moving)

        if dest_index is None:
            dest_index = tw.count()
        if dest_index > end:
            dest_index -= len(moving)
        dest_index = max(0, min(int(dest_index), tw.count()))

        pinned_count = 0
        for i in range(cs.count()):
            w = cs.widget(i)
            if w is not None and bool(getattr(w, 'is_pinned', False)):
                pinned_count += 1
        dest_index = max(dest_index, pinned_count)

        for offset, w in enumerate(moving):
            insert_at = dest_index + offset
            cs.insertWidget(insert_at, w)
            tw.insertTab(insert_at, QWidget(), "")
            try:
                w.update_tab_title()
            except Exception:
                pass

        # 整组拖动保持原分组颜色，不受目标位置分组颜色影响。
        self._split_group_color_after_insertion_if_needed(tw, dest_index, len(moving))

        new_current = dest_index
        if anchor_widget in moving:
            new_current = dest_index + moving.index(anchor_widget)
        tw.setCurrentIndex(new_current)

        self._apply_tab_grouping_for_pane(tw)
        self.save_pinned_tabs()
        self._schedule_session_snapshot()
        return True

    def move_tab_group_across_groups(self, source_tabwidget, block, anchor_widget, dest_tabwidget, dest_index):
        """跨组整块移动标签分组（右键拖拽）。"""
        src_tw, src_cs, src_is_right = self._resolve_group(source_tabwidget)
        dst_tw, dst_cs, dst_is_right = self._resolve_group(dest_tabwidget)
        if src_tw is dst_tw or src_cs is None or dst_cs is None:
            return False
        if not block or len(block) != 2:
            return False

        start, end = int(block[0]), int(block[1])
        if start < 0 or end >= src_cs.count() or start > end:
            return False

        moving = []
        for i in range(start, end + 1):
            w = src_cs.widget(i)
            if w is None or bool(getattr(w, 'is_pinned', False)):
                return False
            moving.append(w)
        block_len = len(moving)
        if block_len <= 0:
            return False

        if not src_is_right and (src_tw.count() - block_len) < 1:
            show_toast(self, tr("拖拽标签"), tr("左侧至少需要保留一个标签页。"), level="warning", duration=2000)
            return False

        dst_count = dst_tw.count()
        if dest_index is None or dest_index < 0 or dest_index > dst_count:
            dest_index = dst_count

        dst_pinned_count = 0
        for i in range(dst_cs.count()):
            w = dst_cs.widget(i)
            if w is not None and bool(getattr(w, 'is_pinned', False)):
                dst_pinned_count += 1
        dest_index = max(int(dest_index), dst_pinned_count)
        dest_index = min(dest_index, dst_tw.count())

        for i in range(end, start - 1, -1):
            w = src_cs.widget(i)
            src_cs.removeWidget(w)
            src_tw.removeTab(i)

        # 整组跨组拖动同样不保留原边界，整组保持原颜色插入。
        self._clear_tab_group_separator_metadata(moving)

        for offset, w in enumerate(moving):
            insert_at = dest_index + offset
            dst_cs.insertWidget(insert_at, w)
            dst_tw.insertTab(insert_at, QWidget(), "")
            try:
                w.update_tab_title()
            except Exception:
                pass

        # 整组拖动保持原分组颜色，不受目标位置分组颜色影响。
        self._split_group_color_after_insertion_if_needed(dst_tw, dest_index, len(moving))

        new_current = dest_index
        if anchor_widget in moving:
            new_current = dest_index + moving.index(anchor_widget)
        dst_tw.setCurrentIndex(new_current)
        self._active_pane = dst_cs.widget(new_current) if new_current >= 0 else self._active_pane

        if src_is_right and src_tw.count() == 0:
            self._teardown_split_group()

        try:
            self._apply_tab_grouping_for_pane(src_tw)
        except Exception:
            pass
        try:
            self._apply_tab_grouping_for_pane(dst_tw)
        except Exception:
            pass
        try:
            self.update_navigation_buttons()
        except Exception:
            pass
        self.save_pinned_tabs()
        self._schedule_session_snapshot()
        return True

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
        was_separator = bool(getattr(content_widget, 'tab_group_separator_after', False))
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
        if was_separator:
            # 单拖边界标签跨组时按普通成员处理，避免在目标位置重建旧分组边界。
            content_widget.tab_group_separator_after = False
            content_widget.tab_group_separator_color = ""
            content_widget.tab_group_separator_name = ""
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
        self._apply_right_neighbor_grouping_for_moved_tabs(dst_tw, [content_widget])
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

    def _hibernate_idle_tabs(self):
        """后台标签超过设定时间未访问时释放其 Shell 视图。"""
        try:
            minutes = int(self.config.get("tab_hibernate_minutes", 30) or 0)
        except (TypeError, ValueError):
            minutes = 30
        if minutes <= 0:
            return 0
        cutoff = time.monotonic() - minutes * 60
        released = 0
        for _tabs, stack in self._all_groups():
            current = stack.currentWidget()
            for index in range(stack.count()):
                tab = stack.widget(index)
                if tab is None or tab is current or not hasattr(tab, 'hibernate_shell_view'):
                    continue
                last_active = getattr(tab, '_last_active_at', None)
                if last_active is not None and last_active <= cutoff and tab.hibernate_shell_view():
                    released += 1
        if released:
            debug_print(f"[Hibernate] Released {released} idle tab(s) after {minutes} min")
            self._schedule_tab_labels()
        return released

    def _schedule_tab_labels(self):
        if not hasattr(self, '_tab_labels_timer'):
            self._tab_labels_timer = QTimer(self)
            self._tab_labels_timer.setSingleShot(True)
            self._tab_labels_timer.timeout.connect(self._refresh_tab_labels)
        self._tab_labels_timer.start(0)

    def _refresh_tab_labels(self):
        if not hasattr(self, 'content_stack'):
            return
        entries = [(tabs, index, stack.widget(index)) for tabs, stack in self._all_groups()
                   for index in range(min(tabs.count(), stack.count()))]
        state = [(_i18n._app_language, id(tabs), id(pane), getattr(pane, 'current_path', ''),
                  bool(getattr(pane, 'is_pinned', False)), _tab_git_branch(pane),
                  bool(getattr(pane, '_hibernated', False))) for tabs, _index, pane in entries]
        if state == getattr(self, '_tab_label_state', None):
            return
        self._tab_label_state = state
        from PyQt5.QtGui import QIcon
        labels = _tab_display_labels([str(getattr(pane, 'current_path', '') or '') for _tabs, _index, pane in entries])
        for (tabs, index, pane), label in zip(entries, labels):
            pane._display_tab_label = label
            pinned = bool(getattr(pane, 'is_pinned', False))
            tabs.setTabText(index, label)
            tabs.setTabIcon(index, _pinned_tab_icon() if pinned else QIcon())
            path = str(getattr(pane, 'current_path', '') or '')
            tip = [tr('已固定')] if pinned else []
            tip.append(path)
            branch = _tab_git_branch(pane)
            if branch:
                tip.append(tr('Git 分支: {}').format(branch))
            if getattr(pane, '_hibernated', False):
                tip.append(tr('已休眠，切换到该标签时重新加载'))
            tabs.setTabToolTip(index, '\n'.join(tip))

    def show_tab_list(self):
        dialog = TabListDialog(self)
        dialog.exec_()
        dialog.deleteLater()

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
        self._update_active_pane_indicator()

    def _update_active_pane_indicator(self):
        if not hasattr(self, 'content_stack'):
            return
        split = bool(getattr(self, '_split_active', False))
        current = self.get_active_pane()
        for tab_widget, stack in self._all_groups():
            if stack is None or tab_widget is None:
                continue
            pane = stack.widget(tab_widget.currentIndex())
            active = not split or pane is current
            tabbar = tab_widget.tabBar()
            if tabbar.property('activePane') != active:
                tabbar.setProperty('activePane', active)
                tabbar.style().unpolish(tabbar)
                tabbar.style().polish(tabbar)
                tabbar.update()
            path_bar = getattr(pane, 'path_bar', None)
            if path_bar is not None:
                path_bar.set_pane_active(active, split)
    
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
        # 快捷方式区与右侧工具组之间始终保留视觉间隔（与 Git 显示状态解耦）。
        if hasattr(self, 'shortcut_git_separator'):
            self.shortcut_git_separator.setVisible(self.config.get("enable_title_shortcuts", True))
        self._refresh_titlebar_layout_now()

    def apply_title_shortcuts_config(self):
        """根据配置显示/隐藏标题栏快捷方式区域。"""
        enable = self.config.get("enable_title_shortcuts", True)
        if hasattr(self, 'title_shortcut_bar'):
            self.title_shortcut_bar.setVisible(enable)
            # 从关闭切回开启时，立即按配置重建按钮，确保已保存快捷方式可见。
            if enable:
                self.refresh_title_shortcuts_ui()
        if hasattr(self, 'shortcut_git_separator'):
            self.shortcut_git_separator.setVisible(enable)
        self._refresh_titlebar_layout_now()

    def _refresh_titlebar_layout_now(self):
        """显隐标题栏控件后立即刷新布局，避免按钮左移延迟到下一次重绘。"""
        titlebar = getattr(self, 'titlebar_widget', None)
        if not titlebar:
            return
        layout = titlebar.layout()
        if layout is not None:
            layout.invalidate()
            layout.activate()
        titlebar.updateGeometry()
        titlebar.update()
        QApplication.sendPostedEvents(None, QEvent.LayoutRequest)

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

    def send_prompt_to_ai(self, prompt: str) -> bool:
        """确保 AI 面板可见并发送一条预置提示词。"""
        cfg = self.config.get("ai_chat", {})
        if not cfg.get("enabled", False):
            show_toast(self, tr("提示"), tr("AI 助手未启用，请先在设置中开启"), level="warning")
            return False
        self._ensure_chat_panel_created()
        if not self.chat_panel.isVisible():
            self.toggle_chat_panel()
        self.update_chat_context()
        try:
            return bool(self.chat_panel.submit_external_prompt(prompt))
        except Exception as e:
            show_toast(self, tr("错误"), tr("❌ 请求失败: {}").format(e), level="error")
            return False

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
        key = _hotkey_text(self.config, 'split_view')
        show_toast(self, tr("分屏对比"), tr("已将当前标签移到右侧。再次按 {} 合并回左侧。").format(key) if key
                   else tr("已将当前标签移到右侧。"), level="info", duration=2000)

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
        self._update_active_pane_indicator()
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
    def add_new_tab(self, path="", is_shell=False, select_file=None, target_tabwidget=None, activate=True,
                    bookmark_group_color=None, bookmark_source_node_id=None,
                    tab_group_separator_after=False, tab_group_separator_color="", tab_group_separator_name=""):
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
            if hasattr(tab, 'set_bottom_statusbar_visible'):
                tab.set_bottom_statusbar_visible(self.config.get("show_bottom_statusbar", True))
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
        tab.bookmark_group_color = bookmark_group_color or ""
        tab.bookmark_source_node_id = str(bookmark_source_node_id or "")
        tab.tab_group_separator_after = bool(tab_group_separator_after)
        tab.tab_group_separator_color = str(tab_group_separator_color or "")
        tab.tab_group_separator_name = str(tab_group_separator_name or "")
        
        # 同时添加到 tab_widget（占位标签）和 content_stack（实际内容）
        tab_index = tab_widget.addTab(QWidget(), "")
        content_stack.addWidget(tab)
        # 启动恢复时也立即用统一逻辑生成短标题，避免显示全路径。
        try:
            tab.update_tab_title()
        except Exception:
            pass
        
        # activate=True：切到该标签（触发 showEvent → 首次导航）。
        # activate=False：不激活，标签保持隐藏，其导航延迟到用户首次切过去（懒加载）。
        if activate:
            tab_widget.setCurrentIndex(tab_index)

        self._apply_tab_group_color(tab_widget, tab_index, tab)
        self._apply_tab_grouping_for_pane(tab_widget)
        if getattr(tab, 'bookmark_group_color', ''):
            try:
                self.populate_bookmark_bar_menu()
            except Exception:
                pass
        
        # 更新导航按钮状态（确保新标签页的按钮状态正确）
        self.update_navigation_buttons()

        self._schedule_session_snapshot()
        
        return tab_index


    def close_tab(self, index, target_tabwidget=None):
        tab_widget, content_stack, is_right = self._resolve_group(target_tabwidget)
        if not is_right and tab_widget.count() == 1:
            panel = getattr(self, '_file_task_panel', None)
            if panel and panel.has_running_tasks():
                self.show_file_tasks()
                return
        tab = content_stack.widget(index) if (content_stack and 0 <= index < content_stack.count()) else None
        closed_group_color = str(getattr(tab, 'bookmark_group_color', '') or '').strip()
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
            self._apply_tab_grouping_for_pane(tab_widget)
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
                self._apply_tab_grouping_for_pane(self.tab_widget)
                self._schedule_session_snapshot()
            else:
                self.close()

        if closed_group_color:
            try:
                self.populate_bookmark_bar_menu()
            except Exception:
                pass


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

    def _deferred_on_tab_changed(self, index):
        """延后重试 tab/content_stack 同步，规避启动期短暂初始化竞态。"""
        self._tab_sync_retry_pending = False
        try:
            if self.tab_widget.currentIndex() != index:
                return
        except Exception:
            return
        self.on_tab_changed(index)

    def on_tab_changed(self, index):
        if index >= 0:
            # 调试信息：检查同步状态
            debug_print(f"[TabSwitch] Tab changed to index {index}")
            if hasattr(self, 'content_stack'):
                debug_print(f"[TabSwitch] content_stack has {self.content_stack.count()} widgets, tab_widget has {self.tab_widget.count()} tabs")

            # 同步 content_stack 的显示（拖拽期间也需要保持同步）
            if hasattr(self, 'content_stack') and index < self.content_stack.count():
                self._tab_sync_retry_attempts = 0
                self.content_stack.setCurrentIndex(index)
                debug_print(f"[TabSwitch] Set content_stack to index {index}")
            else:
                attempts = int(getattr(self, '_tab_sync_retry_attempts', 0))
                pending = bool(getattr(self, '_tab_sync_retry_pending', False))
                if attempts < 3 and not pending:
                    self._tab_sync_retry_attempts = attempts + 1
                    self._tab_sync_retry_pending = True
                    debug_print(
                        f"[TabSwitch] content_stack not ready (retry {self._tab_sync_retry_attempts}/3), "
                        f"deferring sync for index {index}"
                    )
                    QTimer.singleShot(60, lambda idx=index: self._deferred_on_tab_changed(idx))
                else:
                    debug_print(
                        f"[TabSwitch] WARNING: Cannot sync - content_stack count is "
                        f"{self.content_stack.count() if hasattr(self, 'content_stack') else 'N/A'}"
                    )
                return

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
                current_path = str(getattr(tab, 'current_path', '') or '')
                source_id = str(getattr(tab, 'bookmark_source_node_id', '') or '').strip()
                if source_id:
                    self._last_bookmark_node_id = source_id
                    debug_print(f"[GroupInsert] Active tab carries bookmark source id={source_id}")
                elif current_path:
                    children = self._get_bookmark_bar_children()
                    _anchor, matched = self._find_bookmark_anchor_from_path(current_path, children)
                    matched_id = matched.get('id') if isinstance(matched, dict) else ''
                    if matched_id:
                        tab.bookmark_source_node_id = str(matched_id)
                        self._last_bookmark_node_id = str(matched_id)
                        debug_print(
                            f"[GroupInsert] Active tab path matched bookmark id={matched_id}, "
                            f"path='{current_path}'"
                        )
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
        _theme.bind_style(titlebar, "background-color: #f3f3f3;")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(10, 0, 0, 0)
        titlebar_layout.setSpacing(0)
        # 左侧使用弹性空白，所有功能按钮区域始终靠右紧凑排列。
        titlebar_layout.addStretch(1)
        
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
        _theme.bind_style(self.shortcut_git_separator, "background-color: #d0d0d0; max-width: 1px;")
        self.shortcut_git_separator.setFixedWidth(1)
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
        _theme.bind_style(self.git_tools_separator, "background-color: #d0d0d0; max-width: 1px;")
        self.git_tools_separator.setFixedWidth(1)
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
        _theme.bind_style(separator, "background-color: #d0d0d0; max-width: 1px;")
        separator.setFixedWidth(1)
        separator.setFixedHeight(int(20 * getattr(self, 'dpi_scale', 1.0)))
        titlebar_layout.addWidget(separator)
        
        # 标签栏导航按钮（从标签栏移到这里）
        # 后退按钮
        self.back_button = QPushButton("◀")
        self.back_button.setToolTip(_toolbar_hotkey_tooltip(self, 'back_button'))
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
        self.forward_button.setToolTip(_toolbar_hotkey_tooltip(self, 'forward_button'))
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
        self.add_tab_button.setToolTip(_toolbar_hotkey_tooltip(self, 'add_tab_button'))
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
        self.reopen_tab_button.setToolTip(_toolbar_hotkey_tooltip(self, 'reopen_tab_button'))
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

        from PyQt5.QtWidgets import QStyle
        self.tab_list_button = QToolButton(self)
        self.tab_list_button.setToolTip(tr('查找标签'))
        self.tab_list_button.setFixedSize(btn_size, btn_size)
        _set_tool_icon(self.tab_list_button, 'view-list-details', QStyle.SP_FileDialogListView)
        self.tab_list_button.clicked.connect(self.show_tab_list)
        titlebar_layout.addWidget(self.tab_list_button)
        
        # 搜索按钮
        self.search_button = QPushButton("⌕")
        self.search_button.setToolTip(_toolbar_hotkey_tooltip(self, 'search_button'))
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

        from PyQt5.QtWidgets import QStyle
        self.workspace_tools_button = QToolButton(self)
        self.workspace_tools_button.setToolTip(tr("工作区与文件工具"))
        _set_tool_icon(self.workspace_tools_button, 'workspace-tools', QStyle.SP_FileDialogDetailedView)
        self.workspace_tools_button.setFixedSize(btn_size, btn_size)
        self.workspace_tools_button.setPopupMode(QToolButton.InstantPopup)
        tools_menu = QMenu(self.workspace_tools_button)
        self._populate_workspace_tools_menu(tools_menu)
        self.workspace_tools_menu = tools_menu
        self.workspace_tools_button.setMenu(tools_menu)
        titlebar_layout.addWidget(self.workspace_tools_button)

        self.file_tasks_button = QToolButton(self)
        self.file_tasks_button.setFixedSize(int(64 * self.dpi_scale), btn_size)
        self.file_tasks_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.file_tasks_button.clicked.connect(self.show_file_tasks)
        titlebar_layout.addWidget(self.file_tasks_button)
        self._update_task_indicator()

        # 插入分组按钮
        self.insert_group_btn = QPushButton("☰")
        self.insert_group_btn.setToolTip(_toolbar_hotkey_tooltip(self, 'insert_group_btn'))
        self.insert_group_btn.setFixedSize(btn_size, btn_size)
        self.insert_group_btn.setStyleSheet(f"""
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
        self.insert_group_btn.clicked.connect(self.insert_tab_group_marker)
        titlebar_layout.addWidget(self.insert_group_btn)
        
        # 分屏对比按钮（切换右侧第二个独立浏览面板）
        self.split_view_btn = QPushButton("◫")
        self.split_view_btn.setToolTip(_toolbar_hotkey_tooltip(self, 'split_view_btn'))
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
        self.ai_chat_btn.setToolTip(_toolbar_hotkey_tooltip(self, 'ai_chat_btn'))
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
        from PyQt5.QtWidgets import QStyle
        icons = (
            (self.git_log_button, 'app-tortoisegit-log', QStyle.SP_DirLinkIcon),
            (self.git_commit_button, 'app-tortoisegit-commit', QStyle.SP_ArrowUp),
            (self.git_bash_button, 'app-git-bash', QStyle.SP_ComputerIcon),
            (self.cmd_button, 'app-cmd', QStyle.SP_ComputerIcon),
            (self.powershell_button, 'app-powershell', QStyle.SP_DesktopIcon),
            (self.calculator_button, 'accessories-calculator', QStyle.SP_FileDialogDetailedView),
            (self.back_button, 'go-previous', QStyle.SP_ArrowBack),
            (self.forward_button, 'go-next', QStyle.SP_ArrowForward),
            (self.add_tab_button, 'tab-new', QStyle.SP_FileDialogNewFolder),
            (self.reopen_tab_button, 'edit-undo', QStyle.SP_BrowserReload),
            (self.search_button, 'edit-find', QStyle.SP_FileDialogContentsView),
            (self.insert_group_btn, 'view-group', QStyle.SP_DirOpenIcon),
            (self.split_view_btn, 'view-split-left-right', QStyle.SP_TitleBarNormalButton),
            (bookmark_btn, 'user-bookmarks', QStyle.SP_DirIcon),
            (settings_btn, 'preferences-system', QStyle.SP_FileDialogInfoView),
            (self.ai_chat_btn, 'help-contents', QStyle.SP_MessageBoxInformation),
        )
        toolbar_style = (
            'QPushButton, QToolButton { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 0; }'
            'QPushButton:hover:!disabled, QToolButton:hover:!disabled { background: #e5e9ee; }'
            'QPushButton:checked, QToolButton:checked { background: #dceafa; border-color: #739bca; }'
            'QPushButton:focus, QToolButton:focus { border-color: #2f6fdb; }'
        )
        for button, theme, fallback in icons:
            button.setText('')
            button.setFixedSize(btn_size, btn_size)
            _set_tool_icon(button, theme, fallback, max(16, int(18 * self.dpi_scale)))
            _theme.bind_style(button, toolbar_style)
        for button in (self.workspace_tools_button, self.file_tasks_button, self.tab_list_button):
            _theme.bind_style(button, toolbar_style)
            button.setAccessibleName(button.toolTip())
        self.bookmark_button = bookmark_btn
        self.settings_button = settings_btn
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

                # WM_SETTINGCHANGE：可能是 Windows 深浅色切换，去抖后重新读取
                if msg.message == 0x001A:
                    self._schedule_system_theme_check()
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

    def _apply_window_palettes(self):
        """主窗口与主容器的窗口底色（填充边框与内容之间的缝隙）随主题变化。"""
        from PyQt5.QtGui import QColor, QPalette
        for widget in (self, getattr(self, '_main_container', None)):
            if widget is not None:
                palette = widget.palette()
                palette.setColor(QPalette.Window, QColor(_theme.bg('#ffffff')))
                widget.setPalette(palette)

    def _schedule_system_theme_check(self):
        if _theme.normalize_mode(self.config.get("theme")) != 'system':
            return
        timer = getattr(self, '_system_theme_timer', None)
        if timer is None:
            timer = self._system_theme_timer = QTimer(self)
            timer.setSingleShot(True)
            timer.timeout.connect(self.apply_theme_config)
        timer.start(400)

    def apply_theme_config(self):
        """按配置（跟随系统/浅色/深色）应用主题；与当前一致时不做任何事。"""
        if not _theme.set_dark(_theme.resolve_dark(self.config.get("theme", "system"))):
            return False
        self._apply_window_palettes()
        self.apply_tab_group_markers_config()
        chat_panel = getattr(self, 'chat_panel', None)
        if chat_panel is not None:
            chat_panel.apply_theme()
        for _tabs, stack in self._all_groups():
            for index in range(stack.count()):
                pane = stack.widget(index)
                if hasattr(pane, '_git_status_cache'):
                    pane._git_status_cache = None  # 缓存的 Git 摘要带旧配色
        for pane in (self.get_current_tab_widget(), self.get_active_pane()):
            if pane is not None and hasattr(pane, 'update_explorer_status'):
                pane.update_explorer_status()
        self._update_resource_usage_display()
        self._rebuild_shell_views_for_theme()
        return True

    def _rebuild_shell_views_for_theme(self):
        """重建配色与当前主题不符的资源管理器视图；有 Shell 文件操作进度窗口时稍后再试，避免打断复制。"""
        if _shell_file_operation_window_open():
            QTimer.singleShot(2000, self._rebuild_shell_views_for_theme)
            return
        for _tabs, stack in self._all_groups():
            for index in range(stack.count()):
                pane = stack.widget(index)
                if hasattr(pane, 'rebuild_shell_view'):
                    pane.rebuild_shell_view()

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

        def _has_selection(tab):
            if not tab:
                return False
            try:
                if hasattr(tab, '_get_selected_paths'):
                    return bool(tab._get_selected_paths())
            except Exception:
                return False
            return False

        # 某些焦点切换场景下 active pane 可能晚于真实选中区域，
        # 此时回退到当前可见标签页，避免误报“未选中文件”。
        if not _has_selection(current_tab):
            cur_tab = self.get_current_tab_widget()
            if cur_tab is not current_tab and _has_selection(cur_tab):
                current_tab = cur_tab

        names = []
        if current_tab and hasattr(current_tab, 'get_selected_filenames'):
            names = current_tab.get_selected_filenames()
        if (not names) and current_tab and hasattr(current_tab, '_get_selected_paths'):
            try:
                sel_paths = current_tab._get_selected_paths()
                sel_paths = [p for p in (sel_paths or []) if p]
                names = [os.path.basename(p.rstrip('\\/')) for p in sel_paths if os.path.basename(p.rstrip('\\/'))]
            except Exception:
                names = []
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

    def quick_copy_selected_items_for_background_paste(self):
        """Alt+C：复制当前选中项到 TabEx 内部剪贴板（用于 Alt+V 后台粘贴）。"""
        current_tab = self.get_active_pane()
        if not current_tab or not hasattr(current_tab, '_get_selected_paths'):
            show_toast(self, tr("提示"), tr("当前标签不支持快速复制"), level="warning")
            return

        paths = current_tab._get_selected_paths()
        paths = [p for p in (paths or []) if p and os.path.exists(p)]
        if not paths:
            show_toast(self, tr("提示"), tr("请先选择要复制的文件或文件夹"), level="warning")
            return

        self._quick_clipboard_paths = list(paths)
        self._quick_clipboard_mode = 'copy'
        key = _hotkey_text(self.config, 'quick_paste')
        show_toast(self, tr("复制"), tr("已复制 {} 项，按 {} 粘贴到当前目录").format(len(paths), key) if key
                   else tr("已复制 {} 项").format(len(paths)), level="info")

    def quick_paste_to_current_directory(self):
        """Alt+V：将 TabEx 内部剪贴板的内容后台复制到当前目录。"""
        paths = list(getattr(self, '_quick_clipboard_paths', []) or [])
        if not paths:
            key = _hotkey_text(self.config, 'quick_copy')
            show_toast(self, tr("提示"), tr("内部剪贴板为空，请先按 {} 复制").format(key) if key
                       else tr("内部剪贴板为空"), level="warning")
            return

        current_tab = self.get_active_pane()
        if not current_tab or not hasattr(current_tab, '_run_file_batch_op'):
            show_toast(self, tr("提示"), tr("当前标签不支持快速粘贴"), level="warning")
            return

        dst_dir = getattr(current_tab, 'current_path', '')
        if not dst_dir or dst_dir.startswith(('shell:', '::')):
            show_toast(self, tr("提示"), tr("当前目录无效，无法粘贴"), level="warning")
            return

        alive_paths = [p for p in paths if p and os.path.exists(p)]
        if not alive_paths:
            key = _hotkey_text(self.config, 'quick_copy')
            show_toast(self, tr("提示"), tr("源文件不存在，请重新按 {} 复制").format(key) if key
                       else tr("源文件不存在"), level="warning")
            return

        conflict_actions = {}
        # 慢盘上逐项探测会阻塞界面，沿用自动重命名
        if not current_tab._is_slow_path(dst_dir):
            conflicts = _find_copy_conflicts(alive_paths, dst_dir)
            if conflicts:
                dialog = CopyConflictDialog(conflicts, dst_dir, self)
                accepted = dialog.exec_() == QDialog.Accepted
                conflict_actions = dialog.actions()
                dialog.deleteLater()
                self._guard_shortcuts_after_modal()
                if not accepted:
                    return
                if all(conflict_actions.get(p) == 'skip' for p in alive_paths):
                    show_toast(self, tr("提示"), tr("所有项目均已跳过"), level="info")
                    return

        current_tab._run_file_batch_op('copy', alive_paths, dst_dir, conflict_actions=conflict_actions)

    def quick_delete_selected_items(self):
        """Alt+Delete：将选中项移入回收站。"""
        current_tab = self.get_active_pane()
        if not current_tab or not hasattr(current_tab, '_get_selected_paths'):
            show_toast(self, tr("提示"), tr("当前标签不支持快速删除"), level="warning")
            return
        selected_paths = current_tab._get_selected_paths()
        selected_paths = [p for p in (selected_paths or []) if p]
        if not selected_paths:
            show_toast(self, tr("提示"), tr("请先选择要删除的文件或文件夹"), level="warning")
            return
        if hasattr(current_tab, '_run_file_batch_op'):
            current_tab._run_file_batch_op('delete', selected_paths)

    def quick_cancel_background_file_operation(self):
        """Alt+Q：取消当前标签页正在执行的后台复制/删除任务。"""
        current_tab = self.get_active_pane()
        if not current_tab:
            current_tab = self.get_current_tab_widget()
        if not current_tab or not hasattr(current_tab, 'cancel_current_file_batch_op'):
            show_toast(self, tr("提示"), tr("当前标签不支持取消后台任务"), level="warning")
            return
        if not current_tab.cancel_current_file_batch_op():
            show_toast(self, tr("提示"), tr("当前没有可取消的后台复制/删除任务"), level="info")

    def quick_find_in_current_directory(self):
        """通过关键字快速检索当前目录下的文件或文件夹名，并在当前目录中选中目标。"""
        previous = getattr(self, '_quick_find_worker', None)
        if previous is not None and previous.isRunning():
            previous.requestInterruption()
            show_toast(self, tr("提示"), tr("正在取消上一次检索，请稍后重试"), level="info")
            return
        current_tab = self.get_active_pane()
        if not current_tab or not hasattr(current_tab, 'current_path'):
            show_toast(self, tr("提示"), tr("当前没有可用的标签页路径"), level="warning")
            return

        search_root = current_tab.current_path
        if not isinstance(search_root, str) or not search_root or search_root.startswith('shell:') or '::' in search_root:
            show_toast(self, tr("提示"), tr("当前路径不支持快捷检索"), level="warning")
            return
        keyword, ok = QInputDialog.getText(self, tr("快捷定位"), tr("请输入要检索的文件或文件夹关键字："))
        self._guard_shortcuts_after_modal()
        if not ok:
            return

        keyword = keyword.strip()
        if not keyword:
            show_toast(self, tr("提示"), tr("请输入搜索关键词"), level="warning")
            return

        import weakref
        from PyQt5.QtWidgets import QProgressDialog
        worker = QuickFindWorker(search_root, keyword, parent=self)
        worker.tab_ref = weakref.ref(current_tab)
        progress = QProgressDialog(tr("正在检索"), tr("取消"), 0, 0, self)
        progress.setWindowTitle(tr("快捷定位"))
        progress.setWindowModality(Qt.NonModal)
        progress.canceled.connect(worker.requestInterruption)
        worker.dialog = progress
        self._quick_find_worker = worker
        worker.completed.connect(self._quick_find_completed)
        worker.finished.connect(self._quick_find_finished)
        progress.show()
        worker.start()
        _retain_thread_until_finished(worker)

    def _quick_find_finished(self):
        worker = self.sender()
        worker.dialog.close()
        worker.dialog.deleteLater()
        if worker is getattr(self, '_quick_find_worker', None):
            self._quick_find_worker = None

    def _quick_find_completed(self, search_root, matched_paths, error):
        from PyQt5 import sip
        worker = self.sender()
        if worker is not getattr(self, '_quick_find_worker', None) or worker.isInterruptionRequested():
            return
        current_tab = worker.tab_ref()
        if (current_tab is None or sip.isdeleted(current_tab) or current_tab is not self.get_active_pane()
                or current_tab.current_path != search_root):
            return
        worker.dialog.hide()
        if error:
            show_toast(self, tr("错误"), tr("快捷检索失败: {}").format(error), level="error")
            return
        if not matched_paths:
            show_toast(self, tr("提示"), tr("未找到匹配项"), level="info")
            return

        selected_path = matched_paths[0]
        if len(matched_paths) > 1:
            picker = QuickFindResultsDialog(matched_paths, self)
            picker.setWindowTitle(f"选择匹配项（共 {len(matched_paths)} 项）")
            ok = picker.exec_()
            self._guard_shortcuts_after_modal()
            if not ok or not picker.selected_path:
                return
            selected_path = picker.selected_path

        if sip.isdeleted(current_tab) or current_tab.current_path != search_root or current_tab is not self.get_active_pane():
            return
        selected_name = os.path.basename(selected_path)
        current_tab.select_file_in_explorer(selected_name)
        show_toast(self, tr("快捷定位"), tr("已选中: {}").format(selected_name), level="info")
    
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
            for _vk in (0x5A, 0x58, 0x43, 0x56, 0x2E, 0x4C, 0x54, 0x41, 0x57, 0x46, 0x47,
                    0x44, 0x51, 0x09, 0x25, 0x27, 0x26, 0x28, 0x74):
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
    
    def _start_shortcut_listener(self):
        hook = _ShortcutKeyHook(self)
        hook.keyPressed.connect(self._on_shortcut_key)
        if hook.start():
            self._shortcut_key_hook = hook
            debug_print("[ShortcutHook] WH_KEYBOARD_LL installed, polling disabled")
            return
        hook.deleteLater()
        debug_print("[ShortcutHook] Keyboard hook unavailable, falling back to polling")
        self._shortcut_timer.start(SHORTCUT_POLL_ACTIVE_MS)

    def _stop_shortcut_listener(self):
        hook, self._shortcut_key_hook = getattr(self, '_shortcut_key_hook', None), None
        if hook is not None:
            hook.stop()

    def _shortcut_gate(self, wait_for_release):
        """返回 (是否屏蔽快捷键, 本进程是否在前台)。"""
        from PyQt5.QtWidgets import QApplication, QLineEdit, QTextEdit, QPlainTextEdit
        if isinstance(QApplication.focusWidget(), (QLineEdit, QTextEdit, QPlainTextEdit)):
            return True, False
        # 以前台窗口的进程归属判断：嵌入 Shell 视图持有焦点时 activeWindow() 可能为 None
        if _foreground_pid() != os.getpid():
            return True, False
        # 本进程其他 Qt 顶层窗口（如非模态搜索框）持有焦点时不响应
        active_win = QApplication.activeWindow()
        if active_win is not None and active_win is not self:
            return True, False
        if getattr(self, '_shortcut_wait_for_modifier_release', False):
            get_state = ctypes.windll.user32.GetAsyncKeyState
            if wait_for_release and any(get_state(vk) & 0x8000 for vk in (0x11, 0x10, 0x12)):
                return True, True
            self._shortcut_wait_for_modifier_release = False
        if time.monotonic() < getattr(self, '_shortcut_modal_guard_until', 0):
            return True, True
        return False, True

    def _on_shortcut_key(self, vk, ctrl, shift, alt):
        try:
            # 钩子只上报新的按下事件，不存在轮询的残留按键问题，无需等待修饰键松开
            blocked, _foreground = self._shortcut_gate(wait_for_release=False)
            if not blocked:
                self._run_shortcut(vk, ctrl, shift, alt)
        except Exception as e:
            debug_print(f"[ShortcutHook] dispatch error: {e}")

    def _check_shortcuts(self):
        """轮询兜底：仅在键盘钩子安装失败时运行。"""
        try:
            blocked, foreground = self._shortcut_gate(wait_for_release=True)
            interval = SHORTCUT_POLL_ACTIVE_MS if foreground else SHORTCUT_POLL_INACTIVE_MS
            if self._shortcut_timer.interval() != interval:
                self._shortcut_timer.setInterval(interval)
            if blocked:
                self._last_keys_state.clear()
                return
            get_state = ctypes.windll.user32.GetAsyncKeyState
            ctrl, shift, alt = (bool(get_state(vk) & 0x8000) for vk in (0x11, 0x10, 0x12))
            pressed = []
            for vk in getattr(self, '_shortcut_poll_vks', ()):
                down = bool(get_state(vk) & 0x8000)
                if down and not self._last_keys_state.get(vk, False):
                    pressed.append(vk)
                self._last_keys_state[vk] = down
            for vk in pressed:
                if self._run_shortcut(vk, ctrl, shift, alt):
                    return
        except Exception:
            pass

    def _cycle_main_tab(self, step):
        count = self.tab_widget.count()
        if count:
            self.tab_widget.setCurrentIndex((self.tab_widget.currentIndex() + step) % count)

    def _select_tab_by_number(self, number):
        """Ctrl+1..8 切到活动标签组第 N 个标签，Ctrl+9 切到最后一个。"""
        tabs = self.get_active_group_tabwidget()
        count = tabs.count()
        index = count - 1 if number == 9 else number - 1
        if 0 <= index < count:
            tabs.setCurrentIndex(index)

    def _focus_active_path_bar(self):
        current_tab = self.get_active_pane()
        if current_tab and getattr(current_tab, 'path_bar', None):
            current_tab.path_bar.enter_edit_mode()

    def _run_shortcut(self, vk, ctrl, shift, alt):
        """执行按键组合绑定的命令，返回是否命中已启用的快捷键。"""
        command = _hotkey_table(self.config).get((ctrl, shift, alt, vk))
        # Ctrl+Shift+数字 是 Explorer 的视图切换，保留给系统
        if command is None and ctrl and not shift and not alt and 0x31 <= vk <= 0x39:
            command = 'switch_tab_number'
        if command is None:
            return False
        enable_key = _HOTKEY_ENABLE_KEYS.get(command, command)
        if enable_key is not None and not self.config.get("hotkeys", {}).get(enable_key, True):
            return False
        handlers = {
            'new_tab': lambda: self.add_new_tab(),
            'close_tab': self.close_current_tab,
            'reopen_tab': self.reopen_closed_tab,
            'next_tab': lambda: self._cycle_main_tab(1),
            'prev_tab': lambda: self._cycle_main_tab(-1),
            'switch_tab_number': lambda: self._select_tab_by_number(vk - 0x30),
            'search': self.show_search_dialog,
            'quick_find_current_dir': self.quick_find_in_current_directory,
            'go_back': self.go_back_current_tab,
            'go_forward': self.go_forward_current_tab,
            'go_up': self.go_up_current_tab,
            'refresh': self.refresh_current_tab,
            'add_bookmark': self.add_current_tab_bookmark,
            'quick_copy': self.quick_copy_selected_items_for_background_paste,
            'quick_paste': self.quick_paste_to_current_directory,
            'quick_delete': self.quick_delete_selected_items,
            'cancel_file_op': self.quick_cancel_background_file_operation,
            'copy_filename': lambda: self.copy_selected_filename(mode="filename"),
            'copy_filepath': lambda: self.copy_selected_filename(mode="path"),
            'split_view': self.toggle_split_view,
            'insert_group_bookmark': self.insert_tab_group_marker,
            'focus_path_bar': self._focus_active_path_bar,
            'toggle_ai_panel': self.toggle_chat_panel,
        }
        debug_print(f"[Shortcut] vk=0x{vk:02X} ctrl={ctrl} shift={shift} alt={alt} -> {command}")
        handlers[command]()
        return True

    def _apply_hotkey_bindings(self):
        """按当前配置更新轮询兜底的键位、Shell 键盘过滤器和界面上的按键提示。"""
        table = _hotkey_table(self.config)
        poll_vks = {vk for _ctrl, _shift, _alt, vk in table}
        if self.config.get("hotkeys", {}).get("switch_tab_number", True):
            poll_vks.update(range(0x31, 0x3A))
        self._shortcut_poll_vks = tuple(sorted(poll_vks))
        _ieb_keyboard_filter.set_tabex_hotkeys(_hotkey_reserved_combos(self.config))
        for name in _TOOLBAR_HOTKEY_TOOLTIPS:
            button = getattr(self, name, None)
            if button is not None:
                text = _toolbar_hotkey_tooltip(self, name)
                button.setToolTip(text)
                button.setAccessibleName(text)
        for _tabs, stack in self._all_groups():
            for index in range(stack.count()):
                button = getattr(stack.widget(index), 'cancel_file_op_btn', None)
                if button is not None:
                    button.setToolTip(_hotkey_hint(self.config, tr("取消当前后台复制/删除"), 'cancel_file_op'))

    def preview_batch_rename(self):
        tab = self.get_active_pane()
        if tab is None:
            return
        paths = tab._get_selected_paths()
        if not paths:
            show_toast(self, tr("提示"), tr("未选择"), level="warning")
            return
        find_text, accepted = QInputDialog.getText(self, tr("批量重命名"), tr("查找文件名中的文本"))
        if not accepted:
            return
        replacement, accepted = QInputDialog.getText(self, tr("批量重命名"), tr("替换为"))
        if not accepted:
            return
        try:
            plan = _plan_batch_rename(paths, find_text, replacement)
            if not plan:
                show_toast(self, tr("提示"), tr("无变化"), level="info")
                return
            before = '\n'.join(plan) + '\n'
            after = '\n'.join(plan.values()) + '\n'
            if _confirm_file_preview(self, tr("批量重命名预览"), tr("确认重命名以下项目"), before, after):
                self.get_file_task_panel().start_task('rename', list(plan), rename_plan=plan)
                self.show_file_tasks()
        except Exception as error:
            show_toast(self, tr("重命名失败"), str(error), level="error")

    def show_directory_compare(self):
        left = getattr(self.get_current_tab_widget(), 'current_path', '')
        right = ''
        if getattr(self, '_split_active', False):
            right_tab = self.split_content_stack.currentWidget()
            right = getattr(right_tab, 'current_path', '')
        dialog = getattr(self, '_directory_compare_dialog', None)
        if dialog is None:
            dialog = DirectoryCompareDialog(self, left, right)
            self._directory_compare_dialog = dialog
        elif not dialog.running:
            dialog.paths[0].setText(left)
            dialog.paths[1].setText(right)
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _capture_named_workspace(self):
        return SessionController(self)._capture_named_workspace()

    def save_named_workspace(self):
        from PyQt5.QtWidgets import QMessageBox
        name, accepted = QInputDialog.getText(self, tr("保存工作区"), tr("名称"))
        name = name.strip()
        if not accepted or not name:
            return
        workspaces = self.config.setdefault('named_workspaces', {})
        if name in workspaces and QMessageBox.question(
                self, tr("覆盖工作区"), name, QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No) != QMessageBox.Yes:
            return
        workspaces[name] = self._capture_named_workspace()
        self.save_config(immediate=True)
        show_toast(self, tr("工作区"), tr("已保存"), level="success")

    def _choose_named_workspace(self):
        workspaces = self.config.get('named_workspaces', {})
        if not workspaces:
            show_toast(self, tr("工作区"), tr("尚未保存工作区"), level="info")
            return None
        name, accepted = QInputDialog.getItem(self, tr("工作区"), tr("名称"), sorted(workspaces), 0, False)
        return name if accepted else None

    def open_named_workspace(self):
        name = self._choose_named_workspace()
        if name is None:
            return
        state = self.config['named_workspaces'][name]
        groups = state.get('groups', [])
        if not isinstance(groups, list) or not groups:
            return
        failures = []
        for group_index, group in enumerate(groups[:2]):
            if group_index == 1 and not getattr(self, '_split_active', False):
                self._activate_split_layout()
            target = self.tab_widget if group_index == 0 else self.split_tab_widget
            stack = self._content_stack_for(target)
            opened = []
            for entry in group.get('tabs', []):
                path = entry.get('path', '')
                if not isinstance(path, str) or not path:
                    continue
                index = self.add_new_tab(
                    path, is_shell=bool(entry.get('is_shell')), target_tabwidget=target,
                    activate=False, bookmark_group_color=entry.get('bookmark_group_color', ''),
                    tab_group_separator_after=entry.get('tab_group_separator_after', False),
                    tab_group_separator_color=entry.get('tab_group_separator_color', ''),
                    tab_group_separator_name=entry.get('tab_group_separator_name', ''))
                opened.append(index)
                if index < 0:
                    failures.append(path)
                    continue
                tab = stack.widget(index)
                tab.is_pinned = bool(entry.get('is_pinned', False))
                tab.update_tab_title()
            selected = int(group.get('active_index', 0))
            if opened:
                selected = max(0, min(selected, len(opened) - 1))
                if opened[selected] >= 0:
                    target.setCurrentIndex(opened[selected])
            self._apply_tab_grouping_for_pane(target)
        self.save_pinned_tabs()
        self._schedule_session_snapshot()
        if failures:
            show_toast(self, tr("工作区"), tr("部分目录无法打开") + ': ' + '\n'.join(failures), level="warning")

    def delete_named_workspace(self):
        from PyQt5.QtWidgets import QMessageBox
        name = self._choose_named_workspace()
        if name is not None and QMessageBox.question(
                self, tr("删除工作区记录"), name, QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No) == QMessageBox.Yes:
            del self.config['named_workspaces'][name]
            self.save_config(immediate=True)

    def _populate_workspace_tools_menu(self, menu):
        from PyQt5.QtWidgets import QStyle
        menu.addSection(tr('文件操作'))
        self.file_tasks_action = menu.addAction(tr('文件任务'), self.show_file_tasks)
        menu.addAction(tr('目录差异比较...'), self.show_directory_compare)
        menu.addAction(tr('批量重命名预览...'), self.preview_batch_rename)
        workspace = menu.addMenu(tr('工作区'))
        workspace.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))
        workspace.addAction(tr('保存命名工作区...'), self.save_named_workspace)
        workspace.addAction(tr('打开命名工作区...'), self.open_named_workspace)
        workspace.addAction(tr('删除命名工作区...'), self.delete_named_workspace)
        menu.addSection(tr('危险操作'))
        danger = menu.addAction(tr('永久删除选中项...'), self.permanently_delete_selected)
        danger.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxWarning))
        menu.addSection(tr('诊断'))
        self.export_diagnostics_action = menu.addAction(tr('导出崩溃诊断包...'), self.export_diagnostics)
        menu.addAction(tr('检查更新...'), lambda: self.check_for_updates(manual=True))

    def _update_task_indicator(self):
        from PyQt5.QtWidgets import QStyle
        panel = getattr(self, '_file_task_panel', None)
        running, failed = panel.task_counts() if panel is not None else (0, 0)
        summary = tr('文件任务：运行 {}，失败 {}').format(running, failed)
        button = getattr(self, 'file_tasks_button', None)
        if button is not None:
            _set_tool_icon(button, 'task-warning' if failed else 'file-tasks',
                           QStyle.SP_MessageBoxWarning if failed else QStyle.SP_FileDialogDetailedView)
            count = running if running else failed
            button.setText(str(count) if count < 100 else '99+')
            button.setToolTip(summary)
            button.setAccessibleName(summary)
        action = getattr(self, 'file_tasks_action', None)
        if action is not None:
            action.setText(summary if running or failed else tr('文件任务'))

    def get_file_task_panel(self):
        if getattr(self, '_file_task_panel', None) is None:
            self._file_task_panel = FileTaskPanel(self)
            self._file_task_panel.tasks_changed.connect(self._update_task_indicator)
        return self._file_task_panel

    def _capture_diagnostic_snapshot(self):
        if _debuglog._diagnostics is None:
            return
        from PyQt5.QtCore import QT_VERSION_STR, PYQT_VERSION_STR
        try:
            panel = getattr(self, '_file_task_panel', None)
            records = panel.records if panel is not None else []
            tasks = []
            for record in records[-50:]:
                worker = record['worker']
                tasks.append({
                    'operation': record['op'], 'done': record['done'],
                    'sources': len(record['paths']), 'failures': len(record['failed']),
                    'bytes_done': worker.bytes_done if worker is not None else None,
                    'bytes_total': worker.bytes_total if worker is not None else None,
                    'cancel_requested': worker._cancel_requested if worker is not None else False,
                })
            retained = []
            for worker in list(_retained_background_threads):
                try:
                    retained.append({'type': type(worker).__name__, 'running': worker.isRunning()})
                except RuntimeError:
                    pass
            groups, hibernated = [], 0
            for _tabs, stack in self._all_groups():
                groups.append(stack.count())
                hibernated += sum(bool(getattr(stack.widget(index), '_hibernated', False))
                                  for index in range(stack.count()))
            rss_mb = get_process_memory_usage_mb()
            _debuglog._diagnostics.snapshot({
                'qt': QT_VERSION_STR, 'pyqt': PYQT_VERSION_STR,
                'tabs_per_group': groups, 'split': bool(getattr(self, '_split_active', False)),
                'python_threads': threading.active_count(), 'rss_mb': rss_mb,
                'retained_threads': retained[:100], 'tasks': tasks, 'task_count': len(records),
                'search_windows': len(getattr(self, 'search_dialogs', []) or []),
            })
            now = time.monotonic()
            if now - getattr(self, '_last_resource_sample', -RESOURCE_SAMPLE_INTERVAL_S) >= RESOURCE_SAMPLE_INTERVAL_S:
                self._last_resource_sample = now
                _debuglog._diagnostics.sample({
                    'rss_mb': rss_mb, 'python_threads': threading.active_count(),
                    'thread_names': _thread_name_summary(),
                    'tabs': sum(groups), 'hibernated_tabs': hibernated,
                    'running_background_threads': sum(1 for item in retained if item['running']),
                    'running_tasks': sum(1 for record in records if not record['done']),
                    'search_windows': len(getattr(self, 'search_dialogs', []) or []),
                    'toasts': len(_active_toasts),
                })
        except Exception as error:
            _debuglog._diagnostics.record('snapshot_failed', type(error).__name__)

    def export_diagnostics(self):
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        if _debuglog._diagnostics is None:
            QMessageBox.warning(self, tr("诊断包"), tr("诊断记录不可用，请检查本地目录写入权限。"))
            return
        if getattr(self, '_diagnostic_export_worker', None) is not None:
            return
        text = tr("诊断包包含最近运行的版本、异常调用栈、任务状态和筛选后的日志。\n"
                  "不收集配置、聊天、书签或文件内容；路径和常见凭据会自动脱敏。\n"
                  "自动脱敏无法保证识别所有敏感信息，请检查 ZIP 内容后再分享。\n"
                  "仅保存到本地，不会自动上传。是否继续？")
        if QMessageBox.question(self, tr("导出崩溃诊断包"), text,
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        destination, _selected = QFileDialog.getSaveFileName(
            self, tr("导出崩溃诊断包"), f'TabEx-diagnostics-{time.strftime("%Y%m%d-%H%M%S")}.zip',
            'ZIP (*.zip)')
        if not destination:
            return
        if not destination.lower().endswith('.zip'):
            destination += '.zip'
            if os.path.exists(destination):
                if QMessageBox.question(self, tr("覆盖文件"), destination,
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
                    return
        self._capture_diagnostic_snapshot()
        worker = DiagnosticExportWorker(_debuglog._diagnostics, destination, _DEBUG_LOG_PATH)
        self._diagnostic_export_worker = worker
        self.export_diagnostics_action.setEnabled(False)
        worker.finished.connect(self._diagnostic_export_finished)
        worker.start()
        _retain_thread_until_finished(worker)

    def _diagnostic_export_finished(self):
        from PyQt5.QtWidgets import QMessageBox
        worker = self._diagnostic_export_worker
        self._diagnostic_export_worker = None
        self.export_diagnostics_action.setEnabled(True)
        if worker.error:
            QMessageBox.warning(self, tr("诊断包导出失败"), worker.error)
        else:
            destination = worker.destination
            show_toast(self, tr("诊断包已保存"), destination, level='success',
                       action_text=tr('打开位置'), action=lambda: self.add_new_tab(os.path.dirname(destination)))

    def show_file_tasks(self):
        panel = self.get_file_task_panel()
        panel.show()
        panel.raise_()
        panel.activateWindow()

    def permanently_delete_selected(self):
        from PyQt5.QtWidgets import QMessageBox
        tab = self.get_active_pane()
        if tab is None:
            return
        paths = tab._get_selected_paths()
        if not paths:
            return
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle(tr("永久删除"))
        box.setText(tr("选中的文件将永久删除，不进入回收站，无法撤销。是否继续？"))
        box.setDetailedText('\n'.join(paths))
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.setDefaultButton(QMessageBox.No)
        if box.exec_() == QMessageBox.Yes:
            tab._run_file_batch_op('permanent_delete', paths)

    def show_settings_menu(self):
        """显示设置对话框"""
        self.settings_dialog = SettingsDialog(self.config, self)
        dlg = self.settings_dialog
        result = dlg.exec_()
        if result:
            self.apply_theme_config()
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
            self.config["file_op_max_workers"] = dlg.file_op_workers_spin.value()
            self.config["show_tab_group_markers"] = dlg.show_tab_group_markers_cb.isChecked()

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
            self.config["hotkeys"]["switch_tab_number"] = dlg.hotkey_switch_tab_number.isChecked()
            self.config["hotkeys"]["search"] = dlg.hotkey_search.isChecked()
            self.config["hotkeys"]["navigate"] = dlg.hotkey_navigate.isChecked()
            self.config["hotkeys"]["go_up"] = dlg.hotkey_go_up.isChecked()
            self.config["hotkeys"]["refresh"] = dlg.hotkey_refresh.isChecked()
            self.config["hotkeys"]["add_bookmark"] = dlg.hotkey_add_bookmark.isChecked()
            self.config["hotkeys"]["quick_copy"] = dlg.hotkey_quick_copy.isChecked()
            self.config["hotkeys"]["quick_paste"] = dlg.hotkey_quick_paste.isChecked()
            self.config["hotkeys"]["quick_delete"] = dlg.hotkey_quick_delete.isChecked()
            self.config["hotkeys"]["cancel_file_op"] = dlg.hotkey_cancel_file_op.isChecked()
            self.config["hotkeys"]["copy_filename"] = dlg.hotkey_copy_filename.isChecked()
            self.config["hotkeys"]["copy_filepath"] = dlg.hotkey_copy_filepath.isChecked()
            self.config["hotkeys"]["quick_find_current_dir"] = dlg.hotkey_quick_find_current_dir.isChecked()
            self.config["hotkeys"]["split_view"] = dlg.hotkey_split_view.isChecked()
            self.config["hotkeys"]["insert_group_bookmark"] = dlg.hotkey_insert_group_bookmark.isChecked()
            self._apply_hotkey_bindings()
            
            self.save_config()

            # 同步标题栏按钮可见性
            self.apply_tortoisegit_buttons_config()
            self.apply_title_shortcuts_config()
            self.apply_tab_group_markers_config()
            
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
    
    def check_for_updates(self, manual=True):
        """后台查询 GitHub 最新发布；只提示，不下载。"""
        if getattr(self, '_update_worker', None) is not None:
            return
        worker = UpdateCheckWorker(manual)
        worker.completed.connect(self._on_update_check_finished)
        self._update_worker = worker
        worker.start()
        _retain_thread_until_finished(worker)
        if manual:
            show_toast(self, tr("检查更新"), tr("正在查询 GitHub 最新版本…"), level="info", duration=2500)

    def _maybe_auto_check_updates(self):
        if not self.config.get("auto_update_check", False):
            return
        try:
            last_check = float(self.config.get("last_update_check", 0) or 0)
        except (TypeError, ValueError):
            last_check = 0.0
        if time.time() - last_check >= UPDATE_CHECK_INTERVAL_S:
            self.check_for_updates(manual=False)

    def _on_update_check_finished(self, result):
        self._update_worker = None
        manual = bool(result.get('manual'))
        if result.get('error'):
            debug_print(f"[Update] check failed: {result['error']}")
            if manual:
                show_toast(self, tr("检查更新失败"), result['error'], level="warning",
                           action_text=tr("打开发布页"), action=lambda: _open_release_page(UPDATE_RELEASES_PAGE))
            return
        self.config["last_update_check"] = int(time.time())
        version = result.get('version', '')
        url = result.get('url', UPDATE_RELEASES_PAGE)
        remote, local = _version_tuple(version), _version_tuple(APP_VERSION)
        if remote and local and remote > local:
            # 自动检查对同一版本只提醒一次
            if manual or self.config.get("update_notified_version") != version:
                self.config["update_notified_version"] = version
                show_toast(self, tr("发现新版本"), tr("TabEx v{} 已发布，当前为 v{}").format(version, APP_VERSION),
                           level="info", duration=15000, action_text=tr("查看更新"),
                           action=lambda: _open_release_page(url))
        elif manual and remote:
            show_toast(self, tr("检查更新"), tr("当前已是最新版本 v{}").format(APP_VERSION), level="success")
        elif manual:
            show_toast(self, tr("检查更新"), tr("最新发布为 {}，无法读取其版本号").format(result.get('tag', '')),
                       level="info", action_text=tr("查看更新"), action=lambda: _open_release_page(url))
        self.save_config()
    
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
        bar = tw.tabBar() if tw is not None else None
        if bar is not None and hasattr(bar, 'consume_context_menu_suppression'):
            try:
                if bar.consume_context_menu_suppression():
                    return
            except Exception:
                pass
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
        tab.tab_group_separator_after = False
        tab.tab_group_separator_color = ""
        tab.tab_group_separator_name = ""
        tab.bookmark_group_color = ""
        # 重新排序：所有固定的在最左侧（仅作用于该标签组）
        self.sort_tabs_by_pinned(tw)
        self.save_pinned_tabs()

    def unpin_tab(self, tab_index, target_tabwidget=None):
        tw, cs, _is_right = self._resolve_group(target_tabwidget)
        tab = cs.widget(tab_index) if cs is not None else None
        if tab is None:
            return
        tab.is_pinned = False
        move_to_end = tab if not str(getattr(tab, 'bookmark_group_color', '') or '').strip() else None
        self.sort_tabs_by_pinned(tw, move_to_end=move_to_end)
        self.save_pinned_tabs()

    def sort_tabs_by_pinned(self, target_tabwidget=None, move_to_end=None):
        tw, cs, _is_right = self._resolve_group(target_tabwidget)
        if cs is None:
            return
        pinned = []
        unpinned = []
        current_index = tw.currentIndex()
        current_tab = cs.widget(current_index) if current_index >= 0 else None
        for i in range(tw.count()):
            tab = cs.widget(i)
            if hasattr(tab, 'is_pinned') and tab.is_pinned:
                pinned.append(tab)
            else:
                unpinned.append(tab)
        if move_to_end in unpinned:
            unpinned.remove(move_to_end)
            unpinned.append(move_to_end)
        new_tabs = pinned + unpinned
        tabbar = tw.tabBar()
        was_sorting = getattr(tabbar, '_sorting_pinned_tabs', False)
        tabbar._sorting_pinned_tabs = True
        blocked = [(control, control.blockSignals(True)) for control in (tw, cs)]
        try:
            for index, tab in enumerate(new_tabs):
                previous_index = cs.indexOf(tab)
                if previous_index != index:
                    cs.removeWidget(tab)
                    cs.insertWidget(index, tab)
                    tabbar.moveTab(previous_index, index)
            if current_tab is not None:
                tw.setCurrentIndex(cs.indexOf(current_tab))
                cs.setCurrentWidget(current_tab)
        finally:
            tabbar._sorting_pinned_tabs = was_sorting
            for control, was_blocked in reversed(blocked):
                control.blockSignals(was_blocked)
        self._apply_tab_grouping_for_pane(tw)
        self._refresh_tab_labels()
        if hasattr(self, '_on_group_tab_changed'):
            self._on_group_tab_changed(tw, tw.currentIndex())

    def save_pinned_tabs(self):
        """保存固定标签页到config.json（扫描左右两个标签组）"""
        pinned_paths = []
        for _tw, cs in self._all_groups():
            if cs is None:
                continue
            for i in range(cs.count()):
                tab = cs.widget(i)
                if tab and getattr(tab, 'is_pinned', False) and hasattr(tab, 'current_path'):
                    current_path = getattr(tab, 'current_path', '')
                    if not current_path:
                        continue
                    pinned_paths.append(current_path)

        # 更新config并保存
        self.config["pinned_tabs"] = pinned_paths
        self.save_config()
        self._schedule_session_snapshot()
        
        print(f"[Config] Saved {len(pinned_paths)} pinned tabs to config.json")

    def load_pinned_tabs(self):
        """从config.json加载固定标签页"""
        has_pinned = False
        
        # 从config.json读取
        pinned_entries = self.config.get("pinned_tabs", [])
        
        if pinned_entries:
            print(f"[Config] Loading {len(pinned_entries)} pinned tabs from config.json")
            for entry in pinned_entries:
                if isinstance(entry, dict):
                    path = entry.get('path', '')
                    is_shell = bool(entry.get('is_shell', False)) or str(path).startswith('shell:')
                else:
                    path = entry
                    is_shell = str(path).startswith('shell:')
                if path.startswith('shell:') or _tab_target_available(path):
                    try:
                        # 懒加载：固定标签也延迟首次导航，避免启动瞬间多个 Shell 视图同时创建
                        tab = FileExplorerTab(self, path, is_shell=is_shell, defer_nav=True)
                        if hasattr(tab, 'set_bottom_statusbar_visible'):
                            tab.set_bottom_statusbar_visible(self.config.get("show_bottom_statusbar", True))
                        tab.is_pinned = True
                        tab.tab_group_separator_after = False
                        tab.tab_group_separator_color = ""
                        tab.tab_group_separator_name = ""
                        tab.bookmark_group_color = ""
                        # 同时添加到 tab_widget 和 content_stack
                        self.tab_widget.addTab(QWidget(), "")  # tab_widget 只显示标签，内容用占位widget
                        self.content_stack.addWidget(tab)  # 实际内容添加到 content_stack
                        try:
                            tab.update_tab_title()
                        except Exception:
                            pass
                        has_pinned = True
                        print(f"[Config] ✓ Loaded pinned tab: {path}")
                    except Exception as e:
                        print(f"[Config] ✗ Failed to load pinned tab {path}: {e}")
                else:
                    print(f"[Config] ⚠ Skipping non-existent path: {path}")
        else:
            print("[Config] No pinned tabs found in config.json")

        self._apply_tab_grouping_for_pane(self.tab_widget)
        
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
        self._quick_clipboard_paths = []
        self._quick_clipboard_mode = 'copy'
        
        # 启用主窗口拖拽支持
        self.setAcceptDrops(True)
        
        # 加载配置
        self.config = self.load_config()
        
        # 初始化全局调试开关
        set_debug_mode(self.config.get("debug_mode", False))
        set_explorer_monitor_debug(self.config.get("explorer_monitor_debug", False))
        # 须在创建任何 Shell 视图之前应用：原生视图只在新建时完整采用深色
        _theme.set_dark(_theme.resolve_dark(self.config.get("theme", "system")))
        
        # 初始化书签管理器
        self.bookmark_manager = BookmarkManager()
        self._last_bookmark_node_id = None
        self._bookmark_effective_colors = {}
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
        
        # 快捷键：优先键盘钩子（事件驱动）；安装失败时退回定时轮询
        self._last_keys_state = {}
        self._shortcut_key_hook = None
        self._shortcut_timer = QTimer(self)
        self._shortcut_timer.timeout.connect(self._check_shortcuts)
        self._apply_hotkey_bindings()
        self._start_shortcut_listener()

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
        self._housekeeping_timer.start(HOUSEKEEPING_INTERVAL_MS)
        self._tab_hibernate_timer = QTimer(self)
        self._tab_hibernate_timer.timeout.connect(self._hibernate_idle_tabs)
        self._tab_hibernate_timer.start(60 * 1000)
        self._update_check_timer = QTimer(self)
        self._update_check_timer.timeout.connect(self._maybe_auto_check_updates)
        self._update_check_timer.start(60 * 60 * 1000)
        QTimer.singleShot(UPDATE_CHECK_DELAY_MS, self._maybe_auto_check_updates)

        # 状态栏右侧 CPU/内存占用显示（默认关闭，可在设置中开启）
        self._resource_usage_timer = QTimer(self)
        self._resource_usage_timer.timeout.connect(self._update_resource_usage_display)
        self.apply_bottom_statusbar_config()
        self.apply_resource_usage_config()

        
        # 性能优化：延迟加载非关键功能（100ms后加载）
        QTimer.singleShot(100, self._delayed_initialization)

        # 生成并设置 TE 窗口图标：延迟到事件循环启动后执行，
        # 避免 9 张抗锯齿图标（256→16px）的渲染阻塞首帧显示，加快启动感知速度。
        QTimer.singleShot(0, self._setup_window_icon)

    def _setup_window_icon(self):
        """生成并设置 TE 窗口图标（延迟执行，不阻塞启动首帧）。"""
        try:
            te_icon = _build_te_icon()
            self.setWindowIcon(te_icon)
            # 应用级（任务栏）图标已在 main() 早期、主窗口书签菜单创建之前设置。
            # 不在此处调用 QApplication.setWindowIcon：菜单已存在时该调用会向所有顶层
            # 弹出菜单广播图标变更事件；英文界面书签栏溢出时会在弹出菜单间无限递归
            # 传播（WindowIconChange/ApplicationWindowIconChange），导致栈溢出崩溃。
            if hasattr(self, '_title_icon_label'):
                self._title_icon_label.setPixmap(te_icon.pixmap(18, 18))
            print("[Icon] ✓ TE icon set on window")
        except Exception as e:
            print(f"[Icon] ✗ Failed to set TE icon: {e}")
    
    def load_config(self):
        self._config_store = ConfigStore(get_app_data_path("config.json"))
        config = self._config_store.load()
        self._config_recovered_backup = self._config_store.recovered_backup
        return config
    
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
        store = getattr(self, "_config_store", None)
        if store is None:
            store = self._config_store = ConfigStore(get_app_data_path("config.json"))
        result = store.save(self.config)
        self._last_config_written = store.last_written
        if not result and isinstance(self, QWidget) and not getattr(self, '_config_save_warning_shown', False):
            self._config_save_warning_shown = True
            show_toast(self, tr("配置保存失败"), tr("旧配置已保留，请检查目录权限或磁盘空间"), level="warning")
        elif result:
            self._config_save_warning_shown = False
        return result

    def _get_pinned_paths_from_config(self):
        return SessionController(self)._get_pinned_paths_from_config()

    def _collect_cached_tabs(self):
        return SessionController(self)._collect_cached_tabs()

    def _collect_split_session(self):
        return SessionController(self)._collect_split_session()

    def _get_last_active_tab_path(self):
        return SessionController(self)._get_last_active_tab_path()

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
        # 原地修改：该列表由提示模块与主窗口共享
        _active_toasts[:] = alive_toasts[:MAX_ACTIVE_TOASTS]
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
        enabled = self.config.get("show_resource_usage_in_statusbar", False) and \
            self.config.get("show_bottom_statusbar", True)
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

    def apply_bottom_statusbar_config(self):
        """根据配置显示/隐藏所有标签页底部状态栏。"""
        visible = bool(self.config.get("show_bottom_statusbar", True))
        for _tw, cs in self._all_groups():
            if cs is None:
                continue
            for i in range(cs.count()):
                tab = cs.widget(i)
                if tab and hasattr(tab, 'set_bottom_statusbar_visible'):
                    tab.set_bottom_statusbar_visible(visible)

    def _update_resource_usage_display(self):
        """显示整机 CPU/内存占用到当前活动标签的资源标签上，占用过高时变色预警。"""
        if (not self.config.get("show_resource_usage_in_statusbar", False) or
                not self.config.get("show_bottom_statusbar", True)):
            return

        def _color(pct):
            if pct >= RESOURCE_CRIT_PERCENT:
                return _theme.fg("#d32f2f")  # 红：危急
            if pct >= RESOURCE_WARN_PERCENT:
                return _theme.fg("#e67700")  # 橙：偏高
            return _theme.fg("#666")          # 常态

        cpu = get_system_cpu_percent()
        mem = get_system_memory_status()
        cpu_html = (f"<span style='color:{_color(cpu)}'>CPU {cpu:.0f}%</span>"
                    if cpu is not None else f"<span style='color:{_theme.fg('#666')}'>CPU --</span>")
        if mem is not None:
            used_mb, total_mb, pct = mem
            mem_html = (f"<span style='color:{_color(pct)}'>{tr('内存')} "
                        f"{used_mb/1024:.1f}/{total_mb/1024:.1f} GB ({pct}%)</span>")
        else:
            mem_html = ""
        text = f"{cpu_html}&nbsp;&nbsp;&nbsp;{mem_html}".strip()
        tab = self.get_current_tab_widget()
        lbl = getattr(tab, 'resource_label', None) if tab else None
        if lbl and getattr(tab, '_bottom_statusbar_visible', True):
            lbl.setText(text)
            lbl.show()

    def _restore_last_active_tab(self):
        return SessionController(self)._restore_last_active_tab()

    def _restore_split_session(self):
        return SessionController(self)._restore_split_session()

    def save_session_snapshot(self, immediate=False):
        return SessionController(self).save_session_snapshot(immediate)

    def _warn_recovered_data_files(self):
        """提示启动时无法解析、已改名保留的配置/书签文件。"""
        backups = (getattr(self, '_config_recovered_backup', ''),
                   getattr(getattr(self, 'bookmark_manager', None), 'recovered_backup', ''))
        for backup in filter(None, backups):
            name = os.path.basename(backup)
            show_toast(self, tr("数据文件无法读取"),
                       tr("{} 无法解析，原文件已另存为 {}，本次使用默认内容").format(name.split('.broken-')[0], name),
                       level="warning")

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
        prefer_en = (_i18n._app_language == 'en')

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
        # 再次用样式表确保非客户区也以白色填充
        _theme.bind_style(self, "QMainWindow { background: white; }")
        
        
        # 创建主容器，无边距，纯白填充
        main_container = QWidget()
        main_container.setAutoFillBackground(True)
        _theme.bind_style(main_container, "QWidget { background: white; margin: 0px; padding: 0px; border: none; }")
        container_layout = QVBoxLayout(main_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)
        
        # 保存主容器引用，用于应用阴影效果
        self._main_container = main_container
        self._apply_window_palettes()
        
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
        
        tab_css = f"""
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
                background: #f5f5f5;
                border: 1px solid #2F6FDB;
                border-bottom: none;
                margin-top: 0px;
                padding-top: {tab_padding_v + 1}px;
                color: #2F6FDB;
                font-weight: bold;
            }}
            QTabBar::tab:!selected {{
                font-weight: normal;
                margin-top: {tab_margin + 1}px;
            }}
            QTabBar[activePane="false"]::tab:selected {{
                border-color: #a0a5ad;
                color: #505050;
            }}
        """
        _theme.bind_style(tabbar, tab_css)
        # 设置标签文本省略模式 - 左边省略，保留右侧文件/文件夹名称
        tabbar.setElideMode(Qt.ElideLeft)

        # 创建标签栏容器（只显示标签和按钮，不显示内容）
        tab_bar_container = QWidget()
        tab_bar_height = int(32 * getattr(self, 'dpi_scale', 1.0))
        tab_bar_container.setFixedHeight(tab_bar_height)  # 固定高度，只显示标签栏
        _theme.bind_style(tab_bar_container, "background-color: #f3f3f3;")
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
        _theme.bind_style(split_tabbar, tab_css)
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
        _theme.bind_style(self.menu_bar, """
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
        _theme.bind_style(self.content_stack, "background: white;")
        self.content_stack.setAutoFillBackground(True)
        # 允许右侧内容在窗口缩小时被压缩，避免阻止左侧目录树向左拖动
        self.content_stack.setMinimumWidth(0)
        
        self.splitter.addWidget(self.content_stack)
        
        self.splitter.setCollapsible(0, False)  # content_stack 不允许折叠

        # 右侧分屏内容栈（默认不加入分割器，F3 时插入到索引 1）
        self.split_content_stack = QStackedWidget()
        _theme.bind_style(self.split_content_stack, "background: white;")
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
        if not path:
            self._bring_to_front()
            return
        pending = getattr(self, '_open_path_workers', None)
        if pending is None:
            pending = self._open_path_workers = set()
        if len(pending) >= 16:
            show_toast(self, tr("提示"), tr("路径请求过多，请稍后重试"), level="warning")
            return
        worker = OpenPathWorker(path, self)
        pending.add(worker)
        worker.completed.connect(self._open_instance_path_ready)
        worker.finished.connect(self._open_instance_path_finished)
        worker.start()
        _retain_thread_until_finished(worker)

    def _open_instance_path_finished(self):
        self._open_path_workers.discard(self.sender())

    def _open_instance_path_ready(self, path, error):
        if error:
            show_toast(self, tr("错误"), tr("无法打开路径: {}").format(error), level="warning")
            return
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
                    pinned_paths = self._get_pinned_paths_from_config()
                    pinned_norm = {
                        self._normalize_path_for_compare(p)
                        for p in pinned_paths if p
                    }
                    debug_print(f"[App] 恢复 {len(cached_tabs)} 个缓存标签页")
                    for tab_info in cached_tabs:
                        path = tab_info.get('path', '')
                        if path:
                            norm = self._normalize_path_for_compare(path)
                            if norm in pinned_norm:
                                debug_print(tr("[App] 跳过缓存标签（已固定）: {}").format(path))
                                continue
                            # 懒加载：恢复的缓存标签不逐个激活，后台标签首次可见时才导航
                            self.add_new_tab(
                                path,
                                activate=False,
                                bookmark_group_color=(
                                    tab_info.get('bookmark_group_color', '')
                                    or tab_info.get('tab_group_separator_color', '')
                                ),
                                tab_group_separator_after=tab_info.get('tab_group_separator_after', False),
                                tab_group_separator_color=tab_info.get('tab_group_separator_color', ''),
                                tab_group_separator_name=tab_info.get('tab_group_separator_name', ''),
                            )
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
        self._apply_tab_grouping_for_pane(self.tab_widget)
        self._apply_tab_grouping_for_pane(getattr(self, 'split_tab_widget', None))
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
        self._warn_recovered_data_files()
        debug_print("[Performance] Delayed initialization completed")
    
    def start_instance_server(self):
        """实例服务在创建窗口前取得配置写入权，由启动入口管理。"""
        coordinator = getattr(self, '_instance_coordinator', None)
        if coordinator is not None:
            coordinator.set_receiver(self.open_path_signal.emit)
    
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
        """Explorer窗口监听循环：优先响应窗口显示事件，钩子不可用时按间隔轮询。"""
        pythoncom_initialized = False
        show_hook = None
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
            show_hook = self._create_explorer_show_hook()
            # 事件驱动时仍保留低频全量扫描兜底
            fallback_interval = max(10.0, float(self.monitor_interval)) if show_hook else self.monitor_interval
            debug_print(f"[Explorer Monitor] Mode: {'event-driven' if show_hook else 'polling'}, "
                        f"scan interval: {fallback_interval}s")
            
            while self.explorer_monitoring:
                if show_hook:
                    triggered = self._wait_explorer_show_event(show_hook, fallback_interval)
                    if not self.explorer_monitoring:
                        break
                    if triggered:
                        # 给新窗口一点时间设置标题和位置
                        time.sleep(0.1)
                        show_hook['pending'].clear()
                else:
                    time.sleep(self.monitor_interval)
                    # 防抖：如果距离上次检查太近，跳过
                    if time.time() - self.last_check_time < self.monitor_interval * 0.8:
                        continue
                
                self.last_check_time = time.time()
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
                        # 先隐藏外部窗口，避免在读取路径和嵌入期间闪现；不接管时恢复显示
                        hidden = self._set_external_explorer_visible(hwnd, False)
                        handed_off = False
                        try:
                            path = self._get_explorer_path(hwnd)

                            if path:
                                debug_print(f"[Explorer Monitor] ✓ Path: {path}")

                                # 控制面板及其子目录直接在原窗口打开，不拦截
                                if self._is_control_panel_path_for_monitor(path):
                                    debug_print(f"[Explorer Monitor] Control Panel detected, keeping original window")
                                    continue

                                # 非控制面板路径，发送信号并关闭原窗口
                                self.open_path_signal.emit(path)
                                handed_off = True
                                time.sleep(0.2)
                                try:
                                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                                    debug_print(f"[Explorer Monitor] ✓ Closed original Explorer (hwnd={hwnd})")
                                except Exception as e:
                                    debug_print(f"[Explorer Monitor] ✗ Failed to close: {e}")
                            else:
                                debug_print(f"[Explorer Monitor] ✗ Could not get path from {hwnd}")
                        finally:
                            if hidden and not handed_off:
                                self._set_external_explorer_visible(hwnd, True)
                    
                    except Exception as e:
                        debug_print(f"[Explorer Monitor] Error processing window {hwnd}: {e}")
                
                # 更新已知窗口列表
                self.known_explorer_windows = current_explorer_windows
                
        except Exception as e:
            debug_print(f"[Explorer Monitor] Monitor loop error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if show_hook:
                show_hook['user32'].UnhookWinEvent(show_hook['handle'])
            self.explorer_monitoring = False
            self.known_explorer_windows.clear()
            self.monitor_thread = None
            if pythoncom_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
    
    def _create_explorer_show_hook(self):
        """在监听线程上订阅 EVENT_OBJECT_SHOW，开启 Explorer 顶层窗口时置位 pending。"""
        try:
            wt = ctypes.wintypes
            user32 = ctypes.WinDLL('user32', use_last_error=True)
            proc_type = ctypes.WINFUNCTYPE(None, wt.HANDLE, wt.DWORD, wt.HWND, wt.LONG, wt.LONG, wt.DWORD, wt.DWORD)
            user32.SetWinEventHook.argtypes = (wt.DWORD, wt.DWORD, wt.HMODULE, proc_type, wt.DWORD, wt.DWORD, wt.DWORD)
            user32.SetWinEventHook.restype = wt.HANDLE
            user32.UnhookWinEvent.argtypes = (wt.HANDLE,)
            user32.GetClassNameW.argtypes = (wt.HWND, wt.LPWSTR, ctypes.c_int)
            user32.MsgWaitForMultipleObjects.argtypes = (wt.DWORD, ctypes.c_void_p, wt.BOOL, wt.DWORD, wt.DWORD)
            user32.PeekMessageW.argtypes = (ctypes.POINTER(wt.MSG), wt.HWND, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint)
            pending = threading.Event()
            class_buffer = ctypes.create_unicode_buffer(32)

            def on_show(_hook, _event, hwnd, id_object, id_child, _thread, _time):
                if id_object == 0 and id_child == 0 and hwnd:
                    if user32.GetClassNameW(hwnd, class_buffer, 32) and class_buffer.value in ("CabinetWClass", "ExploreWClass"):
                        pending.set()

            callback = proc_type(on_show)
            # EVENT_OBJECT_SHOW；WINEVENT_OUTOFCONTEXT(0) | WINEVENT_SKIPOWNPROCESS(2)
            handle = user32.SetWinEventHook(0x8002, 0x8002, None, callback, 0, 0, 0x0002)
            if not handle:
                debug_print(f"[Explorer Monitor] SetWinEventHook failed err={ctypes.get_last_error()}")
                return None
            return {'user32': user32, 'handle': handle, 'callback': callback, 'pending': pending}
        except Exception as e:
            debug_print(f"[Explorer Monitor] Show hook unavailable: {e}")
            return None

    def _wait_explorer_show_event(self, show_hook, timeout_s):
        """泵消息等待窗口显示事件；返回 True 表示有新 Explorer 窗口。"""
        user32 = show_hook['user32']
        pending = show_hook['pending']
        msg = ctypes.wintypes.MSG()
        deadline = time.monotonic() + timeout_s
        while self.explorer_monitoring and not pending.is_set():
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return False
            # QS_ALLINPUT；分片等待以便及时响应停止请求
            user32.MsgWaitForMultipleObjects(0, None, False, int(min(remaining, 0.25) * 1000), 0x04FF)
            while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        return pending.is_set()

    @staticmethod
    def _set_external_explorer_visible(hwnd, visible):
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW if visible else win32con.SW_HIDE)
            return True
        except Exception as e:
            debug_print(f"[Explorer Monitor] ShowWindow({visible}) failed for {hwnd}: {e}")
            return False

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
                                return _explorer_location_to_path(location, location_name)
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
        chat = getattr(self, 'chat_panel', None)
        if chat is not None and chat.has_running_actions():
            chat.cancel_actions()
            event.ignore()
            show_toast(self, tr("提示"), tr("AI 文件操作正在结束，请稍后再退出"), level="warning")
            return
        if getattr(self, '_diagnostic_export_worker', None) is not None:
            event.ignore()
            show_toast(self, tr("提示"), tr("诊断包正在导出，请完成后再退出"), level="warning")
            return
        panel = getattr(self, '_file_task_panel', None)
        if panel is not None and panel.has_running_tasks():
            event.ignore()
            self.show_file_tasks()
            show_toast(self, tr("提示"), tr("仍有文件任务运行，请等待完成或取消任务后再退出"), level="warning")
            return
        comparison = getattr(self, '_directory_compare_dialog', None)
        if comparison is not None:
            comparison.close()
        quick_find = getattr(self, '_quick_find_worker', None)
        if quick_find is not None:
            quick_find.requestInterruption()
            self._quick_find_worker = None
        for worker in getattr(self, '_open_path_workers', ()):
            worker.requestInterruption()
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
        if getattr(self, 'server_socket', None) is not None:
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

        try:
            self._stop_shortcut_listener()
        except Exception as e:
            print(f"Error stopping shortcut hook: {e}")

        for timer_name in (
            '_shortcut_timer',
            '_session_snapshot_debounce_timer',
            '_session_snapshot_timer',
            '_housekeeping_timer',
            '_tab_hibernate_timer',
            '_update_check_timer',
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

    def open_bookmark_node(self, node):
        if not isinstance(node, dict):
            return
        self._last_bookmark_node_id = node.get('id')
        if self._is_group_separator_node(node):
            show_toast(self, tr("提示"), tr("该分组为分隔标记，不打开路径"), level="info")
            return
        group_color = self._bookmark_effective_colors.get(node.get('id'))
        self.open_bookmark_url(node.get('url', ''), group_color=group_color, bookmark_node_id=node.get('id'))

    def open_bookmark_url(self, url, group_color=None, bookmark_node_id=None):
        # 支持 file:///、file://、shell: 路径和本地绝对路径
        from urllib.parse import unquote
        target_tw = self.get_active_group_tabwidget()
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
                self.add_new_tab(local_path, is_shell=True, target_tabwidget=target_tw,
                                 bookmark_group_color=group_color, bookmark_source_node_id=bookmark_node_id)
            elif _tab_target_available(local_path):
                self.add_new_tab(local_path, target_tabwidget=target_tw,
                                 bookmark_group_color=group_color, bookmark_source_node_id=bookmark_node_id)
            else:
                show_toast(self, tr("路径错误"), tr("路径不存在: {}").format(local_path), level="warning")
        elif url.startswith('shell:'):
            # shell:OneDrive 解析为真实路径（避免Shell.Explorer无法正确显示内容）
            if url.lower() == 'shell:onedrive':
                onedrive_path = os.environ.get('OneDrive', '')
                if onedrive_path and os.path.exists(onedrive_path):
                    self.add_new_tab(onedrive_path, target_tabwidget=target_tw,
                                     bookmark_group_color=group_color, bookmark_source_node_id=bookmark_node_id)
                else:
                    show_toast(self, tr("路径错误"), tr("未找到 OneDrive 文件夹"), level="warning")
            else:
                self.add_new_tab(url, is_shell=True, target_tabwidget=target_tw,
                                 bookmark_group_color=group_color, bookmark_source_node_id=bookmark_node_id)
        elif os.path.isabs(url) and _tab_target_available(url):
            self.add_new_tab(url, target_tabwidget=target_tw,
                             bookmark_group_color=group_color, bookmark_source_node_id=bookmark_node_id)
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
        self._last_bookmark_node_id = bookmark_id
        node = self._find_bookmark_node_by_id(bookmark_id)
        is_group_separator = self._is_group_separator_node(node)
        menu = QMenu(self)

        insert_left_action = menu.addAction(tr("在左侧插入分组"))
        insert_right_action = menu.addAction(tr("在右侧插入分组"))

        if is_group_separator:
            menu.addSeparator()
            rename_group_action = menu.addAction(tr("重命名分组"))
            collapse_group_action = menu.addAction(
                tr("展开分组") if bool(node.get('group_collapsed', False)) else tr("折叠分组")
            )
            close_group_tabs_action = menu.addAction(tr("关闭该分组标签页"))
            keep_only_group_tabs_action = menu.addAction(tr("仅保留当前分组标签页"))
        else:
            rename_group_action = None
            collapse_group_action = None
            close_group_tabs_action = None
            keep_only_group_tabs_action = None

        menu.addSeparator()
        
        delete_action = menu.addAction(tr("🗑️ 删除书签"))
        insert_left_action.triggered.connect(lambda: self._insert_group_relative_to(bookmark_id, side='left'))
        insert_right_action.triggered.connect(lambda: self._insert_group_relative_to(bookmark_id, side='right'))
        if rename_group_action is not None:
            rename_group_action.triggered.connect(lambda: self.rename_group_bookmark(bookmark_id))
        if collapse_group_action is not None:
            collapse_group_action.triggered.connect(lambda: self.toggle_group_collapsed(bookmark_id))
        if close_group_tabs_action is not None:
            close_group_tabs_action.triggered.connect(lambda: self.close_group_tabs(bookmark_id))
        if keep_only_group_tabs_action is not None:
            keep_only_group_tabs_action.triggered.connect(lambda: self.keep_only_current_group_tabs(bookmark_id))
        delete_action.triggered.connect(lambda: self.confirm_delete_bookmark(bookmark_id, bookmark_name))
        
        debug_print(f"[DEBUG] Showing menu...")
        menu.exec_(pos)
        debug_print(f"[DEBUG] Menu closed")

    def _insert_group_relative_to(self, bookmark_id, side='right'):
        if not bookmark_id:
            return False
        self._last_bookmark_node_id = bookmark_id
        return self.insert_group_bookmark(side=side)
    
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

        children = bookmark_bar.get('children', [])
        self._bookmark_effective_colors = self._compute_effective_group_colors(children)
        group_member_map = self._count_group_member_map(children)
        open_count_map = self._count_open_tabs_by_group_color()
        hidden_top_level_ids = set()
        group_collapsed = False
        for node in reversed(children):
            if not isinstance(node, dict):
                continue
            if self._is_group_separator_node(node):
                group_collapsed = bool(node.get('group_collapsed', False))
                continue
            if group_collapsed:
                node_id = node.get('id')
                if node_id:
                    hidden_top_level_ids.add(node_id)
        
        # 存储action/menu到节点的映射
        self.bookmark_actions = {}
        self.bookmark_menus = {}  # 存储QMenu到节点的映射
        action_group_colors = {}
        
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
                is_group_separator = self._is_group_separator_node(node)
                if is_special:
                    action = parent_menu.addAction(name)
                elif is_group_separator:
                    prefix = "◨" if bool(node.get('group_collapsed', False)) else "◧"
                    member_count = int(group_member_map.get(node.get('id'), 0))
                    open_count = int(open_count_map.get(str(node.get('group_color', '')).strip().lower(), 0))
                    badge = tr("成员 {} | 开页 {}").format(member_count, open_count)
                    action = parent_menu.addAction(f"{prefix} {name} [{badge}]")
                else:
                    action = parent_menu.addAction(f"📑 {name}")
                action.triggered.connect(lambda checked, n=node: self.open_bookmark_node(n))
                # 存储action和节点的映射
                self.bookmark_actions[action] = node
        # 直接在菜单栏顶层添加
        menubar = self.menu_bar
        # 先添加所有书签和文件夹
        for child in bookmark_bar['children']:
            if child.get('id') in hidden_top_level_ids:
                continue
            if child.get('type') == 'folder':
                add_menu_items(menubar, child)
                effective_color = self._bookmark_effective_colors.get(child.get('id'))
                if effective_color and menubar.actions():
                    action_group_colors[menubar.actions()[-1]] = effective_color
            elif child.get('type') == 'url':
                # 判断是否为四个常用项目
                special_icons = ["🖥️", "🗔", "🗑️", "🚀", "⬇️"]
                name = child.get('name', '')
                is_special = any(name.startswith(icon) for icon in special_icons)
                is_group_separator = self._is_group_separator_node(child)
                if is_special:
                    action = menubar.addAction(name)
                elif is_group_separator:
                    prefix = "◨" if bool(child.get('group_collapsed', False)) else "◧"
                    member_count = int(group_member_map.get(child.get('id'), 0))
                    open_count = int(open_count_map.get(str(child.get('group_color', '')).strip().lower(), 0))
                    badge = tr("成员 {} | 开页 {}").format(member_count, open_count)
                    action = menubar.addAction(f"{prefix} {name} [{badge}]")
                else:
                    action = menubar.addAction(f"📑 {name}")
                action.triggered.connect(lambda checked, n=child: self.open_bookmark_node(n))
                # 存储action和节点的映射
                self.bookmark_actions[action] = child
                effective_color = self._bookmark_effective_colors.get(child.get('id'))
                if effective_color:
                    action_group_colors[action] = effective_color

        if hasattr(self.menu_bar, 'set_action_group_colors'):
            self.menu_bar.set_action_group_colors(action_group_colors)
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

        if event.type() == QEvent.Show and isinstance(obj, QWidget) and obj.isWindow():
            _theme.apply_window_frame(obj)

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

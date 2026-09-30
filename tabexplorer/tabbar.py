"""标签栏、标签列表与拖放标签控件。"""

import os

from PyQt5.QtCore import QEvent, QPoint, Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from PyQt5.QtWidgets import (
    QDialog, QHeaderView, QLineEdit, QMenuBar, QTabBar, QTabWidget, QToolButton, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout,
)

from .i18n import tr
from .debuglog import debug_print
from .widgets import _set_tool_icon


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
class CustomMenuBar(QMenuBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = parent
        self._action_group_colors = {}

    def set_action_group_colors(self, action_colors):
        self._action_group_colors = dict(action_colors or {})
        self.update()
    
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

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self._action_group_colors:
            return
        try:
            from PyQt5.QtGui import QPainter, QColor
            painter = QPainter(self)
            for action in self.actions():
                color_hex = self._action_group_colors.get(action)
                if not color_hex:
                    continue
                rect = self.actionGeometry(action)
                if not rect.isValid() or rect.width() <= 8:
                    continue
                marker_rect = rect.adjusted(4, rect.height() - 4, -4, -1)
                painter.fillRect(marker_rect, QColor(color_hex))
            painter.end()
        except Exception:
            pass


def _tab_display_labels(paths):
    from collections import Counter
    from pathlib import PureWindowsPath
    normalized = [str(PureWindowsPath(path)).casefold() for path in paths]
    parts = [PureWindowsPath(path).parts for path in paths]
    totals = Counter(normalized)
    seen = Counter()
    labels = []
    special = {'shell:RecycleBinFolder': tr('回收站'), 'shell:MyComputerFolder': tr('此电脑'),
               'shell:Desktop': tr('桌面'), 'shell:NetworkPlacesFolder': tr('网络')}
    for index, path in enumerate(paths):
        label = special.get(path)
        if label is None:
            label = path or tr('新标签页')
            for depth in range(1, len(parts[index]) + 1):
                suffix = parts[index][-depth:]
                label = str(PureWindowsPath(*suffix))
                if not any(normalized[other] != normalized[index]
                           and tuple(part.casefold() for part in parts[other][-depth:])
                           == tuple(part.casefold() for part in suffix)
                           for other in range(len(paths))):
                    break
        seen[normalized[index]] += 1
        if totals[normalized[index]] > 1:
            label += f' [{seen[normalized[index]]}]'
        labels.append(label)
    return labels


class TabListDialog(QDialog):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner
        self.setWindowTitle(tr('查找标签'))
        self.resize(720, 400)
        layout = QVBoxLayout(self)
        self.filter_input = QLineEdit(self)
        self.filter_input.setPlaceholderText(tr('名称或路径'))
        self.filter_input.setClearButtonEnabled(True)
        layout.addWidget(self.filter_input)
        self.table = QTreeWidget(self)
        self.table.setHeaderLabels([tr('标签'), tr('路径'), tr('位置')])
        self.table.setRootIsDecorated(False)
        self.table.setUniformRowHeights(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionMode(QTreeWidget.SingleSelection)
        self.table.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(2, 70)
        layout.addWidget(self.table)
        self.entries = []
        owner._refresh_tab_labels()
        for side, (tabs, stack) in enumerate(owner._all_groups()):
            for index in range(min(tabs.count(), stack.count())):
                pane = stack.widget(index)
                path = str(getattr(pane, 'current_path', '') or '')
                item = QTreeWidgetItem([tabs.tabText(index), path, tr('左侧') if side == 0 else tr('右侧')])
                item.setIcon(0, tabs.tabIcon(index))
                item.setToolTip(1, path)
                item.setData(0, Qt.UserRole, len(self.entries))
                self.entries.append(pane)
                self.table.addTopLevelItem(item)
        self.filter_input.textChanged.connect(self.filter_tabs)
        self.filter_input.returnPressed.connect(self.activate_selected)
        self.filter_input.installEventFilter(self)
        self.table.itemActivated.connect(self.activate_selected)
        self.filter_tabs('')
        self.filter_input.setFocus()

    def filter_tabs(self, text):
        words = text.casefold().split()
        first = None
        for index in range(self.table.topLevelItemCount()):
            item = self.table.topLevelItem(index)
            value = ' '.join(item.text(column) for column in range(3)).casefold()
            visible = all(word in value for word in words)
            item.setHidden(not visible)
            if visible and first is None:
                first = item
        self.table.setCurrentItem(first)

    def eventFilter(self, watched, event):
        if watched is self.filter_input and event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Down, Qt.Key_Up):
            visible = [self.table.topLevelItem(index) for index in range(self.table.topLevelItemCount())
                       if not self.table.topLevelItem(index).isHidden()]
            if visible:
                selected = self.table.currentItem()
                index = visible.index(selected) if selected in visible else 0
                index = max(0, min(len(visible) - 1, index + (1 if event.key() == Qt.Key_Down else -1)))
                self.table.setCurrentItem(visible[index])
                self.table.scrollToItem(visible[index])
            return True
        return super().eventFilter(watched, event)

    def activate_selected(self, *args):
        item = self.table.currentItem()
        if item is None or item.isHidden():
            return
        target = self.entries[item.data(0, Qt.UserRole)]
        for tabs, stack in self.owner._all_groups():
            for index in range(stack.count()):
                if stack.widget(index) is target:
                    tabs.setCurrentIndex(index)
                    self.owner.set_active_pane_to_group(tabs)
                    self.accept()
                    return
        item.setHidden(True)
        self.table.setCurrentItem(None)


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
        self._drag_button = None
        self._press_group_block = None
        self._dragging = False
        self._drag_preview = None
        self._suppress_context_menu_once = False
        self._drop_indicator_active = False
        self._drop_indicator_index = -1
        from PyQt5.QtCore import QTimer
        self._drag_end_timer = QTimer(self)
        self._drag_end_timer.setSingleShot(True)
        self._drag_end_timer.setInterval(150)
        self._drag_end_timer.timeout.connect(self._on_drag_end_timeout)

    def paintEvent(self, event):
        super().paintEvent(event)
        try:
            mw = getattr(self, 'main_window', None)
            if mw is None:
                return
            show_markers = bool(getattr(mw, 'config', {}).get('show_tab_group_markers', True))
            if not show_markers:
                return
            cs = mw._content_stack_for(self._owner_tw()) if hasattr(mw, '_content_stack_for') else None
            if cs is None:
                return

            from PyQt5.QtCore import QPoint
            from PyQt5.QtGui import QPainter, QColor, QPainterPath, QPen
            painter = QPainter(self)
            painter.setRenderHint(QPainter.Antialiasing, True)
            count = min(self.count(), cs.count())
            dpi_scale = float(getattr(mw, 'dpi_scale', 1.0) or 1.0)
            corner_radius = max(4.0, float(int(6 * dpi_scale)) - 0.5)
            for i in range(count):
                tab = cs.widget(i)
                if tab is None:
                    continue
                color_hex = str(getattr(tab, 'bookmark_group_color', '') or '').strip()
                if not color_hex:
                    continue
                rect = self.tabRect(i)
                if not rect.isValid() or rect.width() <= 4:
                    continue

                # 分组使用圆角背景高亮，并内缩到标签内部，避免出现“比标签更大/尖角”的视觉问题。
                base = QColor(color_hex)
                fill = QColor(base)
                fill.setAlpha(72 if i == self.currentIndex() else 52)
                inner = rect.adjusted(2, 2, -2, -1)
                if inner.width() <= 4 or inner.height() <= 4:
                    continue
                path = QPainterPath()
                path.addRoundedRect(float(inner.x()), float(inner.y()), float(inner.width()), float(inner.height()),
                                    corner_radius, corner_radius)
                painter.fillPath(path, fill)

            # 拖拽时明确显示“将插入到这里”的位置指示线。
            if self._drop_indicator_active:
                insert_idx = int(self._drop_indicator_index)
                insert_idx = max(0, min(insert_idx, self.count()))
                x = 6
                if self.count() > 0:
                    if insert_idx >= self.count():
                        last_rect = self.tabRect(self.count() - 1)
                        if last_rect.isValid():
                            x = int(last_rect.right()) + 1
                    else:
                        target_rect = self.tabRect(insert_idx)
                        if target_rect.isValid():
                            x = int(target_rect.left())
                line_pen = QPen(QColor("#D32F2F"))
                line_pen.setWidth(4)
                painter.setPen(line_pen)
                y1 = 8
                y2 = max(y1 + 10, self.height() - 3)
                painter.drawLine(x, y1, x, y2)

                dot_color = QColor("#D32F2F")
                painter.setPen(Qt.NoPen)
                painter.setBrush(dot_color)
                # 倒三角箭头，明确指示插入点
                painter.drawPolygon(
                    QPoint(x - 8, 1),
                    QPoint(x + 8, 1),
                    QPoint(x, y1)
                )
            painter.end()
        except Exception:
            pass

    def set_drop_insert_indicator(self, insert_index):
        try:
            idx = int(insert_index)
        except Exception:
            idx = -1
        idx = max(0, min(idx, self.count()))
        if self._drop_indicator_active and self._drop_indicator_index == idx:
            return
        self._drop_indicator_active = True
        self._drop_indicator_index = idx
        self.update()

    def clear_drop_insert_indicator(self):
        if not self._drop_indicator_active and self._drop_indicator_index < 0:
            return
        self._drop_indicator_active = False
        self._drop_indicator_index = -1
        self.update()

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
        self._drag_button = None
        self._press_group_block = None
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
                        self._drag_button = Qt.LeftButton
            elif event.button() == Qt.RightButton and self.main_window is not None:
                if hasattr(self.main_window, 'set_active_pane_to_group'):
                    try:
                        self.main_window.set_active_pane_to_group(self._owner_tw())
                    except Exception:
                        pass
                idx = self.tabAt(event.pos())
                if idx >= 0:
                    cs = self.main_window._content_stack_for(self._owner_tw())
                    if cs is not None and idx < cs.count():
                        pressed_tab = cs.widget(idx)
                        if pressed_tab is not None and not bool(getattr(pressed_tab, 'is_pinned', False)):
                            block = None
                            if hasattr(self.main_window, '_get_tab_group_block_for_drag'):
                                block = self.main_window._get_tab_group_block_for_drag(self._owner_tw(), idx)
                            if block is not None:
                                self._press_group_block = block
                                self._press_content = pressed_tab
                                self._press_title = tr("整组移动")
                                self._press_index = idx
                                self._press_pos = event.pos()
                                self._drag_button = Qt.RightButton
        except Exception:
            self._press_content = None
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        # 悬停追踪：更新鼠标下的标签索引（供关闭按钮显示等）
        self.hovered_tab = self.tabAt(event.pos())
        # 拖拽判定：左键按住并移动超过阈值 → 进入拖拽，显示跟随光标的浮动预览
        expected_btn = None
        if self._drag_button == Qt.LeftButton:
            expected_btn = Qt.LeftButton
        elif self._drag_button == Qt.RightButton:
            expected_btn = Qt.RightButton
        if (self._press_index >= 0 and expected_btn is not None and (event.buttons() & expected_btn)
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
                dest_tw, dest_index = None, -1
                try:
                    if self.main_window is not None and hasattr(self.main_window, '_pane_group_hit_test'):
                        dest_tw, dest_index = self.main_window._pane_group_hit_test(event.globalPos())
                except Exception:
                    dest_tw, dest_index = None, -1
                if self.main_window is not None and hasattr(self.main_window, '_update_drag_insert_indicators'):
                    self.main_window._update_drag_insert_indicators(dest_tw, dest_index)
                self._update_drag_preview(event.globalPos(), dest_tw)
        super().mouseMoveEvent(event)

    def _update_drag_preview(self, gpos, dest_tw=None):
        """拖拽中显示跟随光标的浮动预览（标签标题气泡），并按目标组区分提示样式。"""
        if dest_tw is None:
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
        drag_button = self._drag_button
        press_content = self._press_content
        press_index = self._press_index
        press_group_block = self._press_group_block
        self._press_content = None
        self._press_index = -1
        self._press_pos = None
        self._drag_button = None
        self._press_group_block = None
        self._dragging = False
        self._hide_drag_preview()
        if self.main_window is not None and hasattr(self.main_window, '_clear_drag_insert_indicators'):
            self.main_window._clear_drag_insert_indicators()
        super().mouseReleaseEvent(event)

        if not was_dragging or press_content is None:
            # 普通点击（未触发拖拽）：清理拖拽标志即可
            if self._is_dragging_tab:
                self._is_dragging_tab = False
                if self.main_window is not None:
                    self.main_window._tab_drag_in_progress = False
            return

        owner = self._owner_tw()
        moved = False
        dest_tw, dest_index = None, -1
        try:
            if self.main_window is not None and hasattr(self.main_window, '_pane_group_hit_test'):
                dest_tw, dest_index = self.main_window._pane_group_hit_test(event.globalPos())
        except Exception:
            dest_tw, dest_index = None, -1

        if drag_button == Qt.LeftButton:
            if event.button() != Qt.LeftButton:
                if self._is_dragging_tab:
                    self._is_dragging_tab = False
                    if self.main_window is not None:
                        self.main_window._tab_drag_in_progress = False
                return

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
                    insert_slot = -1
                    if self.main_window is not None and hasattr(self.main_window, '_tabbar_insert_index_from_global'):
                        insert_slot = self.main_window._tabbar_insert_index_from_global(owner, event.globalPos())
                    if insert_slot < 0:
                        insert_slot = self.count()
                    if 0 <= press_index < self.count():
                        target = insert_slot if insert_slot <= press_index else (insert_slot - 1)
                        target = max(0, min(target, self.count() - 1))
                    else:
                        target = -1
                    if 0 <= target < self.count() and target != press_index:
                        # 单拖边界标签时，落位后应视为普通成员，不保留边界定义。
                        if bool(getattr(press_content, 'tab_group_separator_after', False)):
                            press_content.tab_group_separator_after = False
                            press_content.tab_group_separator_color = ""
                            press_content.tab_group_separator_name = ""
                        self.moveTab(press_index, target)
                        if hasattr(self.main_window, '_apply_right_neighbor_grouping_for_moved_tabs'):
                            self.main_window._apply_right_neighbor_grouping_for_moved_tabs(owner, [press_content])
                        moved = True
                except Exception as _e:
                    debug_print(f"[TabReorder] failed: {_e}")
        elif drag_button == Qt.RightButton:
            if event.button() != Qt.RightButton:
                if self._is_dragging_tab:
                    self._is_dragging_tab = False
                    if self.main_window is not None:
                        self.main_window._tab_drag_in_progress = False
                return

            if press_group_block is None:
                moved = False
            elif dest_tw is not None and dest_tw is not owner:
                try:
                    moved = self.main_window.move_tab_group_across_groups(
                        owner, press_group_block, press_content, dest_tw, dest_index)
                except Exception as _e:
                    debug_print(f"[GroupDrag] cross-group transfer failed: {_e}")
            elif dest_tw is owner:
                try:
                    moved = self.main_window.reorder_tab_group_within_group(
                        owner, press_group_block, dest_index, press_content)
                except Exception as _e:
                    debug_print(f"[GroupDrag] reorder failed: {_e}")
            # 右键拖动发生移动时，抑制这次右键菜单弹出
            if moved:
                self._suppress_context_menu_once = True

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

    def consume_context_menu_suppression(self):
        if self._suppress_context_menu_once:
            self._suppress_context_menu_once = False
            return True
        return False
    
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
        from PyQt5.QtWidgets import QStyle
        close_btn = QToolButton(self)
        close_btn.setToolTip(tr('关闭标签'))
        scale = float(getattr(getattr(self, 'main_window', None), 'dpi_scale', 1.0))
        tab_close_btn_size = max(16, int(16 * scale))
        _set_tool_icon(close_btn, 'window-close', QStyle.SP_TitleBarCloseButton, max(12, int(12 * scale)))
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
        self._schedule_labels()

    def tabRemoved(self, index):
        super().tabRemoved(index)
        self._schedule_labels()

    def _schedule_labels(self):
        owner = getattr(self, 'main_window', None)
        if owner is not None and hasattr(owner, '_schedule_tab_labels'):
            owner._schedule_tab_labels()

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
        if getattr(self, '_sorting_pinned_tabs', False):
            return
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
        try:
            self.main_window._apply_tab_grouping_for_pane(tw)
        except Exception:
            pass
        # 移动后自动检测鼠标下的tab并显示关闭按钮
        self.show_close_button_under_cursor()
        self._schedule_labels()
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

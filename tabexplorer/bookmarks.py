"""书签数据与书签管理窗口。"""

import json
import os
import time

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QHBoxLayout, QInputDialog, QLabel, QPushButton, QTreeWidget, QTreeWidgetItem, QVBoxLayout,
)

from . import theme as _theme
from .paths import get_app_data_path, translate_common_path
from .i18n import tr
from .debuglog import debug_print
from .widgets import show_toast


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


def _preserve_unreadable_file(path):
    """把无法解析的数据文件改名保留，避免随后写入默认内容时覆盖用户数据；返回备份路径。"""
    backup = f"{path}.broken-{time.strftime('%Y%m%d-%H%M%S')}"
    try:
        os.replace(path, backup)
        return backup
    except OSError as e:
        debug_print(f"[Data] 无法保留损坏文件 {path}: {e}")
        return ''


def _bookmark_ids(nodes):
    ids, stack = set(), list(nodes)
    while stack:
        node = stack.pop()
        if isinstance(node, dict):
            if node.get('id'):
                ids.add(str(node['id']))
            stack.extend(node.get('children') or [])
    return ids


def _assign_new_bookmark_ids(nodes, used_ids):
    """导入的节点重新编号：重复 ID 会让按 ID 查找的删除/分组操作命中错误节点。"""
    next_id = int(time.time() * 1000000)
    stack = list(nodes)
    while stack:
        node = stack.pop()
        if not isinstance(node, dict):
            continue
        while str(next_id) in used_ids:
            next_id += 1
        node['id'] = str(next_id)
        used_ids.add(node['id'])
        stack.extend(node.get('children') or [])


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
        self.recovered_backup = ''
        if not os.path.exists(self.config_file):
            debug_print("No bookmark file found, starting with empty bookmarks")
            return {}
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict) and 'roots' in data:
                data = data['roots']
            if not isinstance(data, dict):
                raise ValueError("bookmark root is not an object")
            return data
        except ValueError as e:
            debug_print(f"Failed to load bookmarks: {e}")
            self.recovered_backup = _preserve_unreadable_file(self.config_file)
            return {}
        except OSError as e:
            debug_print(f"Failed to load bookmarks: {e}")
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


class _GroupAwareBookmarkTree(QTreeWidget):
    """书签树拖拽增强：顶层分组块在拖拽阶段保持整体移动，不允许拆组。"""

    def __init__(self, owner_dialog, parent=None):
        super().__init__(parent)
        self._owner_dialog = owner_dialog
        self._drag_block = None
        self._drag_anchor_id = None
        self._preview_block = None
        self._preview_insert_index = -1
        self._preview_drop_valid = True
        self._auto_scroll_edge_px = 24
        self._auto_scroll_step_px = 22

    def _auto_scroll_on_drag(self, pos):
        """拖拽到视口上下边缘时自动滚动，便于长列表跨屏移动。"""
        try:
            bar = self.verticalScrollBar()
            if bar is None:
                return
            y = int(pos.y())
            vh = int(self.viewport().height())
            edge = int(self._auto_scroll_edge_px)
            step = int(self._auto_scroll_step_px)
            if y < edge:
                bar.setValue(max(bar.minimum(), bar.value() - step))
            elif y > max(0, vh - edge):
                bar.setValue(min(bar.maximum(), bar.value() + step))
        except Exception:
            pass

    def startDrag(self, supportedActions):
        self._drag_block = None
        self._drag_anchor_id = None
        self._preview_block = None
        self._preview_insert_index = -1
        self._preview_drop_valid = True
        item = self.currentItem()
        if item is not None and self._owner_dialog is not None:
            self._drag_anchor_id = item.data(0, 1)
            self._drag_block = self._owner_dialog._get_drag_group_block(item)
            self._preview_block = self._drag_block
        self.viewport().update()
        try:
            super().startDrag(supportedActions)
        finally:
            self._drag_block = None
            self._drag_anchor_id = None
            self._preview_block = None
            self._preview_insert_index = -1
            self._preview_drop_valid = True
            self.viewport().update()

    def dragMoveEvent(self, event):
        if self._drag_block is not None:
            self._auto_scroll_on_drag(event.pos())
            self._preview_insert_index = self._owner_dialog._drop_target_top_level_index(event.pos())
            start, end = self._drag_block
            self._preview_drop_valid = not (start <= self._preview_insert_index <= (end + 1))
            self.viewport().update()
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dragLeaveEvent(self, event):
        self._preview_insert_index = -1
        self._preview_drop_valid = True
        self.viewport().update()
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        if self._drag_block is not None and self._owner_dialog is not None:
            moved = self._owner_dialog._move_group_block_in_tree(self._drag_block, event.pos(), self._drag_anchor_id)
            self._drag_block = None
            self._drag_anchor_id = None
            self._preview_block = None
            self._preview_insert_index = -1
            self._preview_drop_valid = True
            self.viewport().update()
            if moved:
                event.acceptProposedAction()
                return
            event.ignore()
            return
        super().dropEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._preview_block is None:
            return
        start, end = self._preview_block
        if start < 0 or end < start or end >= self.topLevelItemCount():
            return
        try:
            from PyQt5.QtGui import QPainter, QColor, QPen
            first_item = self.topLevelItem(start)
            last_item = self.topLevelItem(end)
            if first_item is None or last_item is None:
                return
            first_rect = self.visualItemRect(first_item)
            last_rect = self.visualItemRect(last_item)
            if not first_rect.isValid() or not last_rect.isValid():
                return

            painter = QPainter(self.viewport())

            # 被拖拽分组块高亮
            block_rect = first_rect.united(last_rect).adjusted(1, 0, -1, 0)
            painter.fillRect(block_rect, QColor(100, 181, 246, 48))
            pen = QPen(QColor("#42A5F5"))
            pen.setWidth(2)
            painter.setPen(pen)
            painter.drawRect(block_rect)

            # 分组移动提示徽标
            badge_text = tr("整组移动")
            badge_h = 18
            badge_w = max(70, self.fontMetrics().width(badge_text) + 16)
            badge_rect = block_rect.adjusted(6, 4, -6, -(block_rect.height() - badge_h - 4))
            badge_rect.setWidth(min(badge_w, max(60, block_rect.width() - 12)))
            painter.fillRect(badge_rect, QColor(30, 136, 229, 220))
            painter.setPen(QColor("#FFFFFF"))
            painter.drawText(badge_rect, Qt.AlignCenter, badge_text)

            # 插入位置指示线
            ins = int(self._preview_insert_index)
            if 0 <= ins <= self.topLevelItemCount():
                y = None
                if ins == self.topLevelItemCount():
                    tail_item = self.topLevelItem(self.topLevelItemCount() - 1)
                    if tail_item is not None:
                        tail_rect = self.visualItemRect(tail_item)
                        if tail_rect.isValid():
                            y = tail_rect.bottom() + 1
                else:
                    tgt_item = self.topLevelItem(ins)
                    if tgt_item is not None:
                        tgt_rect = self.visualItemRect(tgt_item)
                        if tgt_rect.isValid():
                            y = tgt_rect.top()
                if y is not None:
                    color_hex = "#1E88E5" if self._preview_drop_valid else "#9E9E9E"
                    insert_pen = QPen(QColor(color_hex))
                    insert_pen.setWidth(3)
                    if not self._preview_drop_valid:
                        insert_pen.setStyle(Qt.DashLine)
                    painter.setPen(insert_pen)
                    painter.drawLine(4, y, max(4, self.viewport().width() - 4), y)
                    if not self._preview_drop_valid:
                        tip_rect = badge_rect.adjusted(0, badge_rect.height() + 4, 48, badge_rect.height() + 20)
                        painter.fillRect(tip_rect, QColor(120, 120, 120, 200))
                        painter.setPen(QColor("#FFFFFF"))
                        painter.drawText(tip_rect, Qt.AlignCenter, tr("不可放置"))

            painter.end()
        except Exception:
            pass


class BookmarkManagerDialog(QDialog):
    def _is_group_separator_node(self, node):
        return bool(isinstance(node, dict) and node.get('type') == 'url' and node.get('is_group_separator'))

    def _find_node_by_id(self, node_id):
        if not node_id:
            return None

        def _walk(node):
            if isinstance(node, dict):
                if node.get('id') == node_id:
                    return node
                for child in node.get('children', []) or []:
                    found = _walk(child)
                    if found is not None:
                        return found
            elif isinstance(node, list):
                for child in node:
                    found = _walk(child)
                    if found is not None:
                        return found
            return None

        tree = self.bookmark_manager.get_tree()
        return _walk(tree.get('bookmark_bar')) if isinstance(tree, dict) else None

    def _is_group_separator_id(self, node_id):
        return self._is_group_separator_node(self._find_node_by_id(node_id))

    def _top_level_block_ranges(self):
        """按当前树的顶层顺序生成分组块区间（start, end，含分隔节点）。"""
        ranges = []
        pending_start = 0
        top_count = self.tree.topLevelItemCount()
        for i in range(top_count):
            item = self.tree.topLevelItem(i)
            node_id = item.data(0, 1)
            if self._is_group_separator_id(node_id):
                ranges.append((pending_start, i))
                pending_start = i + 1
        if pending_start < top_count:
            for j in range(pending_start, top_count):
                item = self.tree.topLevelItem(j)
                if self._is_group_separator_id(item.data(0, 1)):
                    break
            else:
                # 未遇到分隔符：尾部每个节点视为独立块，避免被错误合并
                for k in range(pending_start, top_count):
                    ranges.append((k, k))
                return ranges
            for k in range(pending_start, top_count):
                ranges.append((k, k))
        return ranges

    def _get_drag_group_block(self, item):
        """若拖拽的是顶层分组成员，则返回其分组块区间；否则返回 None。"""
        if item is None or item.parent() is not None:
            return None
        idx = self.tree.indexOfTopLevelItem(item)
        if idx < 0:
            return None
        for start, end in self._top_level_block_ranges():
            if start <= idx <= end:
                if end > start:
                    return (start, end)
                return None
        return None

    def _drop_target_top_level_index(self, pos):
        target = self.tree.itemAt(pos)
        if target is None:
            return self.tree.topLevelItemCount()
        while target.parent() is not None:
            target = target.parent()
        idx = self.tree.indexOfTopLevelItem(target)
        if idx < 0:
            return self.tree.topLevelItemCount()
        rect = self.tree.visualItemRect(target)
        return idx if pos.y() < rect.center().y() else idx + 1

    def _move_group_block_in_tree(self, block, pos, anchor_id=None):
        """在拖拽落点处整体移动顶层分组块（UI层），不拆分成员。"""
        if not block or len(block) != 2:
            return False
        start, end = int(block[0]), int(block[1])
        top_count = self.tree.topLevelItemCount()
        if start < 0 or end >= top_count or start > end:
            return False

        target_index = self._drop_target_top_level_index(pos)
        if start <= target_index <= end + 1:
            return False

        items = []
        while self.tree.topLevelItemCount() > 0:
            items.append(self.tree.takeTopLevelItem(0))
        moving = items[start:end + 1]
        remain = items[:start] + items[end + 1:]

        if target_index > end:
            target_index -= len(moving)
        target_index = max(0, min(target_index, len(remain)))
        reordered = remain[:target_index] + moving + remain[target_index:]

        for it in reordered:
            self.tree.addTopLevelItem(it)

        if anchor_id:
            self.reselect_item_by_id(anchor_id)
        # 拖拽完成后立即把 UI 顺序同步回数据，避免“看起来动了但未保存”。
        self.on_items_moved()
        return True

    def _rebuild_group_block_protected_order(self, original_children, new_children):
        """按原始分组块重建顶层顺序，避免拖拽把“分组+成员”拆散。"""
        if not isinstance(original_children, list) or not isinstance(new_children, list):
            return new_children

        id_to_pos = {}
        for idx, node in enumerate(new_children):
            if isinstance(node, dict) and node.get('id'):
                id_to_pos[node.get('id')] = idx

        blocks = []
        pending = []
        for node in original_children:
            if not isinstance(node, dict):
                continue
            pending.append(node)
            if self._is_group_separator_node(node):
                blocks.append(list(pending))
                pending = []
        for node in pending:
            blocks.append([node])

        scored_blocks = []
        fallback = len(id_to_pos) + 1000
        for i, block in enumerate(blocks):
            positions = []
            for n in block:
                nid = n.get('id') if isinstance(n, dict) else None
                if nid in id_to_pos:
                    positions.append(id_to_pos[nid])
            score = (min(positions) if positions else (fallback + i), i)
            scored_blocks.append((score, block))

        scored_blocks.sort(key=lambda x: x[0])

        result = []
        emitted = set()
        for _score, block in scored_blocks:
            for n in block:
                if not isinstance(n, dict):
                    continue
                nid = n.get('id')
                if not nid or nid in emitted:
                    continue
                emitted.add(nid)
                result.append(n)

        for n in new_children:
            nid = n.get('id') if isinstance(n, dict) else None
            if nid and nid not in emitted:
                emitted.add(nid)
                result.append(n)
        return result

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

            tree = self.bookmark_manager.get_tree()
            original_children = []
            if isinstance(tree, dict) and isinstance(tree.get('bookmark_bar'), dict):
                original_children = list(tree.get('bookmark_bar', {}).get('children', []) or [])
            new_structure = self._rebuild_group_block_protected_order(original_children, new_structure)
            
            debug_print(f"[BookmarkDrag] Rebuilt {len(new_structure)} top-level items")
            
            # 更新书签管理器
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
                node = {
                    'id': node_id,
                    'name': name,
                    'type': 'url',
                    'url': url,
                    'date_added': original.get('date_added', node_id)
                }
                if original.get('is_group_separator'):
                    node['is_group_separator'] = True
                    if 'group_color' in original:
                        node['group_color'] = original.get('group_color')
                    if 'group_collapsed' in original:
                        node['group_collapsed'] = bool(original.get('group_collapsed', False))
                return node
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
        self.tree = _GroupAwareBookmarkTree(self)
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
        _theme.bind_style(drag_hint, "QLabel { color: #666; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }")
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
                # 直接导出内存中的书签：复制磁盘文件会漏掉尚未落盘的延迟保存
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump({"roots": self.bookmark_manager.get_tree()}, f, ensure_ascii=False, indent=2)
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
                current_bar = current_tree.setdefault('bookmark_bar', {
                    "id": str(int(time.time() * 1000000)), "name": tr("书签栏"), "type": "folder", "children": []})
                if 'children' not in current_bar:
                    current_bar['children'] = []
                _assign_new_bookmark_ids(imported_children, _bookmark_ids(current_tree.values()))
                
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

from PyQt5.QtWidgets import QDialog, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QPushButton
# 多层结构书签弹窗
class BookmarkDialog(QDialog):
    def __init__(self, bookmark_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("书签")
        self.resize(500, 600)
        self.bookmark_manager = bookmark_manager
        layout = QVBoxLayout(self)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["名称", "路径"])
        layout.addWidget(self.tree)
        self.populate_tree()
        self.tree.itemDoubleClicked.connect(self.on_item_double_clicked)
        close_btn = QPushButton("关闭")
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
            if os.path.exists(local_path):
                self.accept()
                # 通知主窗口打开新标签页
                if self.parent() and hasattr(self.parent(), 'add_new_tab'):
                    self.parent().add_new_tab(local_path)
            else:
                QMessageBox.warning(self, "路径错误", f"路径不存在: {local_path}")
import sys
# import ctypes  # unused
# import win32con  # unused
# import win32gui  # unused
# import win32api  # unused
# import win32com.shell.shell as shell  # unused
# import win32com.shell.shellcon as shellcon  # unused
# import comtypes
# import comtypes.client
import os
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QListWidget, QLabel, QToolBar, QAction, QMenu, QMessageBox, QInputDialog)  # QDockWidget removed (unused)
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import Qt, QDir, QUrl  # QModelIndex removed (unused)
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
# from PyQt5.QtGui import QIcon  # unused

# Optional native hit-test support (Windows)
try:
    import ctypes
    import win32gui
    import win32con
    HAS_PYWIN = True
except Exception:
    HAS_PYWIN = False

class BookmarkManager:
    def __init__(self, config_file="bookmarks.json"):
        self.config_file = config_file
        self.bookmark_tree = self.load_bookmarks()

    def load_bookmarks(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # 兼容Chrome风格的多层结构
                if 'roots' in data:
                    return data['roots']
                return data
            except Exception as e:
                print(f"Failed to load bookmarks: {e}")
                return {}
        return {}

    def save_bookmarks(self):
        # 只保存根结构
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump({"roots": self.bookmark_tree}, f, ensure_ascii=False, indent=2)

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


class FileExplorerTab(QWidget):
    def update_tab_title(self):
        if hasattr(self, 'current_path'):
            # shell: 路径中文映射
            shell_map = {
                'shell:RecycleBinFolder': '回收站',
                'shell:MyComputerFolder': '此电脑',
                'shell:Desktop': '桌面',
                'shell:NetworkPlacesFolder': '网络',
            }
            path = self.current_path
            display = shell_map.get(path, None)
            if not display and path.startswith('shell:'):
                display = path  # 兜底显示原始shell:路径
            if not display:
                display = path[-8:] if len(path) > 8 else path
            title = ("📌" if getattr(self, 'is_pinned', False) else "") + display
            if self.main_window and hasattr(self.main_window, 'tab_widget'):
                idx = self.main_window.tab_widget.indexOf(self)
                if idx != -1:
                    self.main_window.tab_widget.setTabText(idx, title)

    def start_path_sync_timer(self):
        from PyQt5.QtCore import QTimer
        self._path_sync_timer = QTimer(self)
        self._path_sync_timer.timeout.connect(self.sync_path_bar_with_explorer)
        self._path_sync_timer.start(500)

    def sync_path_bar_with_explorer(self):
        # 通过QAxWidget的LocationURL属性获取当前路径
        try:
            url = self.explorer.property('LocationURL')
            if url and str(url).startswith('file:///'):
                from urllib.parse import unquote
                local_path = unquote(str(url)[8:])
                if os.name == 'nt' and local_path.startswith('/'):
                    local_path = local_path[1:]
                if local_path != self.current_path:
                    self.current_path = local_path
                    if hasattr(self, 'path_bar'):
                        self.path_bar.setText(local_path)
                    self.update_tab_title()
                    # 同步左侧目录树
                    if self.main_window and hasattr(self.main_window, 'expand_dir_tree_to_path'):
                        self.main_window.expand_dir_tree_to_path(local_path)
        except Exception:
            pass

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 路径栏
        self.path_bar = QLineEdit(self)
        self.path_bar.setReadOnly(False)
        self.path_bar.returnPressed.connect(self.on_path_enter)
        # 双击路径栏等效上一级按钮
        self.path_bar.mouseDoubleClickEvent = lambda e: self.go_up(force=True)
        layout.addWidget(self.path_bar)

        # 嵌入Explorer控件
        self.explorer = QAxWidget(self)
        self.explorer.setControl("Shell.Explorer")
        layout.addWidget(self.explorer)
        # 绑定导航完成信号，自动更新路径栏
        self.explorer.dynamicCall('NavigateComplete2(QVariant,QVariant)', None, None)  # 预绑定，防止信号未注册
        self.explorer.dynamicCall('Navigate(const QString&)', QDir.toNativeSeparators(self.current_path))
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
        self.explorer.dynamicCall('NavigateComplete2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('NavigateError(QVariant,QVariant,QVariant,QVariant)', None, None, None, None)
        self.explorer.dynamicCall('DocumentComplete(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('BeforeNavigate2(QVariant,QVariant,QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None, None, None)
        self.explorer.dynamicCall('NewWindow2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('NewWindow3(QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None)
        self.explorer.dynamicCall('OnQuit()', )
        self.explorer.dynamicCall('OnVisible()', )
        self.explorer.dynamicCall('OnToolBar()', )
        self.explorer.dynamicCall('OnMenuBar()', )
        self.explorer.dynamicCall('OnStatusBar()', )
        self.explorer.dynamicCall('OnFullScreen()', )
        self.explorer.dynamicCall('OnTheaterMode()', )
        self.explorer.dynamicCall('OnAddressBar()', )
        self.explorer.dynamicCall('OnResizable()', )
        self.explorer.dynamicCall('OnOffline()', )
        self.explorer.dynamicCall('OnSilent()', )
        self.explorer.dynamicCall('OnRegisterAsBrowser()', )
        self.explorer.dynamicCall('OnRegisterAsDropTarget()', )
        self.explorer.dynamicCall('OnNavigateComplete2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnNavigateError(QVariant,QVariant,QVariant,QVariant)', None, None, None, None)
        self.explorer.dynamicCall('OnDocumentComplete(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnBeforeNavigate2(QVariant,QVariant,QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None, None, None)
        self.explorer.dynamicCall('OnNewWindow2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnNewWindow3(QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None)
        self.explorer.dynamicCall('OnQuit()', )
        # 连接信号
        self.explorer.dynamicCall('NavigateComplete2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('DocumentComplete(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('BeforeNavigate2(QVariant,QVariant,QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None, None, None)
        self.explorer.dynamicCall('NewWindow2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('NewWindow3(QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None)
        self.explorer.dynamicCall('OnNavigateComplete2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnDocumentComplete(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnBeforeNavigate2(QVariant,QVariant,QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None, None, None)
        self.explorer.dynamicCall('OnNewWindow2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnNewWindow3(QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None)


        # 在Explorer下方添加盘符列表（初始隐藏）
        from PyQt5.QtWidgets import QListWidget, QLabel
        self.drive_list = QListWidget()
        self.drive_list.setVisible(False)
        self.drive_list.itemDoubleClicked.connect(self.on_drive_double_clicked)
        layout.addWidget(self.drive_list)
        # 兼容原有空白双击
        self.blank = QLabel()
        # 保持空白区域为固定高度，避免其扩展占满右侧空间
        self.blank.setFixedHeight(10)
        self.blank.setStyleSheet("background: transparent;")
        self.blank.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.blank.mouseDoubleClickEvent = self.blank_double_click
        layout.addWidget(self.blank)

        # 安装事件过滤器以捕获 Explorer 的鼠标按下与双击事件
        try:
            self.explorer.installEventFilter(self)
        except Exception:
            pass

    def event(self, e):
        # 捕获QAxWidget的NavigateComplete2事件
        if e.type() == 50:  # QEvent.MetaCall, QAxWidget信号事件
            if hasattr(e, 'arguments') and hasattr(e, 'signal'):
                if e.signal == 'NavigateComplete2(IDispatch*, QVariant&)':
                    url = str(e.arguments[1])
                    if url.startswith('file:///'):
                        from urllib.parse import unquote
                        local_path = unquote(url[8:])
                        if os.name == 'nt' and local_path.startswith('/'):
                            local_path = local_path[1:]
                        self.current_path = local_path
                        if hasattr(self, 'path_bar'):
                            self.path_bar.setText(local_path)
        return super().event(e)

    def explorer_double_click(self, event):
        # 使用短延迟再检查选中项数量以避免竞态：
        # 如果在检查时没有选中项，则视为点击空白区并返回上一级。
        from PyQt5.QtCore import QTimer
        def _check_and_go_up():
            # 如果按下时已有选中项，认为双击是针对该项，跳过 go_up
            before = getattr(self, '_selected_before_click', None)
            if before is not None:
                try:
                    if int(before) > 0:
                        self._selected_before_click = None
                        return
                except Exception:
                    pass
                cnt = self.explorer.dynamicCall('SelectedItems().Count')
            try:
                cnt = self.explorer.dynamicCall('SelectedItems().Count')
            except Exception:
                try:
                    sel = self.explorer.dynamicCall('SelectedItems()')
                    if sel is not None:
                        cnt = sel.property('Count') if hasattr(sel, 'property') else None
                except Exception:
                    cnt = None
            if cnt is None:
                # 无法确定，按原先约定触发 go_up()
                try:
                    self.go_up(force=True)
                except Exception:
                    pass
                return
            try:
                if int(cnt) == 0:
                        self.go_up(force=True)
            except Exception:
                pass

        QTimer.singleShot(50, _check_and_go_up)

    def on_path_enter(self):
        path = self.path_bar.text().strip()
        # 支持中文特殊路径
        special_map = {
            '回收站': 'shell:RecycleBinFolder',
            '此电脑': 'shell:MyComputerFolder',
            '我的电脑': 'shell:MyComputerFolder',
            '桌面': 'shell:Desktop',
            '网络': 'shell:NetworkPlacesFolder',
            '启动项': 'shell:Startup',
            '开机启动项': 'shell:Startup',
            '启动文件夹': 'shell:Startup',
            'Startup': 'shell:Startup',
        }
        if path in special_map:
            self.navigate_to(special_map[path], is_shell=True)
        elif os.path.exists(path):
            self.navigate_to(path)
        else:
            QMessageBox.warning(self, "路径错误", f"路径不存在: {path}")

    def explorer_mouse_press(self, event):
        # 在鼠标按下时记录当时的选中项数量，用于后续双击判断
        try:
            cnt = None
            try:
                cnt = self.explorer.dynamicCall('SelectedItems().Count')
            except Exception:
                try:
                    sel = self.explorer.dynamicCall('SelectedItems()')
                    if sel is not None:
                        cnt = sel.property('Count') if hasattr(sel, 'property') else None
                except Exception:
                    cnt = None
            self._selected_before_click = int(cnt) if cnt is not None else None
        except Exception:
            self._selected_before_click = None
        # 继续默认处理（不阻止控件行为）
        # 直接返回 None — 不尝试调用 ActiveX 的原始处理（事件仍会被控件处理）
        return None

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
        from PyQt5.QtCore import QEvent, QTimer
        try:
            if obj is self.explorer:
                if event.type() == QEvent.MouseButtonPress:
                    # 记录按下时的选中项数
                    try:
                        cnt = None
                        try:
                            cnt = self.explorer.dynamicCall('SelectedItems().Count')
                        except Exception:
                            try:
                                sel = self.explorer.dynamicCall('SelectedItems()')
                                if sel is not None:
                                    cnt = sel.property('Count') if hasattr(sel, 'property') else None
                            except Exception:
                                cnt = None
                        self._selected_before_click = int(cnt) if cnt is not None else None
                    except Exception:
                        self._selected_before_click = None
                elif event.type() == QEvent.MouseButtonDblClick:
                    # 延迟检查，避免与控件自身处理产生竞态
                    # 记录双击发生前的路径，以便判断双击是否触发了导航
                    try:
                        self._path_before_double = getattr(self, 'current_path', None)
                    except Exception:
                        self._path_before_double = None

                    def _check_and_go_up():
                        # 如果双击期间发生了导航（进入文件夹或打开文件），则跳过 go_up
                        try:
                            before_path = getattr(self, '_path_before_double', None)
                            cur_path = getattr(self, 'current_path', None)
                            if before_path is not None and cur_path is not None and cur_path != before_path:
                                # 导航已发生，跳过上一级
                                self._path_before_double = None
                                return
                        except Exception:
                            pass
                        # 继续原有判断
                        # 如果按下时已有选中项，则认为是对项的双击
                        before = getattr(self, '_selected_before_click', None)
                        if before is not None:
                            try:
                                if int(before) > 0:
                                    self._selected_before_click = None
                                    return
                            except Exception:
                                pass
                        # 原生 HitTest：如果命中某个项，则认为双击是针对项的，跳过 go_up
                        if HAS_PYWIN:
                            try:
                                from PyQt5.QtGui import QCursor
                                gx = QCursor.pos().x()
                                gy = QCursor.pos().y()
                                if self._native_listview_hit_test(gx, gy):
                                    self._selected_before_click = None
                                    return
                            except Exception:
                                pass
                        # 尝试读取当前选中项数量
                        cnt = None
                        try:
                            cnt = self.explorer.dynamicCall('SelectedItems().Count')
                        except Exception:
                            try:
                                sel = self.explorer.dynamicCall('SelectedItems()')
                                if sel is not None:
                                    cnt = sel.property('Count') if hasattr(sel, 'property') else None
                            except Exception:
                                cnt = None
                        # 如果无法确定选中项，或选中为0，则视为空白双击
                        if cnt is None:
                            try:
                                self.go_up(force=True)
                            except Exception:
                                pass
                            return
                        try:
                            if int(cnt) == 0:
                                self.go_up(force=True)
                        except Exception:
                            pass
                        finally:
                            self._selected_before_click = None

                    # try multiple times because folder navigation can be slower;
                    # perform checks at 150ms, 300ms, 600ms before giving up
                    delays = [150, 300, 600]
                    def attempt(idx=0):
                        handled = False
                        try:
                            # reuse the same logic as _check_and_go_up body
                            # check path change first
                            try:
                                before_path = getattr(self, '_path_before_double', None)
                                cur_path = getattr(self, 'current_path', None)
                                if before_path is not None and cur_path is not None and cur_path != before_path:
                                    self._path_before_double = None
                                    return
                            except Exception:
                                pass
                            # if press-time selection existed, skip
                            before = getattr(self, '_selected_before_click', None)
                            if before is not None:
                                try:
                                    if int(before) > 0:
                                        self._selected_before_click = None
                                        return
                                except Exception:
                                    pass
                            # native hit-test
                            if HAS_PYWIN:
                                try:
                                    from PyQt5.QtGui import QCursor
                                    gx = QCursor.pos().x()
                                    gy = QCursor.pos().y()
                                    if self._native_listview_hit_test(gx, gy):
                                        self._selected_before_click = None
                                        return
                                except Exception:
                                    pass
                            # finally check SelectedItems
                            cnt = None
                            try:
                                cnt = self.explorer.dynamicCall('SelectedItems().Count')
                            except Exception:
                                try:
                                    sel = self.explorer.dynamicCall('SelectedItems()')
                                    if sel is not None:
                                        cnt = sel.property('Count') if hasattr(sel, 'property') else None
                                except Exception:
                                    cnt = None
                            if cnt is None:
                                # if there are more attempts left, retry; else treat as blank
                                if idx < len(delays) - 1:
                                    QTimer.singleShot(delays[idx+1] - delays[idx], lambda: attempt(idx+1))
                                    return
                                try:
                                    self.go_up(force=True)
                                except Exception:
                                    pass
                                return
                            try:
                                if int(cnt) == 0:
                                    self.go_up(force=True)
                                return
                            except Exception:
                                pass
                        finally:
                            self._selected_before_click = None

                    QTimer.singleShot(delays[0], lambda: attempt(0))
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def blank_double_click(self, event):
        self.go_up(force=True)

    # 移除 on_document_complete 和 eventFilter 相关内容

    def go_up(self, force=False):
        # 返回上一级目录，盘符根目录时显示所有盘符列表
        # 如果 force=True，则绕过鼠标位置检查（用于按钮或程序化调用）
        if not force:
            # 仅在明确来自空白区域或路径栏的触发时执行，避免误由文件双击触发
            try:
                from PyQt5.QtWidgets import QApplication
                from PyQt5.QtGui import QCursor
                pos = QCursor.pos()
                w = QApplication.widgetAt(pos.x(), pos.y())
                # 允许的触发源：底部空白标签或路径栏
                if w is not self.blank and w is not getattr(self, 'path_bar', None):
                    return
            except Exception:
                # 如果无法判断，保守退出，避免误导航
                return
        if not self.current_path:
            return
        path = self.current_path
        # 判断是否为盘符根目录
        if path.endswith(":\\") or path.endswith(":/"):
            self.show_drive_list()
            return
        parent_path = os.path.dirname(path)
        if parent_path and os.path.exists(parent_path):
            self.navigate_to(parent_path)

    def show_drive_list(self):
        # 列出所有盘符到QListWidget
        import string
        self.drive_list.clear()
        drives = []
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                drives.append(drive)
        if not drives:
            self.drive_list.setVisible(False)
            QMessageBox.information(self, "无可用盘符", "未检测到任何磁盘驱动器。")
            return
        self.drive_list.addItems(drives)
        self.drive_list.setVisible(True)

    def on_drive_double_clicked(self, item):
        drive = item.text()
        self.drive_list.setVisible(False)
        self.navigate_to(drive)
    def __init__(self, parent=None, path="", is_shell=False):
        super().__init__(parent)
        self.main_window = parent
        self.current_path = path if path else QDir.homePath()
        self.setup_ui()
        self.navigate_to(self.current_path, is_shell=is_shell)
        self.start_path_sync_timer()

    # 移除重复的setup_ui，保留带路径栏的实现

    def navigate_to(self, path, is_shell=False):
        # 支持本地路径和shell特殊路径
        if hasattr(self, 'drive_list'):
            self.drive_list.setVisible(False)
        if is_shell:
            self.explorer.dynamicCall("Navigate(const QString&)", path)
            self.current_path = path
            if hasattr(self, 'path_bar'):
                self.path_bar.setText(path)
            self.update_tab_title()
        elif os.path.exists(path):
            url = QDir.toNativeSeparators(path)
            self.explorer.dynamicCall("Navigate(const QString&)", url)
            self.current_path = path
            if hasattr(self, 'path_bar'):
                self.path_bar.setText(path)
            self.update_tab_title()


class DragDropTabWidget(QTabWidget):
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

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if url.isLocalFile():
                    path = url.toLocalFile()
                    if os.path.isdir(path):
                        # 如果是文件夹，打开新标签页
                        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                            self.main_window.add_new_tab(path)
                    elif os.path.isfile(path):
                        # 如果是文件，打开其所在文件夹
                        folder = os.path.dirname(path)
                        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                            self.main_window.add_new_tab(folder)
            event.acceptProposedAction()
        else:
            event.ignore()


class MainWindow(QMainWindow):
    def ensure_default_icons_on_bookmark_bar(self):
        """确保四个常用书签（带图标）始终在最左侧且不会被覆盖。"""
        bm = self.bookmark_manager
        tree = bm.get_tree()
        bar = tree.get('bookmark_bar')
        if not bar or 'children' not in bar:
            return
        # 获取“下载”文件夹路径（跨平台，优先Win用户目录）
        import time
        import os
        from PyQt5.QtCore import QStandardPaths
        downloads_path = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
        if not downloads_path or not os.path.exists(downloads_path):
            # 兜底
            downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
        icon_map = [
            ("🖥️", "此电脑", "shell:MyComputerFolder"),
            ("🗔", "桌面", "shell:Desktop"),
            ("🗑️", "回收站", "shell:RecycleBinFolder"),
            ("🚀", "启动项", "shell:Startup"),
            ("⬇️", "下载", downloads_path),
        ]
        # 移除所有同名（无论有无emoji）
        names_set = set([n for _, n, _ in icon_map])
        bar['children'] = [c for c in bar['children'] if not (c.get('type') == 'url' and any(c.get('name', '').replace(icon, '').strip() == n for icon, n, _ in icon_map))]
        # 插入标准五个项目到最前面
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
            self.go_up_current_tab()


    def go_up_current_tab(self):
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'go_up'):
            current_tab.go_up(force=True)


    def add_new_tab(self, path="", is_shell=False):
        # 默认新建标签页为“此电脑”
        if not path:
            path = 'shell:MyComputerFolder'
            is_shell = True
        tab = FileExplorerTab(self, path, is_shell=is_shell)
        tab.is_pinned = False
        short = path[-8:] if len(path) > 8 else path
        tab_index = self.tab_widget.addTab(tab, short)
        self.tab_widget.setCurrentIndex(tab_index)
        return tab_index


    def close_tab(self, index):
        tab = self.tab_widget.widget(index)
        # 如果是固定标签页，关闭时自动移除固定
        if hasattr(tab, 'is_pinned') and tab.is_pinned:
            tab.is_pinned = False
            self.save_pinned_tabs()
        if self.tab_widget.count() > 1:
            self.tab_widget.removeTab(index)
        else:
            self.close()


    def close_current_tab(self):
        current_index = self.tab_widget.currentIndex()
        self.close_tab(current_index)

    def on_tab_changed(self, index):
        if index >= 0:
            tab = self.tab_widget.widget(index)
            if hasattr(tab, 'current_path'):
                self.setWindowTitle(f"TabExplorer - {tab.current_path}")
                # 展开并选中左侧目录树到当前目录
                self.expand_dir_tree_to_path(tab.current_path)

    def expand_dir_tree_to_path(self, path):
        # 展开并选中左侧目录树到指定路径
        if not hasattr(self, 'dir_model') or not hasattr(self, 'dir_tree'):
            return
        if not path or not os.path.exists(path):
            return
        idx = self.dir_model.index(path)
        if idx.isValid():
            # 递归收集所有父索引
            parents = []
            parent = idx.parent()
            while parent.isValid():
                parents.append(parent)
                parent = parent.parent()
            # 先从根到叶子依次展开
            for p in reversed(parents):
                self.dir_tree.expand(p)
            self.dir_tree.setCurrentIndex(idx)
            self.dir_tree.scrollTo(idx)
    def create_toolbar(self):
        toolbar = QToolBar()
        self.addToolBar(Qt.TopToolBarArea, toolbar)
        # 工具栏保留，可在此添加其它功能按钮
        pass

    def show_bookmark_dialog(self):
        dlg = BookmarkDialog(self.bookmark_manager, self)
        dlg.exec_()

    def tab_context_menu(self, pos):
        tab_index = self.tab_widget.tabBar().tabAt(pos)
        if tab_index < 0:
            return
        tab = self.tab_widget.widget(tab_index)
        is_pinned = hasattr(tab, 'is_pinned') and tab.is_pinned
        menu = QMenu()
        # 图标可用emoji或标准QIcon
        if is_pinned:
            pin_action = QAction("🔨 取消固定", self)
            pin_action.triggered.connect(lambda: self.unpin_tab(tab_index))
            menu.addAction(pin_action)
        else:
            pin_action = QAction("📌 固定", self)
            pin_action.triggered.connect(lambda: self.pin_tab(tab_index))
            menu.addAction(pin_action)

        # 添加“添加书签”菜单项，使用书签emoji
        add_bm_action = QAction("📑 添加书签", self)
        add_bm_action.triggered.connect(lambda: self.add_tab_bookmark(tab))
        menu.addAction(add_bm_action)

        menu.exec_(self.tab_widget.tabBar().mapToGlobal(pos))

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
            QMessageBox.warning(self, "无可用书签文件夹", "请先在 bookmarks.json 中添加至少一个文件夹。")
            return
        # 选择父文件夹
        folder_names = [f"{name} (id:{fid})" for fid, name in folder_list]
        from PyQt5.QtWidgets import QInputDialog
        idx, ok = QInputDialog.getItem(self, "选择书签文件夹", "请选择父文件夹：", folder_names, 0, False)
        if not ok:
            return
        folder_id = folder_list[folder_names.index(idx)][0]
        # 输入书签名称
        name, ok = QInputDialog.getText(self, "书签名称", "请输入书签名称：", text=os.path.basename(tab.current_path))
        if not ok or not name:
            return
        # 保存到 bookmarks.json
        url = "file:///" + tab.current_path.replace("\\", "/")
        if bm.add_bookmark(folder_id, name, url):
            QMessageBox.information(self, "添加成功", "书签已添加！")
            self.populate_bookmark_bar_menu()
        else:
            QMessageBox.warning(self, "添加失败", "未能添加书签，请检查父文件夹。")

    def pin_tab(self, tab_index):
        tab = self.tab_widget.widget(tab_index)
        tab.is_pinned = True
        # 重新排序：所有固定的在最左侧
        self.sort_tabs_by_pinned()
        self.save_pinned_tabs()

    def unpin_tab(self, tab_index):
        tab = self.tab_widget.widget(tab_index)
        tab.is_pinned = False
        self.sort_tabs_by_pinned()
        self.save_pinned_tabs()

    def sort_tabs_by_pinned(self):
        pinned = []
        unpinned = []
        # 记录当前tab对象
        current_index = self.tab_widget.currentIndex()
        current_tab = self.tab_widget.widget(current_index) if current_index >= 0 else None
        for i in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(i)
            if hasattr(tab, 'is_pinned') and tab.is_pinned:
                pinned.append(tab)
            else:
                unpinned.append(tab)
        self.tab_widget.clear()
        new_tabs = pinned + unpinned
        for tab in new_tabs:
            short = tab.current_path[-8:] if len(tab.current_path) > 8 else tab.current_path
            title = ("📌" if getattr(tab, 'is_pinned', False) else "") + short
            self.tab_widget.addTab(tab, title)
        # 恢复原先的tab焦点
        if current_tab is not None:
            for i, tab in enumerate(new_tabs):
                if tab is current_tab:
                    self.tab_widget.setCurrentIndex(i)
                    break

    def save_pinned_tabs(self):
        pinned_paths = []
        for i in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(i)
            if hasattr(tab, 'is_pinned') and tab.is_pinned:
                if hasattr(tab, 'current_path'):
                    pinned_paths.append(tab.current_path)
        with open("pinned_tabs.json", "w", encoding="utf-8") as f:
            json.dump(pinned_paths, f, ensure_ascii=False, indent=2)

    def load_pinned_tabs(self):
        has_pinned = False
        if os.path.exists("pinned_tabs.json"):
            try:
                with open("pinned_tabs.json", "r", encoding="utf-8") as f:
                    paths = json.load(f)
                for path in paths:
                    tab = FileExplorerTab(self, path)
                    tab.is_pinned = True
                    short = path[-8:] if len(path) > 8 else path
                    title = "📌" + short
                    self.tab_widget.addTab(tab, title)
                    has_pinned = True
            except Exception:
                pass
        return has_pinned

    def __init__(self, parent=None):
        super().__init__(parent)
        self.bookmark_manager = BookmarkManager()
        # 检查并自动添加常用书签
        self.ensure_default_bookmarks()
        self.init_ui()

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
                "name": "书签栏",
                "type": "folder",
                "children": []
            }
        bar = tree['bookmark_bar']
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
                make_bm("此电脑", "shell:MyComputerFolder", "🖥️"),
                make_bm("桌面", "shell:Desktop", "🗔"),
                make_bm("回收站", "shell:RecycleBinFolder", "🗑️"),
                make_bm("启动项", "shell:Startup", "🚀"),
            ]
            bm.save_bookmarks()


    def init_ui(self):
        from PyQt5.QtWidgets import QSplitter, QTreeView, QFileSystemModel
        self.setWindowTitle("TabExplorer")
        self.setGeometry(100, 100, 1200, 800)

        # 直接将“收藏夹”里的书签全部列在菜单栏顶层
        self.menuBar().clear()
        self.populate_bookmark_bar_menu()

        # 创建工具栏
        self.create_toolbar()

        # 主分割器，左树右标签
        self.splitter = QSplitter()
        self.splitter.setOrientation(Qt.Horizontal)

        # 左侧目录树
        self.dir_model = QFileSystemModel()
        # 设置根为计算机根目录（显示所有盘符）
        root_path = QDir.rootPath()  # 通常为C:/
        self.dir_model.setRootPath(root_path)
        self.dir_tree = QTreeView()
        self.dir_tree.setModel(self.dir_model)
        # 不设置setRootIndex，或者设置为index("")，这样能显示所有盘符
        # self.dir_tree.setRootIndex(self.dir_model.index(""))
        self.dir_tree.setHeaderHidden(True)
        self.dir_tree.setColumnHidden(1, True)
        self.dir_tree.setColumnHidden(2, True)
        self.dir_tree.setColumnHidden(3, True)
        self.dir_tree.setMinimumWidth(200)
        self.dir_tree.setMaximumWidth(350)
        self.dir_tree.clicked.connect(self.on_dir_tree_clicked)
        self.splitter.addWidget(self.dir_tree)
        # 自动展开所有驱动器根节点（即“我的电脑”下所有盘符）
        drives_parent = self.dir_model.index(root_path)
        for i in range(self.dir_model.rowCount(drives_parent)):
            idx = self.dir_model.index(i, 0, drives_parent)
            self.dir_tree.expand(idx)

        # 右侧原有标签页区域
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # 创建标签页控件（支持拖放）
        self.tab_widget = DragDropTabWidget(self)
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.tab_widget.currentChanged.connect(self.on_tab_changed)

        # 设置选中标签页背景色为淡黄色
        tabbar = self.tab_widget.tabBar()
        tabbar.setStyleSheet("""
            QTabBar::tab {
                border-right: 1px solid #d3d3d3;
                padding-right: 12px;
                padding-left: 12px;
                min-height: 36px;
                font-size: 15px;
            }
            QTabBar::tab:selected {
                background: #FFF9CC;
            }
        """)
        # 绑定双击事件
        tabbar.mouseDoubleClickEvent = self.tabbar_mouse_double_click

        # 添加新标签页按钮和上一级按钮
        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        self.up_button = QPushButton("⬆️")
        self.up_button.setToolTip("上一级目录")
        self.up_button.setFixedHeight(35)
        self.up_button.clicked.connect(self.go_up_current_tab)
        btn_layout.addWidget(self.up_button)
        self.add_tab_button = QPushButton("➕")
        self.add_tab_button.setToolTip("新建标签页")
        self.add_tab_button.setFixedHeight(35)
        self.add_tab_button.clicked.connect(self.add_new_tab)
        btn_layout.addWidget(self.add_tab_button)
        btn_layout.addStretch(1)
        self.tab_widget.setCornerWidget(btn_widget)
        # 双击角落（标签区与按钮之间的空白）等效返回上一级
        btn_widget.mouseDoubleClickEvent = lambda e: self.go_up_current_tab()

        right_layout.addWidget(self.tab_widget)
        self.splitter.addWidget(right_widget)
        self.setCentralWidget(self.splitter)

        # 加载固定标签页
        has_pinned = self.load_pinned_tabs()
        # 添加初始标签页（如无固定标签）
        if not has_pinned and self.tab_widget.count() == 0:
            self.add_new_tab(QDir.homePath())
        # 右键标签页支持固定/取消固定
        self.tab_widget.tabBar().setContextMenuPolicy(Qt.CustomContextMenu)
        self.tab_widget.tabBar().customContextMenuRequested.connect(self.tab_context_menu)

    def on_dir_tree_clicked(self, index):
        # 目录树点击，右侧当前标签页跳转
        if not index.isValid():
            return
        path = self.dir_model.filePath(index)
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'navigate_to'):
            current_tab.navigate_to(path)

    def open_bookmark_url(self, url):
        # 支持 file:///、shell: 路径和本地绝对路径
        from urllib.parse import unquote
        if url.startswith('file:///'):
            local_path = unquote(url[8:])
            if os.name == 'nt' and local_path.startswith('/'):
                local_path = local_path[1:]
            # 检查是否是 shell: 路径
            if local_path.startswith('shell:'):
                self.add_new_tab(local_path, is_shell=True)
            elif os.path.exists(local_path):
                self.add_new_tab(local_path)
            else:
                QMessageBox.warning(self, "路径错误", f"路径不存在: {local_path}")
        elif url.startswith('shell:'):
            self.add_new_tab(url, is_shell=True)
        elif os.path.isabs(url) and os.path.exists(url):
            self.add_new_tab(url)
        else:
            QMessageBox.warning(self, "不支持的书签", f"暂不支持打开此类型书签: {url}")

    def populate_bookmark_bar_menu(self):
        self.ensure_default_icons_on_bookmark_bar()
        self.menuBar().clear()
        bm = self.bookmark_manager
        tree = bm.get_tree()
        bookmark_bar = tree.get('bookmark_bar')
        if not bookmark_bar or 'children' not in bookmark_bar:
            return
        def add_menu_items(parent_menu, node):
            if node.get('type') == 'folder':
                menu = parent_menu.addMenu(f"📁 {node.get('name', '')}")
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
        # 直接在菜单栏顶层添加
        menubar = self.menuBar()
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
        # 添加“书签管理”按钮到菜单栏最右侧（兼容所有系统样式）
        # 方案：添加一个空菜单右对齐，再添加“书签管理”action
        menubar.addSeparator()
        actions = menubar.actions()
        if actions:
            menubar.insertSeparator(actions[-1])
        manage_action = QAction("书签管理", self)
        manage_action.triggered.connect(self.show_bookmark_manager_dialog)
        menubar.addAction(manage_action)

    def show_bookmark_manager_dialog(self):
        dlg = BookmarkManagerDialog(self.bookmark_manager, self)
        dlg.exec_()

# 书签管理对话框（初步框架，后续可扩展重命名/新建/删除等功能）
from PyQt5.QtWidgets import QDialog, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QPushButton, QHBoxLayout, QInputDialog, QMessageBox
class BookmarkManagerDialog(QDialog):
    def __init__(self, bookmark_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("书签管理")
        self.resize(600, 500)
        self.bookmark_manager = bookmark_manager
        layout = QVBoxLayout(self)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["名称", "类型", "路径"])
        self.tree.setColumnWidth(0, 250)  # 第一列宽一些
        layout.addWidget(self.tree)
        self.populate_tree()

        btn_layout = QHBoxLayout()
        # self.rename_btn = QPushButton("重命名")  # 已移除重命名按钮
        # self.rename_btn.clicked.connect(self.rename_item)
        # btn_layout.addWidget(self.rename_btn)
        self.edit_btn = QPushButton("编辑")
        self.edit_btn.clicked.connect(self.edit_item)
        btn_layout.addWidget(self.edit_btn)
        self.delete_btn = QPushButton("删除")
        self.delete_btn.clicked.connect(self.delete_item)
        btn_layout.addWidget(self.delete_btn)
        self.new_folder_btn = QPushButton("新建文件夹")
        self.new_folder_btn.clicked.connect(self.create_folder)
        btn_layout.addWidget(self.new_folder_btn)
        self.up_btn = QPushButton("上移")
        self.up_btn.clicked.connect(self.move_item_up)
        btn_layout.addWidget(self.up_btn)
        self.down_btn = QPushButton("下移")
        self.down_btn.clicked.connect(self.move_item_down)
        btn_layout.addWidget(self.down_btn)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def move_item_up(self):
        item = self.tree.currentItem()
        if not item:
            QMessageBox.warning(self, "未选择", "请先选择要上移的书签或文件夹。")
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
            QMessageBox.warning(self, "未选择", "请先选择要下移的书签或文件夹。")
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
    def __init__(self, bookmark_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("书签管理")
        self.resize(600, 500)
        self.bookmark_manager = bookmark_manager
        layout = QVBoxLayout(self)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["名称", "类型", "路径"])
        self.tree.setColumnWidth(0, 250)  # 第一列宽一些
        layout.addWidget(self.tree)
        self.populate_tree()

        btn_layout = QHBoxLayout()
        # self.rename_btn = QPushButton("重命名")  # 已移除重命名按钮
        # self.rename_btn.clicked.connect(self.rename_item)
        # btn_layout.addWidget(self.rename_btn)
        self.edit_btn = QPushButton("编辑")
        self.edit_btn.clicked.connect(self.edit_item)
        btn_layout.addWidget(self.edit_btn)
        self.delete_btn = QPushButton("删除")
        self.delete_btn.clicked.connect(self.delete_item)
        btn_layout.addWidget(self.delete_btn)
        self.new_folder_btn = QPushButton("新建文件夹")
        self.new_folder_btn.clicked.connect(self.create_folder)
        btn_layout.addWidget(self.new_folder_btn)
        self.up_btn = QPushButton("上移")
        self.up_btn.clicked.connect(self.move_item_up)
        btn_layout.addWidget(self.up_btn)
        self.down_btn = QPushButton("下移")
        self.down_btn.clicked.connect(self.move_item_down)
        btn_layout.addWidget(self.down_btn)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
    def edit_item(self):
        item = self.tree.currentItem()
        if not item:
            QMessageBox.warning(self, "未选择", "请先选择要编辑的书签或文件夹。")
            return
        node_type = item.text(1)
        old_name = item.text(0).lstrip("📁 ").lstrip("📑 ")
        main_window = self.parent() if self.parent() and hasattr(self.parent(), 'populate_bookmark_bar_menu') else None
        if node_type == '文件夹':
            new_name, ok = QInputDialog.getText(self, "编辑文件夹", "请输入新名称：", text=old_name)
            if ok and new_name and new_name != old_name:
                item.setText(0, f"📁 {new_name}")
                self.update_name_in_bookmark_manager(item, new_name)
                self.bookmark_manager.save_bookmarks()
                self.populate_tree()
                if main_window:
                    main_window.populate_bookmark_bar_menu()
        elif node_type == '书签':
            new_name, ok1 = QInputDialog.getText(self, "编辑书签", "请输入新名称：", text=old_name)
            old_url = item.text(2)
            new_url, ok2 = QInputDialog.getText(self, "编辑书签", "请输入新路径：", text=old_url)
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
            QMessageBox.warning(self, "未选择", "请先选择要删除的书签或文件夹。")
            return
        node_id = item.data(0, 1)
        reply = QMessageBox.question(self, "确认删除", "确定要删除选中的项目及其所有子项吗？", QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
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
                item = QTreeWidgetItem([f"📁 {node.get('name', '')}", '文件夹', ''])
                item.setData(0, 1, node.get('id'))
                if parent_item:
                    parent_item.addChild(item)
                else:
                    self.tree.addTopLevelItem(item)
                for child in node.get('children', []):
                    add_node(item, child)
            elif node.get('type') == 'url':
                item = QTreeWidgetItem([f"📑 {node.get('name', '')}", '书签', node.get('url', '')])
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
            QMessageBox.warning(self, "未选择", "请先选择要重命名的书签或文件夹。")
            return
        old_name = item.text(0)
        new_name, ok = QInputDialog.getText(self, "重命名", "请输入新名称：", text=old_name)
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
        if item and item.text(1) == '文件夹':
            parent_id = item.data(0, 1)
        else:
            # 默认加到bookmark_bar根
            parent_id = self.bookmark_manager.get_tree().get('bookmark_bar', {}).get('id')
        folder_name, ok = QInputDialog.getText(self, "新建文件夹", "请输入文件夹名称：")
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


    # BookmarkManagerDialog不再包含标签页相关方法

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("TabExplorer")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

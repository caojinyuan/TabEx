"""
资源管理器监控模块
监控新打开的 Windows Explorer 窗口，获取路径并在 TabExplorer 中打开
"""

import time
import threading
import win32gui
import win32process
import win32con
from collections import defaultdict

try:
    import win32com.client
    HAS_COM = True
except:
    HAS_COM = False


class ExplorerMonitor:
    def __init__(self, main_window):
        self.main_window = main_window
        self.monitoring = False
        self.monitor_thread = None
        self.known_windows = set()  # 已知的窗口句柄
        self.window_paths = {}  # 窗口句柄 -> 路径映射
        
    def start(self):
        """启动监控"""
        if self.monitoring:
            return
        
        # 在启动前记录所有已存在的Explorer窗口
        def record_existing(hwnd, _):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name in ["CabinetWClass", "ExploreWClass"]:
                        self.known_windows.add(hwnd)
            except:
                pass
            return True
        
        win32gui.EnumWindows(record_existing, None)
        print(f"[ExplorerMonitor] 启动时已记录 {len(self.known_windows)} 个现有Explorer窗口")
        
        self.monitoring = True
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()
        print("[ExplorerMonitor] 监控已启动")
    
    def stop(self):
        """停止监控"""
        self.monitoring = False
        print("[ExplorerMonitor] 监控已停止")
    
    def _monitor_loop(self):
        """监控循环"""
        # 在线程中初始化COM
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception as e:
            print(f"[ExplorerMonitor] COM初始化失败: {e}")
        
        try:
            while self.monitoring:
                try:
                    self._check_new_explorer_windows()
                    time.sleep(0.2)  # 改为0.2秒，提高响应速度
                except Exception as e:
                    print(f"[ExplorerMonitor] 监控错误: {e}")
                    time.sleep(1)
        finally:
            # 线程结束时清理COM
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _check_new_explorer_windows(self):
        """检查新的资源管理器窗口"""
        current_windows = set()
        new_windows = []  # 新发现的窗口列表
        
        # 定期输出检查状态（每5秒）
        if not hasattr(self, '_check_counter'):
            self._check_counter = 0
        self._check_counter += 1
        if self._check_counter % 25 == 0:  # 0.2s * 25 = 5s
            print(f"[ExplorerMonitor] 监控运行中... (已检查 {self._check_counter} 次)")
        
        def enum_callback(hwnd, _):
            try:
                # 检查窗口是否可见
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                
                # 获取窗口类名
                class_name = win32gui.GetClassName(hwnd)
                
                # Explorer 窗口的类名是 "CabinetWClass" 或 "ExploreWClass"
                if class_name not in ["CabinetWClass", "ExploreWClass"]:
                    return True
                
                # 获取窗口标题（新窗口可能标题为空，不过滤）
                title = win32gui.GetWindowText(hwnd)
                
                # 记录当前窗口
                current_windows.add(hwnd)
                
                # 检查是否是新窗口
                if hwnd not in self.known_windows:
                    print(f"[ExplorerMonitor] 发现新窗口: '{title}' (hwnd={hwnd}, class={class_name})")
                    new_windows.append((hwnd, title))
                
            except Exception as e:
                print(f"[ExplorerMonitor] 枚举窗口错误: {e}")
            
            return True
        
        # 枚举所有顶层窗口
        win32gui.EnumWindows(enum_callback, None)
        
        # 更新已知窗口列表
        self.known_windows = current_windows
        
        # 处理新窗口（在枚举完成后处理，避免在枚举过程中修改窗口）
        for hwnd, title in new_windows:
            self._handle_new_window(hwnd, title)
    
    def _handle_new_window(self, hwnd, title):
        """处理新打开的资源管理器窗口 - 直接嵌入该窗口而不是创建新标签"""
        try:
            # 延迟并重试获取路径（Win+E 打开的窗口需要更长时间初始化）
            path = None
            max_retries = 5
            
            for retry in range(max_retries):
                # 等待时间递增
                wait_time = 0.1 + (retry * 0.1)  # 0.1, 0.2, 0.3, 0.4, 0.5
                time.sleep(wait_time)
                
                print(f"[ExplorerMonitor] 尝试获取路径 (第{retry+1}次，等待{wait_time}秒)...")
                path = self._get_explorer_path(hwnd)
                
                if path:
                    break
                    
                if retry < max_retries - 1:
                    print(f"[ExplorerMonitor] 路径为空，等待窗口初始化...")
            
            if path:
                print(f"[ExplorerMonitor] 最终获取到路径: {path}")
                
                # 判断是否是 shell: 路径
                is_shell = path.startswith('shell:') or '::' in path
                
                # 在主线程中直接嵌入该Explorer窗口到新标签页
                from PyQt5.QtCore import QMetaObject, Qt, Q_ARG
                QMetaObject.invokeMethod(
                    self.main_window,
                    "embed_existing_explorer",
                    Qt.QueuedConnection,
                    Q_ARG(int, hwnd),
                    Q_ARG(str, path)
                )
                
                # 不再关闭原窗口，因为我们要嵌入它
                # self._close_explorer_window(hwnd)
            else:
                print(f"[ExplorerMonitor] 重试{max_retries}次后仍无法获取路径，窗口标题: {title}")
                print(f"[ExplorerMonitor] 尝试根据标题猜测路径...")
                
                # 根据标题猜测路径
                fallback_path = self._guess_path_from_title(title)
                if fallback_path:
                    print(f"[ExplorerMonitor] 使用备用路径: {fallback_path}")
                    from PyQt5.QtCore import QMetaObject, Qt, Q_ARG
                    QMetaObject.invokeMethod(
                        self.main_window,
                        "embed_existing_explorer",
                        Qt.QueuedConnection,
                        Q_ARG(int, hwnd),
                        Q_ARG(str, fallback_path)
                    )
                
        except Exception as e:
            print(f"[ExplorerMonitor] 处理新窗口错误: {e}")
            import traceback
            traceback.print_exc()
    
    def _guess_path_from_title(self, title):
        """根据窗口标题猜测路径"""
        title_map = {
            '此电脑': 'shell:MyComputerFolder',
            '快速访问': 'shell:MyComputerFolder',
            '回收站': 'shell:RecycleBinFolder',
            '桌面': 'shell:Desktop',
            '网络': 'shell:NetworkPlacesFolder',
            '下载': 'shell:Downloads',
            '文档': 'shell:Personal',
            '图片': 'shell:My Pictures',
            '音乐': 'shell:My Music',
            '视频': 'shell:My Video',
        }
        
        for key, value in title_map.items():
            if key in title:
                return value
        
        return None
    
    def _get_explorer_path(self, hwnd):
        """获取资源管理器窗口的路径"""
        if not HAS_COM:
            return None
        
        try:
            # 使用 Shell.Application COM 对象获取所有窗口
            shell = win32com.client.Dispatch("Shell.Application")
            windows = shell.Windows()
            
            for window in windows:
                try:
                    # 获取窗口句柄
                    window_hwnd = window.HWND
                    
                    if window_hwnd == hwnd:
                        # 获取当前路径
                        try:
                            location = window.LocationURL
                        except Exception as e:
                            print(f"[ExplorerMonitor] 获取LocationURL失败: {e}")
                            # 尝试使用 LocationName
                            try:
                                location_name = window.LocationName
                                print(f"[ExplorerMonitor] LocationName: {location_name}")
                                # 如果是特殊文件夹名称，返回 None 让调用者处理
                                return None
                            except:
                                return None
                        
                        # 处理空 URL
                        if not location or location.strip() == '':
                            print(f"[ExplorerMonitor] URL为空，尝试获取Document...")
                            try:
                                # 尝试通过 Document 获取路径
                                doc = window.Document
                                if hasattr(doc, 'Folder'):
                                    folder = doc.Folder
                                    if hasattr(folder, 'Self'):
                                        path_item = folder.Self
                                        if hasattr(path_item, 'Path'):
                                            doc_path = path_item.Path
                                            print(f"[ExplorerMonitor] 通过Document获取路径: {doc_path}")
                                            
                                            # 转换CLSID为shell路径
                                            if '::' in doc_path:
                                                clsid_map = {
                                                    '::{20D04FE0-3AEA-1069-A2D8-08002B30309D}': 'shell:MyComputerFolder',
                                                    '::{645FF040-5081-101B-9F08-00AA002F954E}': 'shell:RecycleBinFolder',
                                                    '::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}': 'shell:NetworkPlacesFolder',
                                                }
                                                for clsid, shell_path in clsid_map.items():
                                                    if clsid in doc_path:
                                                        print(f"[ExplorerMonitor] CLSID转换为: {shell_path}")
                                                        return shell_path
                                            
                                            return doc_path
                            except Exception as e:
                                print(f"[ExplorerMonitor] 通过Document获取路径失败: {e}")
                            return None
                        
                        print(f"[ExplorerMonitor] 原始URL: {location}")
                        
                        # 转换 file:/// URL 到本地路径
                        if location.startswith('file:///'):
                            from urllib.parse import unquote
                            path = unquote(location[8:])  # 移除 'file:///'
                            
                            # Windows 路径处理
                            if path.startswith('/'):
                                path = path[1:]  # 移除开头的 /
                            
                            path = path.replace('/', '\\')
                            print(f"[ExplorerMonitor] 转换后路径: {path}")
                            return path
                        
                        # 处理特殊路径（如"快速访问"、"此电脑"等）
                        elif '::' in location or location.startswith('shell:'):
                            # CLSID 格式的特殊路径
                            print(f"[ExplorerMonitor] 检测到特殊路径: {location}")
                            
                            # 尝试将常见的 CLSID 转换为 shell: 路径
                            clsid_map = {
                                '::{20D04FE0-3AEA-1069-A2D8-08002B30309D}': 'shell:MyComputerFolder',  # 此电脑
                                '::{645FF040-5081-101B-9F08-00AA002F954E}': 'shell:RecycleBinFolder',  # 回收站
                                '::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}': 'shell:NetworkPlacesFolder',  # 网络
                            }
                            
                            for clsid, shell_path in clsid_map.items():
                                if clsid in location:
                                    print(f"[ExplorerMonitor] 转换为: {shell_path}")
                                    return shell_path
                            
                            # 对于"快速访问"等，默认打开"此电脑"
                            if 'home' in location.lower() or location == '':
                                print(f"[ExplorerMonitor] 快速访问或空路径，使用此电脑")
                                return 'shell:MyComputerFolder'
                            
                            return location
                        
                        # 其他情况尝试直接返回
                        else:
                            print(f"[ExplorerMonitor] 未知格式，直接返回: {location}")
                            return location
                
                except Exception as e:
                    print(f"[ExplorerMonitor] 处理窗口时出错: {e}")
                    continue
            
            print(f"[ExplorerMonitor] 在所有窗口中未找到匹配的句柄: {hwnd}")
            
        except Exception as e:
            print(f"[ExplorerMonitor] 获取路径错误: {e}")
            import traceback
            traceback.print_exc()
        
        return None
    
    def _close_explorer_window(self, hwnd):
        """关闭资源管理器窗口"""
        try:
            # 发送关闭消息
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            print(f"[ExplorerMonitor] 已关闭窗口 (hwnd={hwnd})")
        except Exception as e:
            print(f"[ExplorerMonitor] 关闭窗口错误: {e}")


# ============================================================================
# 书签管理对话框
# ============================================================================

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

# 自定义委托：在文件名列实现省略号在开头
from PyQt5.QtWidgets import QStyledItemDelegate
from PyQt5.QtCore import Qt
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

# 搜索对话框
from PyQt5.QtCore import pyqtSignal as _pyqtSignal
class SearchDialog(QDialog):    
    def __init__(self, search_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"搜索 - {search_path}")
        self.resize(800, 500)
        # 设置窗口标志：可调整大小，带最大化/最小化按钮
        from PyQt5.QtCore import Qt
        self.setWindowFlags(Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.search_path = search_path
        self.main_window = parent
        self.search_thread = None
        self.is_searching = False
        
        # 线程安全的结果队列
        import queue
        self.result_queue = queue.Queue()
        self.ui_update_timer = None
        
        layout = QVBoxLayout(self)
        
        # 搜索选项区域
        search_options = QHBoxLayout()
        
        # 搜索关键词
        search_options.addWidget(QLabel("搜索:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入搜索关键词...")
        self.search_input.returnPressed.connect(self.start_search)
        search_options.addWidget(self.search_input)
        
        # 搜索按钮
        self.search_btn = QPushButton("🔍 搜索")
        self.search_btn.clicked.connect(self.start_search)
        search_options.addWidget(self.search_btn)
        
        # 停止按钮
        self.stop_btn = QPushButton("⏹ 停止")
        self.stop_btn.clicked.connect(self.stop_search)
        self.stop_btn.setEnabled(False)
        search_options.addWidget(self.stop_btn)
        
        layout.addLayout(search_options)
        
        # 搜索路径输入框（可编辑）
        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel("搜索路径:"))
        self.path_input = QLineEdit(search_path)
        self.path_input.setStyleSheet("QLineEdit { color: #0066cc; font-weight: bold; padding: 5px; }")
        self.path_input.setPlaceholderText("输入要搜索的文件夹路径...")
        path_layout.addWidget(self.path_input)
        layout.addLayout(path_layout)
        
        # 搜索类型选择
        type_options = QHBoxLayout()
        self.search_filename_cb = QCheckBox("搜索文件名")
        self.search_filename_cb.setChecked(True)
        type_options.addWidget(self.search_filename_cb)
        
        self.search_content_cb = QCheckBox("搜索文件内容")
        self.search_content_cb.setChecked(True)  # 默认也选中
        type_options.addWidget(self.search_content_cb)
        
        type_options.addStretch(1)
        layout.addLayout(type_options)
        
        # 文件类型过滤
        file_type_layout = QHBoxLayout()
        file_type_layout.addWidget(QLabel("文件类型:"))
        self.file_type_input = QLineEdit()
        self.file_type_input.setPlaceholderText("例如: *.c,*.h,*.xml (留空表示搜索所有类型)")
        self.file_type_input.setText("*.c,*.h,*.xdm,*.arxml,*.xml")  # 默认值
        self.file_type_input.setStyleSheet("QLineEdit { padding: 5px; }")
        file_type_layout.addWidget(self.file_type_input)
        layout.addLayout(file_type_layout)
        
        # 状态标签
        self.status_label = QLabel("就绪")
        layout.addWidget(self.status_label)
        
        # 结果表格
        self.result_list = QTableWidget()
        self.result_list.setColumnCount(4)
        self.result_list.setHorizontalHeaderLabels(["文件名", "类型", "修改日期", "大小"])
        self.result_list.horizontalHeader().setStretchLastSection(False)
        self.result_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.result_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.result_list.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.result_list.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.result_list.setSelectionBehavior(QTableWidget.SelectRows)
        self.result_list.setEditTriggers(QTableWidget.NoEditTriggers)
        self.result_list.cellDoubleClicked.connect(self.on_result_double_clicked)
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
        self.ui_update_timer.start(100)  # 每100ms检查一次队列
    
    def update_ui_from_queue(self):
        """从队列中取出结果并更新UI（在主线程中调用）"""
        try:
            while True:
                item = self.result_queue.get_nowait()
                if item['type'] == 'result':
                    # 添加表格行（排序时暂时禁用以提高性能）
                    sorting_enabled = self.result_list.isSortingEnabled()
                    self.result_list.setSortingEnabled(False)
                    
                    row = self.result_list.rowCount()
                    self.result_list.insertRow(row)
                    # 文件名项 - 使用省略号在开头
                    name_item = QTableWidgetItem(item['name'])
                    # 设置文本省略模式：在开头显示省略号，优先显示文件名
                    from PyQt5.QtCore import Qt
                    name_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    name_item.setToolTip(item['full_path'])  # 添加完整路径的提示
                    self.result_list.setItem(row, 0, name_item)
                    self.result_list.setItem(row, 1, QTableWidgetItem(item['file_type']))
                    self.result_list.setItem(row, 2, QTableWidgetItem(item['date']))
                    self.result_list.setItem(row, 3, QTableWidgetItem(item['size']))
                    # 存储完整路径到第一列的data中
                    self.result_list.item(row, 0).setData(256, item['path'])
                    
                    # 恢复排序状态
                    self.result_list.setSortingEnabled(sorting_enabled)
                elif item['type'] == 'status':
                    self.status_label.setText(item['text'])
                elif item['type'] == 'button':
                    if item['button'] == 'search':
                        self.search_btn.setEnabled(item['enabled'])
                    elif item['button'] == 'stop':
                        self.stop_btn.setEnabled(item['enabled'])
        except:
            pass  # 队列为空
    
    def add_search_result(self, text):
        """添加搜索结果项（通过队列，线程安全）"""
        self.result_queue.put({'type': 'result', 'text': text})
    
    def start_search(self):
        keyword = self.search_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "提示", "请输入搜索关键词")
            return
        
        if not self.search_filename_cb.isChecked() and not self.search_content_cb.isChecked():
            QMessageBox.warning(self, "提示", "请至少选择一种搜索类型")
            return
        
        # 获取并验证搜索路径
        search_path = self.path_input.text().strip()
        if not search_path:
            QMessageBox.warning(self, "提示", "请输入搜索路径")
            return
        
        # 检查路径是否存在
        if not os.path.exists(search_path):
            QMessageBox.warning(self, "路径错误", f"路径不存在:\n{search_path}")
            return
        
        # 检查是否是目录
        if not os.path.isdir(search_path):
            QMessageBox.warning(self, "路径错误", f"路径不是文件夹:\n{search_path}")
            return
        
        # 检查是否是特殊路径（不支持搜索）
        if search_path.startswith('shell:'):
            QMessageBox.warning(self, "不支持", "不支持搜索特殊路径（shell:）")
            return
        
        # 更新搜索路径
        self.search_path = search_path
        
        # 清空之前的结果
        self.result_list.setRowCount(0)
        self.is_searching = True
        self.search_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("搜索中...")
        
        # 获取文件类型过滤
        file_types = self.file_type_input.text().strip()
        
        # 在后台线程执行搜索
        import threading
        self.search_thread = threading.Thread(
            target=self.do_search,
            args=(keyword, self.search_filename_cb.isChecked(), self.search_content_cb.isChecked(), file_types)
        )
        self.search_thread.daemon = True
        self.search_thread.start()
    
    def stop_search(self):
        self.is_searching = False
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("已停止")
    
    def do_search(self, keyword, search_filename, search_content, file_types=""):
        found_count = 0
        keyword_lower = keyword.lower()
        results_buffer = []  # 结果缓冲区
        buffer_size = 20  # 每20个结果批量更新一次
        
        # 解析文件类型过滤（支持*.ext格式，逗号分隔）
        file_extensions = []
        if file_types:
            for ft in file_types.split(','):
                ft = ft.strip()
                if ft.startswith('*.'):
                    file_extensions.append(ft[2:].lower())  # 去掉*.，只保留扩展名
                elif ft.startswith('.'):
                    file_extensions.append(ft[1:].lower())  # 去掉.，只保留扩展名
                elif ft:
                    file_extensions.append(ft.lower())  # 直接使用输入的扩展名
        
        # 调试信息：输出搜索路径
        print(f"[Search] 开始搜索路径: {self.search_path}")
        print(f"[Search] 搜索关键词: {keyword}")
        print(f"[Search] 搜索文件名: {search_filename}, 搜索内容: {search_content}")
        print(f"[Search] 文件类型过滤: {file_extensions if file_extensions else '所有类型'}")
        
        def matches_file_type(filename):
            """检查文件是否匹配文件类型过滤"""
            if not file_extensions:  # 如果没有设置过滤，匹配所有文件
                return True
            # 获取文件扩展名（不含点）
            _, ext = os.path.splitext(filename)
            if ext:
                ext = ext[1:].lower()  # 去掉点号并转为小写
                return ext in file_extensions
            return False
        
        try:
            scanned_files = 0
            folder_count = 0
            for root, dirs, files in os.walk(self.search_path):
                if not self.is_searching:
                    print("[Search] 搜索被中断")
                    break
                
                folder_count += 1
                # 每处理10个文件夹更新一次状态（减少更新频率）
                if folder_count % 10 == 0:
                    # 通过队列更新状态
                    status_text = f"搜索中... 已扫描 {scanned_files} 个文件，找到 {found_count} 个结果"
                    self.result_queue.put({'type': 'status', 'text': status_text})
                
                # 搜索文件夹名
                if search_filename:
                    for dirname in dirs:
                        if not self.is_searching:
                            break
                        
                        if keyword_lower in dirname.lower():
                            found_count += 1
                            dir_path = os.path.join(root, dirname)
                            
                            # 获取文件夹信息
                            try:
                                stat_info = os.stat(dir_path)
                                mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                                size_str = "-"  # 文件夹不显示大小
                            except:
                                mtime = "-"
                                size_str = "-"
                            
                            results_buffer.append({
                                'path': dir_path,
                                'name': f"📁 {dirname}",
                                'full_path': f"📁 {dir_path}",
                                'file_type': '文件夹',
                                'date': mtime,
                                'size': size_str
                            })
                            
                            # 批量更新UI
                            if len(results_buffer) >= buffer_size:
                                for item in results_buffer:
                                    self.result_queue.put({'type': 'result', **item})
                                results_buffer.clear()
                
                # 搜索文件名和文件内容
                for filename in files:
                    if not self.is_searching:
                        print("[Search] 搜索被中断（文件循环）")
                        break
                        break
                    
                    # 检查文件类型过滤
                    if not matches_file_type(filename):
                        # 调试：显示被过滤的文件（仅对特定文件名）
                        if 'TstMgr' in filename or scanned_files < 5:
                            print(f"[Search] 文件被类型过滤跳过: {filename}")
                        continue  # 跳过不匹配的文件类型
                    
                    scanned_files += 1
                    file_path = os.path.join(root, filename)
                    matched = False
                    match_type = ""
                    
                    # 调试：显示正在搜索的特定文件
                    if 'TstMgr_RtnSound.c' in filename:
                        print(f"[Search] 正在搜索文件: {file_path}")
                        print(f"[Search] 搜索文件名: {search_filename}, 搜索内容: {search_content}")
                    
                    # 搜索文件名
                    if search_filename and keyword_lower in filename.lower():
                        matched = True
                        match_type = "📄"
                    
                    # 搜索文件内容（不管文件名是否匹配，只要勾选了搜索内容就搜索）
                    if search_content and not matched:
                        # 调试信息
                        if 'TstMgr_RtnSound.c' in filename:
                            print(f"[Search] 开始搜索文件内容: {file_path}")
                        
                        try:
                            # 分块读取大文件，每次读取100MB
                            chunk_size = 100 * 1024 * 1024  # 100MB
                            file_size = os.path.getsize(file_path)
                            
                            # 尝试多种编码方式读取文件内容
                            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                            content_matched = False
                            
                            for encoding in encodings:
                                try:
                                    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                                        if file_size <= chunk_size:
                                            # 小文件直接全部读取
                                            content = f.read()
                                            if keyword_lower in content.lower():
                                                matched = True
                                                match_type = "📄"
                                                content_matched = True
                                                # 调试信息
                                                if 'TstMgr_RtnSound.c' in filename:
                                                    print(f"[Search] ✓ 在文件内容中找到关键词 (编码: {encoding})")
                                                break
                                        else:
                                            # 大文件分块读取
                                            overlap = len(keyword) * 2  # 重叠区域，防止关键词被分割
                                            while True:
                                                chunk = f.read(chunk_size)
                                                if not chunk:
                                                    break
                                                if keyword_lower in chunk.lower():
                                                    matched = True
                                                    match_type = "📄"
                                                    content_matched = True
                                                    break
                                                # 回退overlap字节，避免关键词跨块
                                                if len(chunk) == chunk_size:
                                                    f.seek(f.tell() - overlap)
                                            if content_matched:
                                                break
                                except UnicodeDecodeError:
                                    # 尝试下一个编码
                                    continue
                                except Exception as e:
                                    # 其他错误，记录日志并尝试下一个编码
                                    print(f"[Search] 读取文件失败 {file_path} (编码 {encoding}): {e}")
                                    continue
                        except Exception as e:
                            # 如果无法以文本方式读取，记录日志并跳过该文件
                            print(f"[Search] 无法读取文件 {file_path}: {e}")
                            pass
                    
                    if matched:
                        found_count += 1
                        
                        # 获取文件信息
                        try:
                            stat_info = os.stat(file_path)
                            mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                            size_bytes = stat_info.st_size
                            # 格式化大小
                            if size_bytes < 1024:
                                size_str = f"{size_bytes} B"
                            elif size_bytes < 1024 * 1024:
                                size_str = f"{size_bytes / 1024:.1f} KB"
                            elif size_bytes < 1024 * 1024 * 1024:
                                size_str = f"{size_bytes / (1024 * 1024):.1f} MB"
                            else:
                                size_str = f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
                        except:
                            mtime = "-"
                            size_str = "-"
                        
                        # 获取文件名和扩展名
                        name_without_ext, file_ext = os.path.splitext(filename)
                        # 获取不带扩展名的完整路径
                        path_without_ext = os.path.join(root, name_without_ext)
                        if file_ext:
                            file_type = file_ext[1:].upper()  # 去掉点并转大写
                        else:
                            file_type = "无"
                        
                        results_buffer.append({
                            'path': file_path,
                            'name': f"{match_type} {path_without_ext}",
                            'full_path': f"{match_type} {file_path}",
                            'file_type': file_type,
                            'date': mtime,
                            'size': size_str
                        })
                        
                        # 批量更新UI（每20个结果更新一次）
                        if len(results_buffer) >= buffer_size:
                            # 将结果放入队列
                            for item in results_buffer:
                                self.result_queue.put({'type': 'result', **item})
                            results_buffer.clear()
        except Exception as e:
            print(f"Search error: {e}")
        
        # 添加剩余的结果
        if results_buffer:
            for item in results_buffer:
                self.result_queue.put({'type': 'result', **item})
        
        # 调试信息
        print(f"[Search] 搜索完成，共扫描 {scanned_files} 个文件，找到 {found_count} 个结果")
        
        # 重置搜索状态（先重置，避免后续更新被跳过）
        self.is_searching = False
        
        # 搜索完成，更新UI状态（通过队列）
        final_status = f"搜索完成，共找到 {found_count} 个结果（扫描了 {scanned_files} 个文件）"
        self.result_queue.put({'type': 'status', 'text': final_status})
        self.result_queue.put({'type': 'button', 'button': 'search', 'enabled': True})
        self.result_queue.put({'type': 'button', 'button': 'stop', 'enabled': False})
        
        print(f"[Search] UI更新已调度（使用队列）")
    
    def on_result_double_clicked(self, row, column):
        """双击搜索结果，打开文件所在文件夹或文件夹本身"""
        # 从第一列获取存储的完整路径
        path_item = self.result_list.item(row, 0)
        if path_item:
            file_path = path_item.data(256)  # 获取存储的完整路径
            
            if os.path.exists(file_path):
                # 如果是文件夹，直接打开文件夹；如果是文件，打开文件所在文件夹
                if os.path.isdir(file_path):
                    folder_path = file_path
                else:
                    folder_path = os.path.dirname(file_path)
                # 不关闭搜索对话框，保持独立
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    self.main_window.add_new_tab(folder_path)

import sys
import os
import json
import subprocess
import string
import time
import socket
import threading
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QListWidget, QLabel, QToolBar, QAction, QMenu, QMessageBox, QInputDialog, QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView)  # QDockWidget removed (unused)
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import Qt, QDir, QUrl, pyqtSignal, pyqtSlot, Q_ARG, QObject  # QModelIndex removed (unused)
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QMouseEvent
# from PyQt5.QtGui import QIcon  # unused


# Optional native hit-test support (Windows)
try:
    import ctypes
    import win32gui
    import win32con
    HAS_PYWIN = True
except Exception:
    HAS_PYWIN = False

# 面包屑导航路径栏
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
            
            # CLSID 路径映射（防止未转换的CLSID路径）
            clsid_map = {
                '::{20D04FE0-3AEA-1069-A2D8-08002B30309D}': '此电脑',
                '::{645FF040-5081-101B-9F08-00AA002F954E}': '回收站',
                '::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}': '网络',
            }
            
            path = self.current_path
            display = None
            
            # 先检查 shell: 路径
            display = shell_map.get(path, None)
            
            # 检查 CLSID 路径
            if not display and '::' in path:
                for clsid, name in clsid_map.items():
                    if clsid in path:
                        display = name
                        break
            
            # 检查是否是其他 shell: 路径
            if not display and path.startswith('shell:'):
                display = path  # 兜底显示原始shell:路径
            
            # 最后使用路径本身
            if not display:
                display = path
                
            # 统一对所有display做长度限制
            is_pinned = getattr(self, 'is_pinned', False)
            max_len = 12 if is_pinned else 16  # 固定标签页显示更短，为📌图标留空间
            if len(display) > max_len:
                display = display[-max_len:]
            pin_prefix = "📌" if is_pinned else ""
            title = pin_prefix + display
            print(f"DEBUG update_tab_title: path={path}, is_pinned={is_pinned}, pin_prefix='{pin_prefix}', title='{title}'")
            if self.main_window and hasattr(self.main_window, 'tab_widget'):
                idx = self.main_window.tab_widget.indexOf(self)
                if idx != -1:
                    self.main_window.tab_widget.setTabText(idx, title)
                    print(f"DEBUG: Set tab {idx} text to '{title}'")
    
    def refresh_explorer(self):
        """刷新Explorer视图以更新覆盖图标（如Git Tortoise图标）"""
        try:
            # 方法1: 发送F5键刷新
            self.explorer.dynamicCall('Refresh()')
        except:
            try:
                # 方法2: 重新导航到当前路径
                if hasattr(self, 'current_path') and self.current_path:
                    is_shell = self.current_path.startswith('shell:')
                    self.navigate_to(self.current_path, is_shell=is_shell)
            except Exception as e:
                print(f"[refresh_explorer] 刷新失败: {e}")

    def start_path_sync_timer(self):
        from PyQt5.QtCore import QTimer
        self._path_sync_timer = QTimer(self)
        self._path_sync_timer.timeout.connect(self.sync_path_bar_with_explorer)
        self._path_sync_timer.start(500)

    def sync_path_bar_with_explorer(self):
        """同步Explorer窗口的当前路径到current_path"""
        # 如果是真实嵌入的Explorer窗口
        if hasattr(self, 'explorer_hwnd'):
            try:
                import win32gui
                
                # 获取窗口标题，通常包含当前路径
                title = win32gui.GetWindowText(self.explorer_hwnd)
                
                # Explorer窗口标题就是当前文件夹路径
                if title and title != self.current_path:
                    # 验证是否是有效路径
                    import os
                    if os.path.exists(title):
                        print(f"[sync_path] 路径变化: {self.current_path} -> {title}")
                        self.current_path = title
                        self.update_tab_title()
                        # 同步左侧目录树
                        if self.main_window and hasattr(self.main_window, 'expand_dir_tree_to_path'):
                            self.main_window.expand_dir_tree_to_path(title)
            except Exception as e:
                pass
            return
        
        # 如果是QAxWidget的Shell.Explorer控件（降级方案）
        if not hasattr(self, 'explorer'):
            return
        
        try:
            url = self.explorer.property('LocationURL')
            if url:
                url_str = str(url)
                local_path = None
                
                # 处理 file:/// 本地路径
                if url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = unquote(url_str[8:])
                    if os.name == 'nt' and local_path.startswith('/'):
                        local_path = local_path[1:]
                # 处理 shell: 特殊路径
                elif url_str.startswith('shell:') or '::' in url_str:
                    # Shell特殊文件夹，通常以 shell: 或包含 CLSID (::)
                    # 这些路径我们已经在 current_path 中维护，无需更新
                    return
                
                if local_path and local_path != self.current_path:
                    self.current_path = local_path
                    if hasattr(self, 'path_bar'):
                        self.path_bar.set_path(local_path)
                    self.update_tab_title()
                    # 同步左侧目录树
                    if self.main_window and hasattr(self.main_window, 'expand_dir_tree_to_path'):
                        self.main_window.expand_dir_tree_to_path(local_path)
        except Exception:
            pass

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 嵌入真正的Windows资源管理器窗口（完整功能）
        try:
            self.embed_real_explorer()
        except Exception as e:
            print(f"[setup_ui] 嵌入真实Explorer失败: {e}")
            # 降级到使用简单控件
            self.use_simple_explorer()
    
    def embed_real_explorer(self):
        """嵌入真正的Windows资源管理器窗口"""
        import subprocess
        import win32gui
        import win32con
        import win32process
        from PyQt5.QtCore import QTimer
        from PyQt5.QtWidgets import QWidget
        import time
        
        # 创建容器widget
        self.explorer_container = QWidget()
        self.layout().addWidget(self.explorer_container)
        
        # 如果传入了已存在的Explorer窗口句柄，直接嵌入它
        if hasattr(self, 'existing_explorer_hwnd') and self.existing_explorer_hwnd:
            print(f"[embed_real_explorer] 嵌入已存在的Explorer窗口: {self.existing_explorer_hwnd}")
            self._embed_window(self.existing_explorer_hwnd)
            return
        
        # 否则，启动新的Explorer进程
        # 记录启动前的所有Explorer窗口
        existing_windows = set()
        def record_existing(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                class_name = win32gui.GetClassName(hwnd)
                if class_name in ['CabinetWClass', 'ExploreWClass']:
                    existing_windows.add(hwnd)
            return True
        
        win32gui.EnumWindows(record_existing, None)
        
        # 启动explorer进程
        path = QDir.toNativeSeparators(self.current_path)
        # 使用 /e 参数打开资源管理器（带左侧树），/root 设置根目录
        cmd = f'explorer.exe /e,/root,"{path}"'
        
        print(f"[embed_real_explorer] 启动命令: {cmd}")
        subprocess.Popen(cmd, shell=True)
        
        # 等待窗口创建并获取句柄
        self.retry_count = 0
        def find_and_embed():
            try:
                self.retry_count += 1
                
                # 查找新创建的Explorer窗口（不在existing_windows中的）
                new_windows = []
                
                def callback(hwnd, extra):
                    if win32gui.IsWindowVisible(hwnd) and hwnd not in existing_windows:
                        class_name = win32gui.GetClassName(hwnd)
                        if class_name in ['CabinetWClass', 'ExploreWClass']:
                            new_windows.append(hwnd)
                    return True
                
                win32gui.EnumWindows(callback, None)
                
                explorer_hwnd = new_windows[0] if new_windows else None
                
                if explorer_hwnd:
                    print(f"[embed_real_explorer] 找到Explorer窗口: {explorer_hwnd}")
                    self._embed_window(explorer_hwnd)
                else:
                    # 最多重试15次（3秒）
                    if self.retry_count < 15:
                        print(f"[embed_real_explorer] 未找到Explorer窗口，重试 {self.retry_count}/15...")
                        QTimer.singleShot(200, find_and_embed)
                    else:
                        print(f"[embed_real_explorer] 超时未找到窗口，降级到简单控件")
                        self.use_simple_explorer()
                    
            except Exception as e:
                print(f"[embed_real_explorer] 嵌入失败: {e}")
                import traceback
                traceback.print_exc()
        
        # 增加初始延迟到800ms，给Explorer更多启动时间
        QTimer.singleShot(800, find_and_embed)
    
    def _embed_window(self, explorer_hwnd):
        """嵌入Explorer窗口到容器中"""
        import win32gui
        import win32con
        
        # 将Explorer窗口设置为子窗口并嵌入
        container_hwnd = int(self.explorer_container.winId())
        
        # 先隐藏Explorer窗口，避免嵌入过程中的闪烁和抖动
        win32gui.ShowWindow(explorer_hwnd, win32con.SW_HIDE)
        
        # 移除Explorer窗口的边框和标题栏
        style = win32gui.GetWindowLong(explorer_hwnd, win32con.GWL_STYLE)
        style = style & ~win32con.WS_CAPTION & ~win32con.WS_THICKFRAME & ~win32con.WS_BORDER
        win32gui.SetWindowLong(explorer_hwnd, win32con.GWL_STYLE, style)
        
        # 移除扩展样式中的边框
        ex_style = win32gui.GetWindowLong(explorer_hwnd, win32con.GWL_EXSTYLE)
        ex_style = ex_style & ~win32con.WS_EX_CLIENTEDGE & ~win32con.WS_EX_WINDOWEDGE
        win32gui.SetWindowLong(explorer_hwnd, win32con.GWL_EXSTYLE, ex_style)
        
        # 设置为子窗口（WS_CHILD样式）
        win32gui.SetParent(explorer_hwnd, container_hwnd)
        
        # 调整大小以充满容器，使用SetWindowPos确保位置正确
        rect = self.explorer_container.rect()
        win32gui.SetWindowPos(
            explorer_hwnd,
            0,  # HWND_TOP
            0, 0,  # x, y 坐标
            rect.width(), rect.height(),
            win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_FRAMECHANGED
        )
        
        # 最后显示窗口
        win32gui.ShowWindow(explorer_hwnd, win32con.SW_SHOW)
        
        # 强制刷新窗口以确保正确显示
        win32gui.InvalidateRect(explorer_hwnd, None, True)
        win32gui.UpdateWindow(explorer_hwnd)
        
        # 保存句柄以便后续使用
        self.explorer_hwnd = explorer_hwnd
        
        # 监听容器大小变化
        self.explorer_container.resizeEvent = self.on_container_resize
        
        print(f"[_embed_window] Explorer窗口已嵌入")
    
    def on_container_resize(self, event):
        """容器大小改变时调整嵌入的Explorer窗口大小"""
        if hasattr(self, 'explorer_hwnd'):
            try:
                import win32gui
                import win32con
                rect = self.explorer_container.rect()
                # 强制设置窗口位置和大小，确保始终从(0,0)开始
                win32gui.SetWindowPos(
                    self.explorer_hwnd,
                    win32con.HWND_TOP,
                    0, 0,  # x, y - 确保从容器左上角开始
                    rect.width(), rect.height(),
                    win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED
                )
                # 强制刷新窗口
                win32gui.InvalidateRect(self.explorer_hwnd, None, True)
                win32gui.UpdateWindow(self.explorer_hwnd)
            except Exception as e:
                print(f"[on_container_resize] 调整窗口大小失败: {e}")
    
    def use_simple_explorer(self):
        """降级方案：使用简单的Shell.Explorer控件"""
        self.explorer = QAxWidget(self)
        self.explorer.setControl("Shell.Explorer.2")
        self.layout().addWidget(self.explorer)
        self.explorer.dynamicCall('Navigate(const QString&)', QDir.toNativeSeparators(self.current_path))
        print(f"[setup_ui] 使用简单Explorer控件")
        self.explorer.dynamicCall('OnNavigateComplete2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnDocumentComplete(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnBeforeNavigate2(QVariant,QVariant,QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None, None, None)
        self.explorer.dynamicCall('OnNewWindow2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('OnNewWindow3(QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None)

        # 注意：真实Explorer窗口的事件处理由Windows本身管理
        # 我们不需要installEventFilter，因为不再使用QAxWidget

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
                            self.path_bar.set_path(local_path)
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

    def on_path_bar_changed(self, path):
        """处理面包屑路径栏的路径变化"""
        path = path.strip()
        # 处理cmd命令
        if path.lower() == 'cmd':
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    subprocess.Popen(['cmd', '/K', 'cd', '/d', current_dir], creationflags=subprocess.CREATE_NEW_CONSOLE)
                    # 恢复路径栏显示当前路径
                    self.path_bar.set_path(current_dir)
                else:
                    QMessageBox.warning(self, "错误", "当前路径无效，无法打开命令行")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"无法打开命令行: {e}")
            return
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
        # 返回上一级目录，盘符根目录时导航到"此电脑"
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
        # 判断是否为盘符根目录，导航到"此电脑"
        if path.endswith(":\\") or path.endswith(":/"):
            self.navigate_to('shell:MyComputerFolder', is_shell=True)
            return
        parent_path = os.path.dirname(path)
        if parent_path and os.path.exists(parent_path):
            self.navigate_to(parent_path)

    def __init__(self, parent=None, path="", is_shell=False, existing_hwnd=None):
        super().__init__(parent)
        self.main_window = parent
        self.current_path = path if path else QDir.homePath()
        # 如果传入了已存在的Explorer窗口句柄，保存它
        self.existing_explorer_hwnd = existing_hwnd
        self.setup_ui()
        if not existing_hwnd:  # 只有新创建的标签页需要导航
            self.navigate_to(self.current_path, is_shell=is_shell)
        self.start_path_sync_timer()

    # 移除重复的setup_ui，保留带路径栏的实现

    def navigate_to(self, path, is_shell=False):
        # 支持本地路径和shell特殊路径
        
        # 自动检测 CLSID 路径并转换
        if '::' in path and not is_shell:
            clsid_map = {
                '::{20D04FE0-3AEA-1069-A2D8-08002B30309D}': 'shell:MyComputerFolder',
                '::{645FF040-5081-101B-9F08-00AA002F954E}': 'shell:RecycleBinFolder',
                '::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}': 'shell:NetworkPlacesFolder',
            }
            for clsid, shell_path in clsid_map.items():
                if clsid in path:
                    print(f"[navigate_to] 检测到CLSID路径，转换为: {shell_path}")
                    path = shell_path
                    is_shell = True
                    break
        
        # 如果使用真实Explorer窗口，需要通过COM导航
        if hasattr(self, 'explorer_hwnd'):
            try:
                import win32com.client
                shell = win32com.client.Dispatch("Shell.Application")
                # 查找我们的Explorer窗口
                for window in shell.Windows():
                    try:
                        if int(window.HWND) == self.explorer_hwnd:
                            # 导航到新路径
                            if is_shell or path.startswith('shell:'):
                                window.Navigate(path)
                            else:
                                window.Navigate(QDir.toNativeSeparators(path))
                            self.current_path = path
                            self.update_tab_title()
                            return
                    except:
                        continue
            except Exception as e:
                print(f"[navigate_to] COM导航失败: {e}")
        
        # 降级到使用QAxWidget
        if hasattr(self, 'explorer') and self.explorer:
            if is_shell or path.startswith('shell:'):
                self.explorer.dynamicCall("Navigate(const QString&)", path)
                self.current_path = path
                if hasattr(self, 'path_bar'):
                    self.path_bar.set_path(path)
                self.update_tab_title()
            elif os.path.exists(path):
                url = QDir.toNativeSeparators(path)
                self.explorer.dynamicCall("Navigate(const QString&)", url)
                self.current_path = path
                if hasattr(self, 'path_bar'):
                    self.path_bar.set_path(path)
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
    
    def mouseDoubleClickEvent(self, event):
        """捕获 TabWidget 区域的双击事件"""
        from PyQt5.QtCore import QPoint
        # 获取 TabBar 的几何位置
        tabbar = self.tabBar()
        # 将事件位置转换为 TabBar 的坐标系
        tabbar_pos = tabbar.mapFrom(self, event.pos())
        
        print(f"[DEBUG] TabWidget double click: pos={event.pos()}, tabbar_pos={tabbar_pos}")
        print(f"[DEBUG] TabBar rect: {tabbar.rect()}")
        
        # 检查点击是否在 TabBar 的矩形范围内（使用 TabBar 自己的坐标系）
        in_tabbar = tabbar.rect().contains(tabbar_pos)
        print(f"[DEBUG] In TabBar: {in_tabbar}")
        
        if in_tabbar:
            # 在 TabBar 内，检查是否点击在空白区域
            clicked_tab = tabbar.tabAt(tabbar_pos)
            print(f"[DEBUG] Clicked tab index: {clicked_tab}")
            
            if clicked_tab == -1:
                # 空白区域，打开新标签页
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    print(f"[DEBUG] Opening new tab from TabBar blank area...")
                    self.main_window.add_new_tab()
                    return
        else:
            # 不在 TabBar 内，检查是否在标签页头部区域（TabBar 右侧的空白）
            # 获取 TabWidget 的 TabBar 所在的区域高度
            if event.pos().y() < tabbar.height():
                print(f"[DEBUG] Click is in tab header area but outside TabBar")
                # 这是标签头和按钮之间的空白区域，打开新标签页
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    print(f"[DEBUG] Opening new tab from header blank area...")
                    self.main_window.add_new_tab()
                    return
        
        super().mouseDoubleClickEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
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


# 自定义 TabBar 以支持双击空白区域打开新标签页和悬停显示关闭按钮
from PyQt5.QtWidgets import QTabBar, QToolButton
from PyQt5.QtCore import QEvent, QPoint
from PyQt5.QtGui import QIcon
class CustomTabBar(QTabBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = None
        self.hovered_tab = -1  # 当前鼠标悬停的标签页索引
        self.setMouseTracking(True)  # 启用鼠标追踪
        self.setMovable(True)  # 启用标签页拖拽排序
        # 连接标签移动信号
        self.tabMoved.connect(self.on_tab_moved)
    
    def event(self, event):
        # 拦截所有事件，确保双击事件能被处理
        if event.type() == QEvent.MouseButtonDblClick:
            print(f"[DEBUG] TabBar event: MouseButtonDblClick")
            self.mouseDoubleClickEvent(event)
            return True
        return super().event(event)
    
    def mouseDoubleClickEvent(self, event):
        print(f"[DEBUG] TabBar double click event triggered")
        # 获取点击位置
        pos = event.pos()
        # 判断是否点在空白区域（没有点在任何标签页上）
        clicked_tab = self.tabAt(pos)
        print(f"[DEBUG] Clicked tab: {clicked_tab}, pos: ({pos.x()}, {pos.y()}), count: {self.count()}")
        
        # 如果点击在空白区域，或点击在最后一个标签右侧的空白处
        is_blank = clicked_tab == -1
        if not is_blank and self.count() > 0:
            # 检查是否点击在最后一个标签页的右侧
            last_tab_rect = self.tabRect(self.count() - 1)
            print(f"[DEBUG] Last tab right edge: {last_tab_rect.right()}")
            if pos.x() > last_tab_rect.right():
                is_blank = True
        
        print(f"[DEBUG] Is blank area: {is_blank}, has main_window: {self.main_window is not None}")
        
        if is_blank:
            # 点击在空白区域，打开新标签页
            if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                print(f"[DEBUG] Opening new tab from TabBar...")
                self.main_window.add_new_tab()
                event.accept()
                return
        
        # 如果点击在标签页上，调用默认行为
        super().mouseDoubleClickEvent(event)
    
    def mouseMoveEvent(self, event):
        """追踪鼠标位置，更新悬停的标签页"""
        tab_index = self.tabAt(event.pos())
        if tab_index != self.hovered_tab:
            # 移除旧的关闭按钮
            if self.hovered_tab >= 0:
                self.setTabButton(self.hovered_tab, QTabBar.RightSide, None)
            
            # 添加新的关闭按钮
            self.hovered_tab = tab_index
            if self.hovered_tab >= 0:
                close_btn = QToolButton(self)
                close_btn.setText("×")
                close_btn.setFixedSize(16, 16)
                close_btn.setStyleSheet("""
                    QToolButton {
                        border: none;
                        background: transparent;
                        color: #666;
                        font-size: 18px;
                        font-weight: bold;
                        padding: 0px;
                        margin: 0px;
                    }
                    QToolButton:hover {
                        background: #ff6b6b;
                        color: white;
                        border-radius: 8px;
                    }
                """)
                close_btn.clicked.connect(lambda: self.close_tab_at_index(self.hovered_tab))
                self.setTabButton(self.hovered_tab, QTabBar.RightSide, close_btn)
        
        super().mouseMoveEvent(event)
    
    def leaveEvent(self, event):
        """鼠标离开标签栏时移除关闭按钮"""
        if self.hovered_tab >= 0:
            self.setTabButton(self.hovered_tab, QTabBar.RightSide, None)
            self.hovered_tab = -1
        super().leaveEvent(event)
    
    def close_tab_at_index(self, index):
        """关闭指定索引的标签页"""
        if self.main_window and hasattr(self.main_window, 'close_tab'):
            self.main_window.close_tab(index)
    
    def on_tab_moved(self, from_index, to_index):
        """标签页移动后的处理，确保固定标签页始终在左侧"""
        if not self.main_window:
            return
        
        # 获取被移动的标签页
        moved_tab = self.main_window.tab_widget.widget(to_index)
        if not moved_tab:
            return
        
        # 检查是否违反固定标签页规则
        is_pinned = getattr(moved_tab, 'is_pinned', False)
        
        # 统计固定标签页的数量
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


class MainWindow(QMainWindow):
    # 定义信号用于从服务器线程通知主线程打开新标签
    open_path_signal = pyqtSignal(str)
    
    def ensure_default_icons_on_bookmark_bar(self):
        """确保五个常用书签（带图标）始终在最左侧且不会被覆盖。"""
        bm = self.bookmark_manager
        tree = bm.get_tree()
        bar = tree.get('bookmark_bar')
        if not bar or 'children' not in bar:
            return
        import time
        import os
        from PyQt5.QtCore import QStandardPaths
        downloads_path = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
        if not downloads_path or not os.path.exists(downloads_path):
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
            self.add_new_tab()


    def go_up_current_tab(self):
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'go_up'):
            current_tab.go_up(force=True)
    
    @pyqtSlot()
    @pyqtSlot(str)
    @pyqtSlot(str, bool)
    def add_new_tab(self, path="", is_shell=False):
        # 默认新建标签页为“此电脑”
        if not path:
            path = 'shell:MyComputerFolder'
            is_shell = True
        tab = FileExplorerTab(self, path, is_shell=is_shell)
        tab.is_pinned = False
        short = path[-16:] if len(path) > 16 else path
        tab_index = self.tab_widget.addTab(tab, short)
        self.tab_widget.setCurrentIndex(tab_index)
        
        # 激活窗口（当从其他实例接收到路径时）
        self.activateWindow()
        self.raise_()
        
        return tab_index

    @pyqtSlot(int, str)
    def embed_existing_explorer(self, hwnd, path):
        """嵌入已存在的Explorer窗口到新标签页"""
        print(f"[MainWindow] 嵌入已存在Explorer窗口: hwnd={hwnd}, path={path}")
        
        # 创建新标签页，但使用特殊标记告诉它嵌入已存在的窗口
        is_shell = path.startswith('shell:') or '::' in path
        tab = FileExplorerTab(self, path, is_shell=is_shell, existing_hwnd=hwnd)
        tab.is_pinned = False
        short = path[-16:] if len(path) > 16 else path
        tab_index = self.tab_widget.addTab(tab, short)
        self.tab_widget.setCurrentIndex(tab_index)
        
        # 激活窗口
        self.activateWindow()
        self.raise_()
        
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
                # 刷新Explorer视图以更新覆盖图标
                if hasattr(tab, 'refresh_explorer'):
                    from PyQt5.QtCore import QTimer
                    # 延迟刷新，确保视图已完全切换
                    QTimer.singleShot(100, tab.refresh_explorer)
                # 触发嵌入窗口的resize以适应容器大小
                if hasattr(tab, 'explorer_container') and hasattr(tab, 'explorer_hwnd'):
                    from PyQt5.QtCore import QTimer
                    # 立即调整一次
                    self._resize_embedded_explorer(tab)
                    # 延迟再次调整，确保完全显示
                    QTimer.singleShot(100, lambda: self._resize_embedded_explorer(tab))
    
    def _resize_embedded_explorer(self, tab):
        """调整标签页中嵌入的Explorer窗口大小"""
        try:
            import win32gui
            import win32con
            if hasattr(tab, 'explorer_hwnd') and hasattr(tab, 'explorer_container'):
                rect = tab.explorer_container.rect()
                # 确保窗口可见
                win32gui.ShowWindow(tab.explorer_hwnd, win32con.SW_SHOW)
                # 使用SetWindowPos确保位置从(0,0)开始，并使用FRAMECHANGED强制重绘
                win32gui.SetWindowPos(
                    tab.explorer_hwnd,
                    win32con.HWND_TOP,
                    0, 0,
                    rect.width(), rect.height(),
                    win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED
                )
                # 强制刷新窗口内容
                win32gui.InvalidateRect(tab.explorer_hwnd, None, True)
                win32gui.UpdateWindow(tab.explorer_hwnd)
        except Exception as e:
            pass  # 静默处理错误

    def expand_dir_tree_to_path(self, path):
        # 展开并选中左侧目录树到指定路径（自定义模型）
        if not hasattr(self, 'custom_tree_model') or not hasattr(self, 'dir_tree'):
            return
        if not path or not os.path.exists(path):
            return
        # 如果是网络路径，直接返回，不展开目录树
        if path.startswith('\\\\'):
            return
        
        # 规范化路径
        path = os.path.normpath(path)
        
        # 获取路径的驱动器号
        drive = os.path.splitdrive(path)[0]  # 如 "D:"
        if not drive:
            return
        
        # 在树中查找对应的驱动器节点
        drive_item = None
        for i in range(self.custom_tree_model.rowCount()):
            item = self.custom_tree_model.item(i)
            if hasattr(item, 'fs_path') and item.fs_path and item.fs_path.startswith(drive):
                drive_item = item
                break
        
        if not drive_item:
            return
        
        # 展开驱动器节点
        drive_index = self.custom_tree_model.indexFromItem(drive_item)
        self.dir_tree.expand(drive_index)
        
        # 确保驱动器子目录已加载
        if drive_item.rowCount() == 0 or drive_item.child(0).text() == "":
            self.on_custom_tree_expanded(drive_index)
        
        # 分解路径为各级目录
        rel_path = os.path.splitdrive(path)[1]  # 去掉驱动器号，如 "\project\test"
        if rel_path.startswith('\\') or rel_path.startswith('/'):
            rel_path = rel_path[1:]
        
        if not rel_path:  # 如果是驱动器根目录
            self.dir_tree.setCurrentIndex(drive_index)
            self.dir_tree.scrollTo(drive_index)
            return
        
        # 逐级查找并展开路径
        path_parts = rel_path.split(os.sep)
        current_item = drive_item
        
        for part in path_parts:
            if not part:
                continue
            
            # 在当前节点的子节点中查找
            found = False
            for i in range(current_item.rowCount()):
                child = current_item.child(i)
                if child.text() == part:
                    # 找到匹配的子节点
                    child_index = self.custom_tree_model.indexFromItem(child)
                    self.dir_tree.expand(child_index)
                    
                    # 确保子目录已加载
                    if child.rowCount() == 0 or child.child(0).text() == "":
                        self.on_custom_tree_expanded(child_index)
                    
                    current_item = child
                    found = True
                    break
            
            if not found:
                # 如果找不到，可能是因为目录还未加载
                break
        
        # 选中并滚动到最后找到的节点
        final_index = self.custom_tree_model.indexFromItem(current_item)
        self.dir_tree.setCurrentIndex(final_index)
        self.dir_tree.scrollTo(final_index)
    

    

    
    def create_toolbar(self):
        toolbar = QToolBar()
        self.addToolBar(Qt.TopToolBarArea, toolbar)
        # 工具栏保留，可在此添加其它功能按钮
        pass

    def show_bookmark_dialog(self):
        dlg = BookmarkDialog(self.bookmark_manager, self)
        dlg.exec_()
    
    def show_search_dialog(self):
        """显示搜索对话框（非模态）"""
        current_tab = self.tab_widget.currentWidget()
        if not current_tab or not hasattr(current_tab, 'current_path'):
            QMessageBox.warning(self, "提示", "请先打开一个文件夹")
            return
        
        search_path = current_tab.current_path
        
        # 不支持搜索特殊路径
        if search_path.startswith('shell:'):
            QMessageBox.warning(self, "提示", "不支持搜索特殊路径（shell:）")
            return
        
        if not os.path.exists(search_path):
            QMessageBox.warning(self, "提示", f"路径不存在: {search_path}")
            return
        
        # 创建非模态对话框
        dlg = SearchDialog(search_path, self)
        # 保存对话框引用，防止被垃圾回收
        if not hasattr(self, 'search_dialogs'):
            self.search_dialogs = []
        self.search_dialogs.append(dlg)
        
        # 对话框关闭时从列表中移除
        dlg.finished.connect(lambda: self.search_dialogs.remove(dlg) if dlg in self.search_dialogs else None)
        
        # 非模态显示，不阻塞主窗口
        dlg.show()

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
            # 先添加标签页（临时标题）
            self.tab_widget.addTab(tab, "")
            # 然后调用update_tab_title更新标题（会考虑shell路径映射和图标）
            tab.update_tab_title()
        # 恢复原先的tab焦点
        if current_tab is not None:
            for i, tab in enumerate(new_tabs):
                if tab is current_tab:
                    self.tab_widget.setCurrentIndex(i)
                    break

    def load_config(self):
        """加载配置文件"""
        config_file = "config.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                print(f"[Config] 加载配置失败: {e}")
                return {}
        return {}
    
    def save_config(self, config):
        """保存配置文件"""
        config_file = "config.json"
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            print(f"[Config] 配置已保存")
        except Exception as e:
            print(f"[Config] 保存配置失败: {e}")
    
    def save_pinned_tabs(self):
        pinned_paths = []
        for i in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(i)
            if hasattr(tab, 'is_pinned') and tab.is_pinned:
                if hasattr(tab, 'current_path'):
                    pinned_paths.append(tab.current_path)
        
        # 保存到统一配置文件
        config = self.load_config()
        config['pinned_tabs'] = pinned_paths
        self.save_config(config)

    def load_pinned_tabs(self):
        has_pinned = False
        config = self.load_config()
        paths = config.get('pinned_tabs', [])
        if paths:
            try:
                pass  # paths已经从config获取
                for path in paths:
                    tab = FileExplorerTab(self, path)
                    tab.is_pinned = True
                    short = path[-12:] if len(path) > 12 else path
                    pin_prefix = "📌"
                    title = pin_prefix + short
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
        
        # 初始化资源管理器监控（默认不启动）
        self.explorer_monitor = None
        self._init_explorer_monitor()

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

        # 隐藏左侧目录树，因为右侧Explorer会显示自己的导航窗格
        # 如果需要恢复，取消以下注释：
        """
        # 左侧目录树 - 使用自定义模型显示固定书签和盘符
        self.dir_tree = QTreeView()
        self.dir_tree.setHeaderHidden(True)
        self.dir_tree.setMinimumWidth(200)
        self.dir_tree.setMaximumWidth(350)
        self.dir_tree.clicked.connect(self.on_custom_tree_clicked)
        self.populate_custom_dir_tree()
        self.splitter.addWidget(self.dir_tree)
        """

        # 右侧原有标签页区域
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # 创建标签页控件（支持拖放）
        self.tab_widget = DragDropTabWidget(self)
        self.tab_widget.setTabsClosable(False)  # 禁用默认关闭按钮，使用自定义悬停关闭按钮
        self.tab_widget.currentChanged.connect(self.on_tab_changed)

        # 使用自定义 TabBar 支持双击空白区域打开新标签页
        custom_tabbar = CustomTabBar()
        custom_tabbar.main_window = self
        self.tab_widget.setTabBar(custom_tabbar)

        # 设置选中标签页背景色为淡黄色
        tabbar = self.tab_widget.tabBar()
        tabbar.setAcceptDrops(True)
        tabbar.setStyleSheet("""
            QTabBar::tab {
                border-right: 1px solid #d3d3d3;
                padding-right: 12px;
                padding-left: 12px;
                min-height: 36px;
                min-width: 120px;
                font-size: 15px;
            }
            QTabBar::tab:selected {
                background: #FFF9CC;
            }
        """)

        # 添加导航和新标签页按钮
        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        
        # 新建标签页按钮
        self.add_tab_button = QPushButton("➕")
        self.add_tab_button.setToolTip("新建标签页")
        self.add_tab_button.setFixedHeight(35)
        self.add_tab_button.setFixedWidth(35)
        self.add_tab_button.clicked.connect(self.add_new_tab)
        btn_layout.addWidget(self.add_tab_button)
        
        # 搜索按钮
        self.search_button = QPushButton("🔍")
        self.search_button.setToolTip("搜索当前文件夹")
        self.search_button.setFixedHeight(35)
        self.search_button.setFixedWidth(35)
        self.search_button.clicked.connect(self.show_search_dialog)
        btn_layout.addWidget(self.search_button)
        
        btn_layout.addStretch(1)
        self.tab_widget.setCornerWidget(btn_widget)
        
        # 为 btn_widget 添加双击事件处理，双击空白区域打开新标签页
        def btn_widget_double_click(event):
            print(f"[DEBUG] btn_widget double click event triggered")
            # 检查点击位置是否在按钮之外的空白区域
            from PyQt5.QtWidgets import QApplication
            child = btn_widget.childAt(event.pos())
            print(f"[DEBUG] Clicked child widget: {child}")
            if child is None:
                # 点击在空白区域
                print(f"[DEBUG] Opening new tab from btn_widget blank area")
                self.add_new_tab()
        
        btn_widget.mouseDoubleClickEvent = btn_widget_double_click

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
        
        # 连接信号
        self.open_path_signal.connect(self.handle_open_path_from_instance)
        
        # 添加F5快捷键刷新当前Explorer视图
        from PyQt5.QtWidgets import QShortcut
        from PyQt5.QtGui import QKeySequence
        refresh_shortcut = QShortcut(QKeySequence("F5"), self)
        refresh_shortcut.activated.connect(self.refresh_current_tab)
        
        # 启动单实例通信服务器
        self.start_instance_server()
    
    def refresh_current_tab(self):
        """刷新当前标签页的Explorer视图"""
        current_tab = self.tab_widget.currentWidget()
        if current_tab and hasattr(current_tab, 'refresh_explorer'):
            current_tab.refresh_explorer()
            print("[MainWindow] 已刷新当前Explorer视图")

    def handle_open_path_from_instance(self, path):
        """处理从其他实例接收到的路径（在主线程中）"""
        print(f"[MainWindow] Opening path in new tab: {path}")
        self.add_new_tab(path)
        # 激活并置顶窗口
        self.activateWindow()
        self.raise_()
        self.showNormal()  # 如果最小化则恢复
    
    def start_instance_server(self):
        """启动本地服务器监听其他实例的请求"""
        def server_thread():
            try:
                server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                server.bind(('127.0.0.1', 58923))  # 使用固定端口
                server.listen(5)
                server.settimeout(1.0)  # 设置超时，使线程可以退出
                self.server_socket = server
                print("[Server] Instance server started on port 58923")
                
                while getattr(self, 'server_running', True):
                    try:
                        conn, addr = server.accept()
                        data = conn.recv(4096).decode('utf-8')
                        conn.close()
                        
                        if data:
                            print(f"[Server] Received path: {data}")
                            # 使用信号在主线程中打开新标签页
                            self.open_path_signal.emit(data)
                    except socket.timeout:
                        continue
                    except Exception as e:
                        print(f"[Server] Connection error: {e}")
                        continue
            except Exception as e:
                print(f"[Server] Failed to start server: {e}")
        
        self.server_running = True
        server_thread_obj = threading.Thread(target=server_thread, daemon=True)
        server_thread_obj.start()
        # 等待服务器启动
        time.sleep(0.2)

    def closeEvent(self, event):
        """窗口关闭时停止服务器和监控"""
        self.server_running = False
        if hasattr(self, 'server_socket'):
            try:
                self.server_socket.close()
            except:
                pass
        
        # 停止资源管理器监控
        if self.explorer_monitor:
            self.explorer_monitor.stop()
        
        super().closeEvent(event)

    def resizeEvent(self, event):
        """窗口大小改变时，调整所有嵌入的Explorer窗口大小"""
        super().resizeEvent(event)
        # 遍历所有标签页，调整嵌入窗口大小
        for i in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(i)
            if hasattr(tab, 'explorer_hwnd') and hasattr(tab, 'explorer_container'):
                try:
                    import win32gui
                    import win32con
                    rect = tab.explorer_container.rect()
                    # 使用SetWindowPos确保位置和大小都正确
                    win32gui.SetWindowPos(
                        tab.explorer_hwnd,
                        win32con.HWND_TOP,
                        0, 0,
                        rect.width(), rect.height(),
                        win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED
                    )
                    # 强制刷新
                    win32gui.InvalidateRect(tab.explorer_hwnd, None, True)
                    win32gui.UpdateWindow(tab.explorer_hwnd)
                except:
                    pass




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
                action = parent_menu.addAction(f"📑 {node.get('name', '')}")
                url = node.get('url', '')
                action.triggered.connect(lambda checked, u=url: self.open_bookmark_url(u))
        # 直接在菜单栏顶层添加所有书签（前5个是默认书签，已有图标）
        menubar = self.menuBar()
        # 添加所有书签和文件夹
        for idx, child in enumerate(bookmark_bar['children']):
            if child.get('type') == 'folder':
                add_menu_items(menubar, child)
            elif child.get('type') == 'url':
                # 前5个默认书签已有emoji图标，不需要添加📑前缀
                if idx < 5:
                    action = menubar.addAction(child.get('name', ''))
                else:
                    action = menubar.addAction(f"📑 {child.get('name', '')}")
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
        
        # 添加"设置"菜单
        settings_action = QAction("⚙️ 设置", self)
        settings_action.triggered.connect(self.show_settings_dialog)
        menubar.addAction(settings_action)

    def show_bookmark_manager_dialog(self):
        dlg = BookmarkManagerDialog(self.bookmark_manager, self)
        dlg.exec_()
    
    def show_settings_dialog(self):
        """显示设置对话框"""
        dlg = SettingsDialog(self)
        dlg.exec_()
    
    def _init_explorer_monitor(self):
        """初始化资源管理器监控"""
        try:
            self.explorer_monitor = ExplorerMonitor(self)
            # 根据配置决定是否启动监控
            config = self.load_config()
            if config.get('enable_explorer_monitor', True):
                self.explorer_monitor.start()
                print("[MainWindow] 资源管理器监控模块已加载并自动启动")
            else:
                print("[MainWindow] 资源管理器监控模块已加载但未启动（配置已禁用）")
        except ImportError as e:
            print(f"[MainWindow] 无法加载监控模块: {e}")
            print("[MainWindow] 请确保已安装: pip install pywin32 psutil")
            self.explorer_monitor = None
        except Exception as e:
            print(f"[MainWindow] 监控模块初始化错误: {e}")
            self.explorer_monitor = None

# 设置对话框
class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.resize(500, 300)
        self.main_window = parent
        
        layout = QVBoxLayout(self)
        
        # 接管资源管理器选项
        from PyQt5.QtWidgets import QCheckBox, QLabel, QGroupBox
        monitor_group = QGroupBox("资源管理器监控")
        monitor_layout = QVBoxLayout()
        
        self.takeover_checkbox = QCheckBox("⚡ 接管资源管理器（Win+E等）")
        self.takeover_checkbox.setToolTip(
            "开启后，使用Win+E或双击文件夹等方式打开的资源管理器窗口\n"
            "会自动在TabExplorer中打开新标签页。"
        )
        
        # 从配置文件加载设置
        config = self.main_window.load_config()
        self.takeover_checkbox.setChecked(config.get('enable_explorer_monitor', True))
        
        monitor_layout.addWidget(self.takeover_checkbox)
        
        info_label = QLabel(
            "注意：此功能需要安装依赖包：\n"
            "pip install pywin32 psutil"
        )
        info_label.setStyleSheet("color: gray; font-size: 10px;")
        monitor_layout.addWidget(info_label)
        
        monitor_group.setLayout(monitor_layout)
        layout.addWidget(monitor_group)
        
        layout.addStretch()
        
        # 按钮
        button_layout = QHBoxLayout()
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self.save_settings)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
    
    def save_settings(self):
        """保存设置"""
        config = self.main_window.load_config()
        
        # 更新设置
        old_monitor_state = config.get('enable_explorer_monitor', True)
        new_monitor_state = self.takeover_checkbox.isChecked()
        config['enable_explorer_monitor'] = new_monitor_state
        
        # 保存配置
        self.main_window.save_config(config)
        
        # 如果监控状态改变，更新监控
        if old_monitor_state != new_monitor_state:
            if new_monitor_state:
                if self.main_window.explorer_monitor:
                    self.main_window.explorer_monitor.start()
                    QMessageBox.information(self, "已启用", "资源管理器监控已启用")
            else:
                if self.main_window.explorer_monitor:
                    self.main_window.explorer_monitor.stop()
                    QMessageBox.information(self, "已禁用", "资源管理器监控已禁用")
        
        self.accept()


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
            print(f"[Client] Successfully sent path to existing instance: {path}")
            return True
        except Exception as e:
            print(f"[Client] Attempt {attempt + 1}/{max_retries} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(0.1)  # 短暂等待后重试
            continue
    print("[Client] No existing instance found, starting new instance")
    return False

def main():
    # 支持命令行参数：打开指定路径
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
    
    # 启动新实例
    app = QApplication(sys.argv)
    app.setApplicationName("TabExplorer")
    window = MainWindow()
    
    # 如果有路径参数，在新窗口中打开
    # 注意：MainWindow.__init__() 已经调用了 load_pinned_tabs()
    if path_to_open:
        # 无论是否有固定标签页，都添加目标路径作为新标签
        window.add_new_tab(path_to_open)
    
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

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
class BreadcrumbPathBar(QWidget):
    """类似Windows资源管理器的面包屑路径栏，支持点击层级跳转"""
    pathChanged = pyqtSignal(str)  # 当路径改变时发出信号
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_path = ""
        self.edit_mode = False
        self.init_ui()
    
    def init_ui(self):
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(3, 0, 3, 0)
        self.layout.setSpacing(0)
        
        # 路径编辑框（编辑模式时显示）
        self.path_edit = QLineEdit(self)
        self.path_edit.setFixedHeight(30)
        self.path_edit.setStyleSheet("QLineEdit { font-size: 12pt; padding: 3px; border: 1px solid #ccc; }")
        self.path_edit.hide()
        self.path_edit.returnPressed.connect(self.on_edit_finished)
        self.path_edit.editingFinished.connect(self.exit_edit_mode)
        
        # 面包屑容器（显示模式时显示）
        self.breadcrumb_widget = QWidget(self)
        self.breadcrumb_widget.setStyleSheet("QWidget { background: #e8f5e9; }")
        self.breadcrumb_layout = QHBoxLayout(self.breadcrumb_widget)
        self.breadcrumb_layout.setContentsMargins(0, 0, 0, 0)
        self.breadcrumb_layout.setSpacing(0)
        self.breadcrumb_layout.addStretch(1)
        
        self.layout.addWidget(self.breadcrumb_widget)
        self.layout.addWidget(self.path_edit)
        
        # 设置整体样式
        self.setStyleSheet("""
            BreadcrumbPathBar {
                background: #e8f5e9;
                border: 1px solid #ccc;
                border-radius: 2px;
            }
        """)
        self.setFixedHeight(30)
    
    def set_path(self, path):
        """设置并显示路径"""
        self.current_path = path
        if not self.edit_mode:
            self.update_breadcrumbs()
    
    def update_breadcrumbs(self):
        """更新面包屑显示"""
        # 清空现有的面包屑
        while self.breadcrumb_layout.count() > 1:  # 保留最后的stretch
            item = self.breadcrumb_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        if not self.current_path:
            return
        
        # 处理特殊路径
        if self.current_path.startswith('shell:'):
            # shell路径直接显示为一个标签
            label = ClickableLabel(self.current_path, self.current_path)
            label.clicked.connect(self.on_segment_clicked)
            self.breadcrumb_layout.insertWidget(0, label)
            return
        
        # 分割路径
        parts = []
        if os.name == 'nt':
            # Windows路径
            path = self.current_path.replace('/', '\\')
            
            # 检查是否是网络路径（UNC路径）
            is_unc = path.startswith('\\\\')
            
            segments = path.split('\\')
            
            # 构建累积路径
            accumulated = ""
            segment_index = 0
            for i, segment in enumerate(segments):
                if not segment:
                    continue
                
                if is_unc and segment_index == 0:
                    # UNC路径的服务器名
                    accumulated = '\\\\' + segment
                    parts.append((segment, accumulated))
                    segment_index += 1
                elif is_unc and segment_index == 1:
                    # UNC路径的共享名
                    accumulated += '\\' + segment
                    parts.append((segment, accumulated))
                    segment_index += 1
                elif i == 0 and ':' in segment:
                    # 盘符
                    accumulated = segment + '\\'
                    parts.append((segment, accumulated))
                    segment_index += 1
                else:
                    if accumulated and not accumulated.endswith('\\'):
                        accumulated += '\\'
                    accumulated += segment
                    parts.append((segment, accumulated))
                    segment_index += 1
        else:
            # Unix路径
            segments = self.current_path.split('/')
            accumulated = ""
            for segment in segments:
                if not segment:
                    continue
                accumulated += '/' + segment
                parts.append((segment, accumulated))
        
        # 创建面包屑标签
        for i, (name, full_path) in enumerate(parts):
            # 创建可点击的标签
            label = ClickableLabel(name, full_path)
            label.clicked.connect(self.on_segment_clicked)
            self.breadcrumb_layout.insertWidget(i * 2, label)
            
            # 添加分隔符（除了最后一个）
            if i < len(parts) - 1:
                separator = QLabel(">")
                separator.setStyleSheet("QLabel { color: #888; font-size: 11pt; padding: 0 2px; }")
                self.breadcrumb_layout.insertWidget(i * 2 + 1, separator)
    
    def on_segment_clicked(self, path):
        """点击某个层级时触发"""
        self.current_path = path
        self.pathChanged.emit(path)
        self.update_breadcrumbs()
    
    def enter_edit_mode(self):
        """进入编辑模式"""
        self.edit_mode = True
        self.breadcrumb_widget.hide()
        self.path_edit.setText(self.current_path)
        self.path_edit.show()
        self.path_edit.setFocus()
        self.path_edit.selectAll()
    
    def exit_edit_mode(self):
        """退出编辑模式"""
        if self.edit_mode:
            self.edit_mode = False
            self.path_edit.hide()
            self.breadcrumb_widget.show()
            self.update_breadcrumbs()
    
    def on_edit_finished(self):
        """编辑完成时触发"""
        new_path = self.path_edit.text().strip()
        if new_path and new_path != self.current_path:
            self.current_path = new_path
            self.pathChanged.emit(new_path)
        self.exit_edit_mode()
    
    def mouseDoubleClickEvent(self, event: QMouseEvent):
        """双击进入编辑模式"""
        self.enter_edit_mode()
        super().mouseDoubleClickEvent(event)
    
    def mousePressEvent(self, event: QMouseEvent):
        """单击也可以进入编辑模式（点击空白处）"""
        if event.button() == Qt.LeftButton:
            # 检查是否点击在面包屑标签上
            child = self.childAt(event.pos())
            if child is None or child == self.breadcrumb_widget:
                self.enter_edit_mode()
        super().mousePressEvent(event)


class ClickableLabel(QLabel):
    """可点击的标签，用于面包屑导航"""
    clicked = pyqtSignal(str)
    
    def __init__(self, text, path, parent=None):
        super().__init__(text, parent)
        self.path = path
        self.setStyleSheet("""
            QLabel {
                color: #003d7a;
                font-size: 11pt;
                padding: 2px 2px;
                border-radius: 2px;
            }
            QLabel:hover {
                background-color: #cce5ff;
                text-decoration: underline;
            }
        """)
        self.setCursor(Qt.PointingHandCursor)
    
    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.path)
        super().mousePressEvent(event)


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

    def start_path_sync_timer(self):
        from PyQt5.QtCore import QTimer
        self._path_sync_timer = QTimer(self)
        self._path_sync_timer.timeout.connect(self.sync_path_bar_with_explorer)
        self._path_sync_timer.start(500)

    def sync_path_bar_with_explorer(self):
        # 通过QAxWidget的LocationURL属性获取当前路径
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
                    # 只在非程序化导航时添加到历史记录
                    if not self._navigating_programmatically and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    # 同步左侧目录树
                    if self.main_window and hasattr(self.main_window, 'expand_dir_tree_to_path'):
                        self.main_window.expand_dir_tree_to_path(local_path)
        except Exception:
            pass

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 面包屑路径栏
        self.path_bar = BreadcrumbPathBar(self)
        self.path_bar.pathChanged.connect(self.on_path_bar_changed)
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


        # 兼容原有空白双击
        from PyQt5.QtWidgets import QLabel
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

    def __init__(self, parent=None, path="", is_shell=False):
        super().__init__(parent)
        self.main_window = parent
        self.current_path = path if path else QDir.homePath()
        # 浏览历史记录
        self.history = []
        self.history_index = -1
        # 标志：是否正在程序化导航（用于防止sync时重复添加历史）
        self._navigating_programmatically = False
        self.setup_ui()
        self.navigate_to(self.current_path, is_shell=is_shell)
        self.start_path_sync_timer()

    # 移除重复的setup_ui，保留带路径栏的实现

    def navigate_to(self, path, is_shell=False, add_to_history=True):
        # 支持本地路径和shell特殊路径
        if is_shell:
            self.explorer.dynamicCall("Navigate(const QString&)", path)
            self.current_path = path
            if hasattr(self, 'path_bar'):
                self.path_bar.set_path(path)
            self.update_tab_title()
            # 添加到历史记录
            if add_to_history:
                self._add_to_history(path)
        elif os.path.exists(path):
            url = QDir.toNativeSeparators(path)
            self.explorer.dynamicCall("Navigate(const QString&)", url)
            self.current_path = path
            if hasattr(self, 'path_bar'):
                self.path_bar.set_path(path)
            self.update_tab_title()
            # 添加到历史记录
            if add_to_history:
                self._add_to_history(path)
    
    def _add_to_history(self, path):
        """添加路径到历史记录"""
        # 如果当前不在历史末尾，删除当前位置之后的所有历史
        if self.history_index < len(self.history) - 1:
            self.history = self.history[:self.history_index + 1]
        # 添加新路径（避免重复添加相同路径）
        if not self.history or self.history[-1] != path:
            self.history.append(path)
            self.history_index = len(self.history) - 1
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
            self.add_new_tab()


    def go_up_current_tab(self):
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'go_up'):
            current_tab.go_up(force=True)
    
    def go_back_current_tab(self):
        """后退当前标签页"""
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'go_back'):
            current_tab.go_back()
    
    def go_forward_current_tab(self):
        """前进当前标签页"""
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'go_forward'):
            current_tab.go_forward()
    
    def update_navigation_buttons(self):
        """更新前进后退按钮状态"""
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'can_go_back'):
            self.back_button.setEnabled(current_tab.can_go_back())
        else:
            self.back_button.setEnabled(False)
        
        if hasattr(current_tab, 'can_go_forward'):
            self.forward_button.setEnabled(current_tab.can_go_forward())
        else:
            self.forward_button.setEnabled(False)


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
        # 更新导航按钮状态（确保新标签页的按钮状态正确）
        self.update_navigation_buttons()
        
        # 激活窗口（当从其他实例接收到路径时）
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
            # 更新导航按钮状态
            self.update_navigation_buttons()

    def expand_dir_tree_to_path(self, path):
        # 展开并选中左侧目录树到指定路径
        if not hasattr(self, 'dir_model') or not hasattr(self, 'dir_tree'):
            return
        if not path or not os.path.exists(path):
            return
        # 如果是网络路径，直接返回，不展开目录树
        if path.startswith('\\\\'):
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
            path = self.dir_model.filePath(idx)
            # 隐藏网络驱动器（UNC路径以\\开头）
            if path.startswith('\\\\'):
                self.dir_tree.setRowHidden(i, drives_parent, True)
            else:
                self.dir_tree.expand(idx)

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
        
        # 后退按钮
        self.back_button = QPushButton("←")
        self.back_button.setToolTip("后退")
        self.back_button.setFixedHeight(35)
        self.back_button.setFixedWidth(35)
        self.back_button.clicked.connect(self.go_back_current_tab)
        self.back_button.setEnabled(False)
        btn_layout.addWidget(self.back_button)
        
        # 前进按钮
        self.forward_button = QPushButton("→")
        self.forward_button.setToolTip("前进")
        self.forward_button.setFixedHeight(35)
        self.forward_button.setFixedWidth(35)
        self.forward_button.clicked.connect(self.go_forward_current_tab)
        self.forward_button.setEnabled(False)
        btn_layout.addWidget(self.forward_button)
        
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
        
        # 启动单实例通信服务器
        self.start_instance_server()

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
        """窗口关闭时停止服务器"""
        self.server_running = False
        if hasattr(self, 'server_socket'):
            try:
                self.server_socket.close()
            except:
                pass
        super().closeEvent(event)


    def on_dir_tree_clicked(self, index):
        # 目录树点击，右侧当前标签页跳转
        if not index.isValid():
            return
        path = self.dir_model.filePath(index)
        current_tab = self.tab_widget.currentWidget()
        if hasattr(current_tab, 'navigate_to'):
            current_tab.navigate_to(path)

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

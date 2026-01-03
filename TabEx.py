# 全局搜索缓存（LRU缓存，最多缓存50个搜索结果）
from collections import OrderedDict
import hashlib

class SearchCache:
    """搜索结果缓存，使用LRU策略"""
    def __init__(self, max_size=50):
        self.cache = OrderedDict()
        self.max_size = max_size
    
    def get_key(self, search_path, keyword, search_filename, search_content, file_types):
        """生成缓存键"""
        key_str = f"{search_path}|{keyword}|{search_filename}|{search_content}|{file_types}"
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

# 全局搜索缓存实例
_search_cache = SearchCache()

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
        
        # 线程安全的结果队列（限制大小防止内存溢出）
        import queue
        self.result_queue = queue.Queue(maxsize=1000)  # 最多缓存1000个待显示的结果
        self.ui_update_timer = None
        self.queue_overflow_count = 0  # 队列溢出计数
        
        # 结果限制配置（防止UI卡死）
        self.max_results = 5000  # 最多显示5000个结果
        self.current_result_count = 0
        
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
        
        # 清除缓存按钮
        self.clear_cache_btn = QPushButton("🗑️ 清除缓存")
        self.clear_cache_btn.setToolTip("清除所有搜索缓存")
        self.clear_cache_btn.clicked.connect(self.clear_search_cache)
        search_options.addWidget(self.clear_cache_btn)
        
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
        self.ui_update_timer.start(50)  # 每50ms检查一次队列（加快消费速度）
    
    def update_ui_from_queue(self):
        """从队列中取出结果并更新UI（在主线程中调用，批量优化）"""
        try:
            # 批量处理结果（一次处理最多200个，加快队列消费）
            batch_results = []
            batch_count = 0
            max_batch = 200  # 从50增加到20
            
            while batch_count < max_batch:
                try:
                    item = self.result_queue.get_nowait()
                    
                    if item['type'] == 'result':
                        batch_results.append(item)
                        batch_count += 1
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
                except:
                    break  # 队列为空
            
            # 批量添加结果到表格（性能优化）
            if batch_results:
                from PyQt5.QtCore import Qt
                
                # 禁用更新以提高性能
                self.result_list.setUpdatesEnabled(False)
                
                for item in batch_results:
                    # 检查是否超过最大结果数
                    if self.current_result_count >= self.max_results:
                        break
                    
                    row = self.result_list.rowCount()
                    self.result_list.insertRow(row)
                    
                    # 文件名项
                    name_item = QTableWidgetItem(item['name'])
                    name_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    name_item.setToolTip(item['full_path'])
                    self.result_list.setItem(row, 0, name_item)
                    self.result_list.setItem(row, 1, QTableWidgetItem(item['file_type']))
                    self.result_list.setItem(row, 2, QTableWidgetItem(item['date']))
                    self.result_list.setItem(row, 3, QTableWidgetItem(item['size']))
                    self.result_list.item(row, 0).setData(256, item['path'])
                    
                    self.current_result_count += 1
                
                # 重新启用更新
                self.result_list.setUpdatesEnabled(True)
                
        except Exception as e:
            print(f"[Search] UI update error: {e}")
    
    def add_search_result(self, text):
        """添加搜索结果项（通过队列，线程安全）"""
        self.result_queue.put({'type': 'result', 'text': text})
    
    def clear_search_cache(self):
        """清除所有搜索缓存"""
        global _search_cache
        _search_cache.clear()
        QMessageBox.information(self, "提示", "搜索缓存已清除")
    
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
        
        # 获取文件类型过滤
        file_types = self.file_type_input.text().strip()
        
        # 检查缓存
        global _search_cache
        cache_key = _search_cache.get_key(
            search_path, keyword, 
            self.search_filename_cb.isChecked(), 
            self.search_content_cb.isChecked(), 
            file_types
        )
        cached_results = _search_cache.get(cache_key)
        
        if cached_results is not None:
            # 使用缓存结果
            print(f"[Search] 使用缓存结果，共 {len(cached_results)} 个")
            self.result_list.setRowCount(0)
            self.status_label.setText("正在加载缓存结果...")
            
            # 批量添加缓存结果
            sorting_enabled = self.result_list.isSortingEnabled()
            self.result_list.setSortingEnabled(False)
            
            for item in cached_results:
                self.result_queue.put({'type': 'result', **item})
            
            self.result_list.setSortingEnabled(sorting_enabled)
            self.status_label.setText(f"搜索完成（缓存），共找到 {len(cached_results)} 个结果")
            return
        
        # 清空之前的结果
        self.result_list.setRowCount(0)
        self.current_result_count = 0  # 重置计数器
        
        # 搜索期间完全禁用排序（性能优化）
        self.result_list.setSortingEnabled(False)
        
        self.is_searching = True
        self.search_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText(f"搜索中... (最多显示{self.max_results}个结果)")
        
        # 在后台线程执行搜索
        import threading
        self.search_thread = threading.Thread(
            target=self.do_search,
            args=(keyword, self.search_filename_cb.isChecked(), self.search_content_cb.isChecked(), file_types, cache_key)
        )
        self.search_thread.daemon = True
        self.search_thread.start()
    
    def stop_search(self):
        self.is_searching = False
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("已停止")
    
    def do_search(self, keyword, search_filename, search_content, file_types="", cache_key=None):
        found_count = 0
        keyword_lower = keyword.lower()
        results_buffer = []  # 结果缓冲区
        buffer_size = 100  # 每100个结果批量更新一次（优化：从20增加到100）
        all_results = []  # 存储所有结果用于缓存（限制大小）
        
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
        
        def is_text_file(file_path, sample_size=1024):
            """智能检测文件是否为文本文件（读取前N字节检测）"""
            try:
                with open(file_path, 'rb') as f:
                    sample = f.read(sample_size)
                    if not sample:
                        return True  # 空文件视为文本文件
                    
                    # 检测二进制字符的比例
                    # 文本文件中NULL字节和其他控制字符应该很少
                    null_count = sample.count(b'\x00')
                    
                    # 如果包含NULL字节，很可能是二进制文件
                    if null_count > 0:
                        return False
                    
                    # 统计非打印字符（排除常见的\r\n\t）
                    non_printable = 0
                    for byte in sample:
                        # ASCII可打印字符范围：32-126，加上\r(13)\n(10)\t(9)
                        if byte < 9 or (byte > 13 and byte < 32) or byte > 126:
                            non_printable += 1
                    
                    # 如果超过10%是非打印字符，可能是二进制文件
                    if non_printable / len(sample) > 0.1:
                        return False
                    
                    return True
            except:
                return False  # 无法读取则视为二进制
        
        # 对于文件内容搜索，使用编译的正则表达式可能更快（可选优化）
        # 但Python的内置字符串搜索已经很快，这里保持简单
        
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
            skipped_binary_files = 0  # 跳过的二进制文件数
            for root, dirs, files in os.walk(self.search_path):
                if not self.is_searching:
                    print("[Search] 搜索被中断")
                    break
                
                folder_count += 1
                
                # 搜索文件夹名
                if search_filename:
                    for dirname in dirs:
                        if not self.is_searching:
                            break
                        
                        # 检查是否达到结果限制
                        if found_count >= max_results:
                            results_limited = True
                            break
                        
                        # 使用Python内置的字符串搜索（已优化）
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
                            
                            result_item = {
                                'path': dir_path,
                                'name': f"📁 {dirname}",
                                'full_path': f"📁 {dir_path}",
                                'file_type': '文件夹',
                                'date': mtime,
                                'size': size_str
                            }
                            results_buffer.append(result_item)
                            all_results.append(result_item)  # 保存到缓存列表
                            
                            # 批量更新UI（队列满时等待）
                            if len(results_buffer) >= buffer_size:
                                for item in results_buffer:
                                    # 使用带超时的put，队列满时等待最多0.5秒
                                    try:
                                        self.result_queue.put({'type': 'result', **item}, timeout=0.5)
                                    except:
                                        # 超时后跳过该结果
                                        self.queue_overflow_count += 1
                                results_buffer.clear()
                
                # 检查是否达到结果限制
                if found_count >= max_results:
                    results_limited = True
                    break
                
                # 搜索文件名和文件内容
                for filename in files:
                    if not self.is_searching:
                        print("[Search] 搜索被中断（文件循环）")
                        break
                    
                    # 检查是否达到结果限制
                    if found_count >= max_results:
                        results_limited = True
                        break
                    
                    # 优化：如果只搜索文件名，快速过滤不匹配的文件
                    if search_filename and not search_content:
                        if keyword_lower not in filename.lower():
                            continue  # 文件名不匹配，跳过
                    
                    # 检查文件类型过滤
                    if not matches_file_type(filename):
                        # 调试：显示被过滤的文件（仅对特定文件名）
                        if 'TstMgr' in filename or scanned_files < 5:
                            print(f"[Search] 文件被类型过滤跳过: {filename}")
                        continue  # 跳过不匹配的文件类型
                    
                    scanned_files += 1
                    
                    # 每扫描100个文件更新一次状态（减少队列压力）
                    if scanned_files % 100 == 0:
                        # 使用带超时的put，确保状态能更新
                        status_text = f"搜索中... 已扫描 {scanned_files} 个文件，找到 {found_count} 个结果"
                        try:
                            self.result_queue.put({'type': 'status', 'text': status_text}, timeout=0.1)
                        except:
                            pass  # 超时后继续搜索
                    
                    file_path = os.path.join(root, filename)
                    matched = False
                    match_type = ""
                    
                    # 调试：显示正在搜索的特定文件
                    if 'TstMgr_RtnSound.c' in filename:
                        print(f"[Search] 正在搜索文件: {file_path}")
                        print(f"[Search] 搜索文件名: {search_filename}, 搜索内容: {search_content}")
                    
                    # 搜索文件名（Python内置优化）
                    if search_filename and keyword_lower in filename.lower():
                        matched = True
                        match_type = "📄"
                    
                    # 搜索文件内容（智能检测文本文件）
                    if search_content and not matched:
                        # 获取文件扩展名
                        _, ext = os.path.splitext(filename)
                        file_ext = ext[1:].lower() if ext else ''  # 去掉点号并转小写
                        
                        # 1. 首先检查黑名单（明确的二进制文件）
                        if file_ext in binary_file_extensions:
                            skipped_binary_files += 1
                            continue
                        
                        # 2. 对于其他文件，智能检测是否为文本文件
                        if not is_text_file(file_path):
                            skipped_binary_files += 1
                            continue
                        
                        # 调试信息
                        if 'TstMgr_RtnSound.c' in filename:
                            print(f"[Search] 开始搜索文件内容: {file_path}")
                        
                        try:
                            # 获取文件大小
                            file_size = os.path.getsize(file_path)
                            
                            # 分块读取大文件，每次读取10MB
                            chunk_size = 10 * 1024 * 1024  # 10MB
                            
                            # 尝试多种编码方式读取文件内容
                            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                            content_matched = False
                            
                            for encoding in encodings:
                                try:
                                    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                                        if file_size <= chunk_size:
                                            # 小文件直接全部读取（Python内置字符串搜索已优化）
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
                                            # 大文件分块读取（优化：使用内存映射）
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
                        
                        result_item = {
                            'path': file_path,
                            'name': f"{match_type} {path_without_ext}",
                            'full_path': f"{match_type} {file_path}",
                            'file_type': file_type,
                            'date': mtime,
                            'size': size_str
                        }
                        results_buffer.append(result_item)
                        all_results.append(result_item)  # 保存到缓存列表
                        
                        # 批量更新UI（每20个结果更新一次）
                        if len(results_buffer) >= buffer_size:
                            # 将结果放入队列（队列满时等待）
                            for item in results_buffer:
                                try:
                                    # 使用带超时的put，队列满时等待最多0.5秒
                                    self.result_queue.put({'type': 'result', **item}, timeout=0.5)
                                except:
                                    # 超时后跳过该结果
                                    self.queue_overflow_count += 1
                            results_buffer.clear()
        except Exception as e:
            print(f"Search error: {e}")
        
        # 添加剩余的结果（队列满时等待）
        if results_buffer:
            for item in results_buffer:
                try:
                    self.result_queue.put({'type': 'result', **item}, timeout=0.5)
                except:
                    self.queue_overflow_count += 1
        
        # 调试信息
        print(f"[Search] 搜索完成，共扫描 {scanned_files} 个文件，找到 {found_count} 个结果")
        if search_content and skipped_binary_files > 0:
            print(f"[Search] 跳过 {skipped_binary_files} 个二进制文件（不搜索内容）")
        if self.queue_overflow_count > 0:
            print(f"[Search] ⚠️ 队列溢出 {self.queue_overflow_count} 次（部分结果未显示）")
        
        # 将结果存入缓存（限制缓存大小，防止内存溢出）
        if cache_key and all_results:
            global _search_cache
            # 只缓存前max_results个结果
            cached_results = all_results[:max_results]
            _search_cache.put(cache_key, cached_results)
            print(f"[Search] 已将 {len(cached_results)} 个结果存入缓存")
        
        # 重置搜索状态（先重置，避免后续更新被跳过）
        self.is_searching = False
        self.queue_overflow_count = 0  # 重置溢出计数
        
        # 搜索完成，更新UI状态（使用带超时的put，避免卡死）
        if results_limited:
            final_status = f"搜索完成（已限制），显示前 {found_count} 个结果（扫描了 {scanned_files} 个文件）⚠️"
        else:
            final_status = f"搜索完成，共找到 {found_count} 个结果（扫描了 {scanned_files} 个文件）"
        
        # 使用超时put，防止队列满时卡死
        try:
            self.result_queue.put({'type': 'status', 'text': final_status}, timeout=1)
            self.result_queue.put({'type': 'button', 'button': 'search', 'enabled': True}, timeout=1)
            self.result_queue.put({'type': 'button', 'button': 'stop', 'enabled': False}, timeout=1)
            # 搜索完成后启用排序
            self.result_queue.put({'type': 'enable_sorting'}, timeout=1)
        except:
            print(f"[Search] ⚠️ 队列满，最终状态更新失败")
        
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
import queue
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QListWidget, QLabel, QToolBar, QAction, QMenu, QMessageBox, QInputDialog, QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView, QSizePolicy)  # QDockWidget removed (unused)
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import Qt, QDir, QUrl, pyqtSignal, pyqtSlot, Q_ARG, QObject, QSize  # QModelIndex removed (unused)
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QMouseEvent, QCursor
# from PyQt5.QtGui import QIcon  # unused


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
        print(f"Failed to setup Windows API monitoring: {e}")

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
        # 优化：延迟保存机制，避免频繁写入磁盘
        self._save_timer = None
        self._pending_save = False

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

    def save_bookmarks(self, immediate=False):
        # 优化：延迟保存，避免频繁操作时多次写入
        if immediate:
            # 立即保存
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({"roots": self.bookmark_tree}, f, ensure_ascii=False, indent=2)
            self._pending_save = False
        else:
            # 延迟保存：500ms后执行
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
        # 优化：从500ms改为1000ms，减少CPU占用
        self._path_sync_timer.start(1000)

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
        
        # 初始导航到当前路径（在setup_ui最后调用，确保所有设置已应用）
        self.explorer.dynamicCall('Navigate(const QString&)', QDir.toNativeSeparators(self.current_path))

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
        # 优化：使用多重检测避免误判
        # 1. 检查鼠标按下时是否点击在项目上（原生API）
        # 2. 检查是否有选中项
        # 3. 检查路径是否发生变化（多次检查）
        from PyQt5.QtCore import QTimer
        
        # 记录双击前的路径，用于检测是否发生了导航
        try:
            self._path_before_double = getattr(self, 'current_path', None)
        except Exception:
            self._path_before_double = None
        
        def _check_and_go_up():
            # 优先级1: 如果按下时点击测试显示点击在项目上，直接跳过
            if getattr(self, '_clicked_on_item', False):
                self._clicked_on_item = False
                return
            
            # 优先级2: 如果路径已经改变（说明进入了文件夹），跳过
            try:
                before_path = getattr(self, '_path_before_double', None)
                cur_path = getattr(self, 'current_path', None)
                if before_path is not None and cur_path is not None and cur_path != before_path:
                    # 路径已改变，说明双击了文件夹并进入了，清理标志
                    self._path_before_double = None
                    return
            except Exception:
                pass
            
            # 优先级3: 如果按下时有选中项，跳过
            before = getattr(self, '_selected_before_click', None)
            if before is not None:
                try:
                    if int(before) > 0:
                        return
                except Exception:
                    pass
            
            # 确认是空白区域双击，返回上一级
            try:
                self.go_up(force=True)
            except Exception:
                pass
            finally:
                # 清理标志
                self._path_before_double = None

        # 延迟150ms检查，给Explorer充足的时间完成导航
        QTimer.singleShot(150, _check_and_go_up)

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
        # 在鼠标按下时记录当时的选中项数量和鼠标位置点击测试结果
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
                    # 取消所有之前待处理的双击检查，避免多次操作时的累积效应
                    try:
                        for timer in getattr(self, '_pending_double_click_timers', []):
                            try:
                                timer.stop()
                            except Exception:
                                pass
                        self._pending_double_click_timers = []
                    except Exception:
                        self._pending_double_click_timers = []
                    
                    # 为这次双击生成唯一ID
                    self._double_click_id = getattr(self, '_double_click_id', 0) + 1
                    current_click_id = self._double_click_id
                    
                    # 延迟检查，避免与控件自身处理产生竞态
                    # 记录双击发生前的路径，以便判断双击是否触发了导航
                    path_before = getattr(self, 'current_path', None)
                    selected_before = getattr(self, '_selected_before_click', None)
                    
                    print(f"[DoubleClick] ID={current_click_id}, path_before='{path_before}', selected_before={selected_before}")

                    # try multiple times because folder navigation can be slower;
                    # perform checks at 150ms, 300ms, 600ms, 1000ms before giving up
                    delays = [150, 300, 600, 1000]
                    def attempt(idx=0):
                        # 检查这次双击是否还是当前的双击（通过ID验证）
                        # 如果ID不匹配，说明有新的双击发生了，放弃当前检查
                        current_id = getattr(self, '_double_click_id', 0)
                        if current_id != current_click_id:
                            print(f"[DoubleClick] ID mismatch: current={current_id} vs expected={current_click_id}, abort")
                            return
                        
                        cur_path = getattr(self, 'current_path', None)
                        print(f"[DoubleClick] Check attempt {idx} (ID={current_click_id}): path_before='{path_before}' -> cur_path='{cur_path}'")
                        
                        handled = False
                        try:
                            # reuse the same logic as _check_and_go_up body
                            # check path change first
                            try:
                                if path_before is not None and cur_path is not None and cur_path != path_before:
                                    # 路径已改变，说明双击触发了导航（进入文件夹），不需要go_up
                                    print(f"[DoubleClick] Path changed (ID={current_click_id}): '{path_before}' -> '{cur_path}', skip go_up")
                                    return
                            except Exception as e:
                                print(f"[DoubleClick] Path check exception: {e}")
                                pass
                            # if press-time selection existed, skip
                            before = getattr(self, '_selected_before_click', None)
                            if before is not None:
                                try:
                                    if int(before) > 0:
                                        print(f"[DoubleClick] Had selection before click (ID={current_click_id}): {before}, skip go_up")
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
                                        print(f"[DoubleClick] Hit test positive (ID={current_click_id}), skip go_up")
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
                            
                            print(f"[DoubleClick] SelectedItems count (ID={current_click_id}): {cnt}")
                            
                            if cnt is None:
                                # SelectedItems().Count API不可用，使用其他方法判断
                                
                                # 首先检查路径是否变化
                                try:
                                    cur_path = getattr(self, 'current_path', None)
                                    if path_before is not None and cur_path is not None and cur_path != path_before:
                                        # 路径已改变，说明双击触发了导航（进入文件夹），不需要go_up
                                        print(f"[DoubleClick] cnt=None but path changed (ID={current_click_id}): '{path_before}' -> '{cur_path}', skip go_up")
                                        return
                                except Exception:
                                    pass
                                
                                # 路径未变化，继续判断
                                # 如果还有重试机会，继续等待
                                if idx < len(delays) - 1:
                                    print(f"[DoubleClick] cnt=None, schedule retry {idx+1} (ID={current_click_id})")
                                    timer = QTimer()
                                    timer.setSingleShot(True)
                                    timer.timeout.connect(lambda: attempt(idx+1))
                                    timer.start(delays[idx+1] - delays[idx])
                                    if hasattr(self, '_pending_double_click_timers'):
                                        self._pending_double_click_timers.append(timer)
                                    return
                                
                                # 最后一次尝试，路径未变化，cnt仍为None
                                # 优先使用native hit-test判断是否点击在项目上
                                if HAS_PYWIN:
                                    try:
                                        from PyQt5.QtGui import QCursor
                                        gx = QCursor.pos().x()
                                        gy = QCursor.pos().y()
                                        if self._native_listview_hit_test(gx, gy):
                                            print(f"[DoubleClick] Final check: hit test positive (ID={current_click_id}), likely clicked on item, skip go_up")
                                            return
                                    except Exception as e:
                                        print(f"[DoubleClick] Hit test exception: {e}")
                                        pass
                                
                                # 使用selected_before来判断：如果按下时有选中项，可能是文件双击
                                before = getattr(self, '_selected_before_click', None)
                                if before is not None and before > 0:
                                    # 按下时有选中项，可能是文件双击（打开文件不改变路径）
                                    print(f"[DoubleClick] cnt=None but had selection before (ID={current_click_id}): {before}, skip go_up")
                                    return
                                
                                # 没有选中项，路径也没变化，hit-test也是负的，很可能是空白双击
                                print(f"[DoubleClick] Execute go_up (ID={current_click_id}): cnt=None but no selection, no hit, path unchanged")
                                try:
                                    self.go_up(force=True)
                                except Exception:
                                    pass
                                return
                            try:
                                if int(cnt) == 0:
                                    print(f"[DoubleClick] Execute go_up (ID={current_click_id}): cnt=0, blank area double-click")
                                    self.go_up(force=True)
                                else:
                                    print(f"[DoubleClick] Has selection (ID={current_click_id}): cnt={cnt}, skip go_up")
                                return
                            except Exception:
                                pass
                        finally:
                            self._selected_before_click = None

                    timer = QTimer()
                    timer.setSingleShot(True)
                    timer.timeout.connect(lambda: attempt(0))
                    timer.start(delays[0])
                    if hasattr(self, '_pending_double_click_timers'):
                        self._pending_double_click_timers.append(timer)
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def blank_double_click(self, event):
        self.go_up(force=True)

    # 移除 on_document_complete 和 eventFilter 相关内容

    def go_up(self, force=False):
        # 返回上一级目录，盘符根目录时导航到"此电脑"
        # 如果 force=True，则绕过鼠标位置检查（用于按钮或程序化调用）
        current_path_before = getattr(self, 'current_path', None)
        print(f"[go_up] Called with force={force}, current_path='{current_path_before}'")
        
        if not force:
            # 仅在明确来自空白区域或路径栏的触发时执行，避免误由文件双击触发
            try:
                from PyQt5.QtWidgets import QApplication
                from PyQt5.QtGui import QCursor
                pos = QCursor.pos()
                w = QApplication.widgetAt(pos.x(), pos.y())
                # 允许的触发源：底部空白标签或路径栏
                if w is not self.blank and w is not getattr(self, 'path_bar', None):
                    print(f"[go_up] Rejected: not from valid source (widget={w})")
                    return
            except Exception:
                # 如果无法判断，保守退出，避免误导航
                print(f"[go_up] Rejected: exception in widget check")
                return
        if not self.current_path:
            print(f"[go_up] Rejected: no current_path")
            return
        path = self.current_path
        # 判断是否为盘符根目录，导航到"此电脑"
        if path.endswith(":\\") or path.endswith(":/"):
            print(f"[go_up] Root directory '{path}', navigate to MyComputer")
            self.navigate_to('shell:MyComputerFolder', is_shell=True)
            return
        parent_path = os.path.dirname(path)
        if parent_path and os.path.exists(parent_path):
            print(f"[go_up] Navigate from '{path}' to '{parent_path}'")
            self.navigate_to(parent_path)
        else:
            print(f"[go_up] Rejected: parent_path '{parent_path}' invalid or not exists")

    def __init__(self, parent=None, path="", is_shell=False):
        super().__init__(parent)
        self.main_window = parent
        self.current_path = path if path else QDir.homePath()
        # 浏览历史记录
        self.history = []
        self.history_index = -1
        # 标志：是否正在程序化导航（用于防止sync时重复添加历史）
        self._navigating_programmatically = False
        # 用于跟踪待处理的双击检查定时器
        self._pending_double_click_timers = []
        # 双击事件唯一ID，用于区分不同的双击操作
        self._double_click_id = 0
        self.setup_ui()
        self.navigate_to(self.current_path, is_shell=is_shell)
        self.start_path_sync_timer()

    # 移除重复的setup_ui，保留带路径栏的实现

    def navigate_to(self, path, is_shell=False, add_to_history=True):
        old_path = getattr(self, 'current_path', None)
        print(f"[navigate_to] From '{old_path}' to '{path}' (is_shell={is_shell})")
        
        # 取消所有待处理的双击检查定时器，因为路径即将改变
        try:
            cancelled_count = 0
            for timer in getattr(self, '_pending_double_click_timers', []):
                try:
                    timer.stop()
                    cancelled_count += 1
                except Exception:
                    pass
            self._pending_double_click_timers = []
            if cancelled_count > 0:
                print(f"[navigate_to] Cancelled {cancelled_count} pending double-click timers")
        except Exception:
            pass
        
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
            # 使用Navigate2来获得更好的控制
            # navNoHistory = 0x2, navNoReadFromCache = 0x4, navNoWriteToCache = 0x8
            # navAllowAutosearch = 0x10, navBrowserBar = 0x20, navHyperlink = 0x40
            try:
                # 尝试使用Navigate2获得更好的刷新效果
                self.explorer.dynamicCall("Navigate2(QVariant,QVariant,QVariant,QVariant,QVariant)", 
                                         url, 0, "", None, None)
            except:
                # 回退到普通Navigate
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
    def create_custom_titlebar(self, main_layout):
        """创建自定义标题栏，包含窗口控制按钮和功能按钮"""
        titlebar = QWidget()
        titlebar.setFixedHeight(32)
        titlebar.setStyleSheet("background-color: #f0f0f0; border-bottom: 1px solid #ccc;")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(10, 0, 0, 0)
        titlebar_layout.setSpacing(0)
        
        # 窗口标题
        title_label = QLabel("TabExplorer")
        title_label.setStyleSheet("font-weight: bold; font-size: 11pt; color: #333;")
        titlebar_layout.addWidget(title_label)
        
        # 用于拖动窗口
        self.titlebar_widget = titlebar
        self.drag_position = None
        
        titlebar_layout.addStretch()
        
        # 书签管理按钮
        bookmark_btn = QPushButton("📑")
        bookmark_btn.setToolTip("书签管理")
        bookmark_btn.setFixedSize(40, 32)
        bookmark_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 14pt;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
            QPushButton:pressed {
                background: rgba(0, 0, 0, 0.2);
            }
        """)
        bookmark_btn.clicked.connect(self.show_bookmark_manager_dialog)
        titlebar_layout.addWidget(bookmark_btn)
        
        # 设置按钮
        settings_btn = QPushButton("⚙️")
        settings_btn.setToolTip("设置")
        settings_btn.setFixedSize(40, 32)
        settings_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 14pt;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
            QPushButton:pressed {
                background: rgba(0, 0, 0, 0.2);
            }
        """)
        settings_btn.clicked.connect(self.show_settings_menu)
        titlebar_layout.addWidget(settings_btn)
        
        # 最小化按钮
        min_btn = QPushButton("—")
        min_btn.setFixedSize(45, 32)
        min_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 16pt;
                font-weight: bold;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
        """)
        min_btn.clicked.connect(self.showMinimized)
        titlebar_layout.addWidget(min_btn)
        
        # 最大化/还原按钮
        self.max_btn = QPushButton("□")
        self.max_btn.setFixedSize(45, 32)
        self.max_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 16pt;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
        """)
        self.max_btn.clicked.connect(self.toggle_maximize)
        titlebar_layout.addWidget(self.max_btn)
        
        # 关闭按钮
        close_btn = QPushButton("✕")
        close_btn.setFixedSize(45, 32)
        close_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 16pt;
            }
            QPushButton:hover {
                background: #e81123;
                color: white;
            }
        """)
        close_btn.clicked.connect(self.close)
        titlebar_layout.addWidget(close_btn)
        
        main_layout.addWidget(titlebar)
    
    def toggle_maximize(self):
        """切换最大化/还原窗口"""
        if self.isMaximized():
            self.showNormal()
            self.max_btn.setText("□")
        else:
            self.showMaximized()
            self.max_btn.setText("❐")
    
    def mousePressEvent(self, event):
        """鼠标按下事件 - 用于拖动窗口"""
        if event.button() == Qt.LeftButton and hasattr(self, 'titlebar_widget'):
            # 检查点击位置是否在标题栏内
            titlebar_rect = self.titlebar_widget.geometry()
            if titlebar_rect.contains(event.pos()):
                self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
                event.accept()
        super().mousePressEvent(event)
    
    def mouseMoveEvent(self, event):
        """鼠标移动事件 - 拖动窗口"""
        if event.buttons() == Qt.LeftButton and self.drag_position is not None:
            if not self.isMaximized():
                self.move(event.globalPos() - self.drag_position)
                event.accept()
        super().mouseMoveEvent(event)
    
    def mouseReleaseEvent(self, event):
        """鼠标释放事件"""
        self.drag_position = None
        super().mouseReleaseEvent(event)
    
    def create_custom_titlebar(self, main_layout):
        """创建自定义标题栏，包含窗口控制按钮和功能按钮"""
        titlebar = QWidget()
        titlebar.setFixedHeight(32)
        titlebar.setStyleSheet("background-color: #f0f0f0; border-bottom: 1px solid #ccc;")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(10, 0, 0, 0)
        titlebar_layout.setSpacing(0)
        
        # 窗口标题
        title_label = QLabel("TabExplorer")
        title_label.setStyleSheet("font-weight: bold; font-size: 11pt; color: #333;")
        titlebar_layout.addWidget(title_label)
        
        # 用于拖动窗口
        self.titlebar_widget = titlebar
        self.drag_position = None
        
        titlebar_layout.addStretch()
        
        # 书签管理按钮
        bookmark_btn = QPushButton("📑")
        bookmark_btn.setToolTip("书签管理")
        bookmark_btn.setFixedSize(40, 32)
        bookmark_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 14pt;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
            QPushButton:pressed {
                background: rgba(0, 0, 0, 0.2);
            }
        """)
        bookmark_btn.clicked.connect(self.show_bookmark_manager_dialog)
        titlebar_layout.addWidget(bookmark_btn)
        
        # 设置按钮
        settings_btn = QPushButton("⚙️")
        settings_btn.setToolTip("设置")
        settings_btn.setFixedSize(40, 32)
        settings_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 14pt;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
            QPushButton:pressed {
                background: rgba(0, 0, 0, 0.2);
            }
        """)
        settings_btn.clicked.connect(self.show_settings_menu)
        titlebar_layout.addWidget(settings_btn)
        
        # 最小化按钮
        min_btn = QPushButton("—")
        min_btn.setFixedSize(45, 32)
        min_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 16pt;
                font-weight: bold;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
        """)
        min_btn.clicked.connect(self.showMinimized)
        titlebar_layout.addWidget(min_btn)
        
        # 最大化/还原按钮
        self.max_btn = QPushButton("☐")
        self.max_btn.setFixedSize(45, 32)
        self.max_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 16pt;
            }
            QPushButton:hover {
                background: rgba(0, 0, 0, 0.1);
            }
        """)
        self.max_btn.clicked.connect(self.toggle_maximize)
        titlebar_layout.addWidget(self.max_btn)
        
        # 关闭按钮
        close_btn = QPushButton("✕")
        close_btn.setFixedSize(45, 32)
        close_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                font-size: 16pt;
            }
            QPushButton:hover {
                background: #e81123;
                color: white;
            }
        """)
        close_btn.clicked.connect(self.close)
        titlebar_layout.addWidget(close_btn)
        
        main_layout.addWidget(titlebar)
    
    def toggle_maximize(self):
        """切换最大化/还原窗口"""
        if self.isMaximized():
            self.showNormal()
            self.max_btn.setText("☐")
        else:
            self.showMaximized()
            self.max_btn.setText("❐")
    
    def mousePressEvent(self, event):
        """鼠标按下事件 - 用于拖动窗口或调整大小"""
        if event.button() == Qt.LeftButton:
            # 检测是否在边缘（调整大小）
            edge = self.detect_edge(event.pos())
            if edge and not self.isMaximized():
                self.resizing = True
                self.resize_direction = edge
                self.resize_start_pos = event.globalPos()
                self.resize_start_geometry = self.geometry()
                event.accept()
                return
            
            # 检查点击位置是否在标题栏内（拖动窗口）
            if hasattr(self, 'titlebar_widget'):
                titlebar_rect = self.titlebar_widget.geometry()
                if titlebar_rect.contains(event.pos()):
                    self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
                    event.accept()
        super().mousePressEvent(event)
    
    def mouseMoveEvent(self, event):
        """鼠标移动事件 - 拖动窗口或调整大小"""
        if event.buttons() == Qt.LeftButton:
            # 调整窗口大小
            if self.resizing and self.resize_direction:
                self.resize_window(event.globalPos())
                # 调整大小时保持调整大小的光标
                self.update_cursor(self.resize_direction)
                event.accept()
                return
            
            # 拖动窗口
            if self.drag_position is not None and not self.isMaximized():
                self.move(event.globalPos() - self.drag_position)
                event.accept()
                return
        else:
            # 只有在没有按键按下时才更新光标（仅在未最大化时）
            if not self.isMaximized():
                edge = self.detect_edge(event.pos())
                if edge:
                    # 在边缘，显示调整大小光标
                    self.update_cursor(edge)
                    event.accept()
                    return
                else:
                    # 不在边缘，恢复默认光标（通过QApplication恢复）
                    self.update_cursor(None)
        
        super().mouseMoveEvent(event)
    
    def mouseReleaseEvent(self, event):
        """鼠标释放事件"""
        self.drag_position = None
        self.resizing = False
        self.resize_direction = None
        # 释放后恢复默认光标，避免停留在调整大小形状
        try:
            self.update_cursor(None)
        except Exception:
            pass
        super().mouseReleaseEvent(event)
    
    def leaveEvent(self, event):
        """鼠标离开窗口时恢复覆盖的光标"""
        from PyQt5.QtWidgets import QApplication
        if getattr(self, 'cursor_overridden', False):
            QApplication.restoreOverrideCursor()
            self.cursor_overridden = False
        super().leaveEvent(event)
    
    def detect_edge(self, pos):
        """检测鼠标是否在窗口边缘，返回边缘方向"""
        rect = self.rect()
        margin = self.resize_margin
        
        left = pos.x() <= margin
        right = pos.x() >= rect.width() - margin
        top = pos.y() <= margin
        bottom = pos.y() >= rect.height() - margin
        
        if top and left:
            return 'top-left'
        elif top and right:
            return 'top-right'
        elif bottom and left:
            return 'bottom-left'
        elif bottom and right:
            return 'bottom-right'
        elif top:
            return 'top'
        elif bottom:
            return 'bottom'
        elif left:
            return 'left'
        elif right:
            return 'right'
        return None
    
    def update_cursor(self, edge):
        """根据边缘位置更新鼠标光标（使用QApplication覆盖，避免子控件干扰）"""
        from PyQt5.QtGui import QCursor
        from PyQt5.QtWidgets import QApplication

        def apply(shape):
            if not self.cursor_overridden:
                QApplication.setOverrideCursor(QCursor(shape))
                self.cursor_overridden = True
            else:
                # 已覆盖则变更形状
                QApplication.changeOverrideCursor(QCursor(shape))

        def clear():
            if self.cursor_overridden:
                QApplication.restoreOverrideCursor()
                self.cursor_overridden = False

        if edge == 'top-left' or edge == 'bottom-right':
            apply(Qt.SizeFDiagCursor)
        elif edge == 'top-right' or edge == 'bottom-left':
            apply(Qt.SizeBDiagCursor)
        elif edge == 'top' or edge == 'bottom':
            apply(Qt.SizeVerCursor)
        elif edge == 'left' or edge == 'right':
            apply(Qt.SizeHorCursor)
        else:
            clear()
    
    def resize_window(self, global_pos):
        """根据鼠标位置调整窗口大小"""
        delta = global_pos - self.resize_start_pos
        old_geo = self.resize_start_geometry
        
        # 计算新的位置和大小
        new_x = old_geo.x()
        new_y = old_geo.y()
        new_width = old_geo.width()
        new_height = old_geo.height()
        
        # 根据拖动方向调整窗口位置和大小
        if 'left' in self.resize_direction:
            new_x = old_geo.x() + delta.x()
            new_width = old_geo.width() - delta.x()
        
        if 'right' in self.resize_direction:
            new_width = old_geo.width() + delta.x()
        
        if 'top' in self.resize_direction:
            new_y = old_geo.y() + delta.y()
            new_height = old_geo.height() - delta.y()
        
        if 'bottom' in self.resize_direction:
            new_height = old_geo.height() + delta.y()
        
        # 应用最小尺寸限制
        if new_width < self.minimumWidth():
            if 'left' in self.resize_direction:
                new_x = old_geo.x() + old_geo.width() - self.minimumWidth()
            new_width = self.minimumWidth()
        
        if new_height < self.minimumHeight():
            if 'top' in self.resize_direction:
                new_y = old_geo.y() + old_geo.height() - self.minimumHeight()
            new_height = self.minimumHeight()
        
        # 一次性设置新的几何形状
        self.setGeometry(new_x, new_y, new_width, new_height)
    
    def mouseDoubleClickEvent(self, event):
        """鼠标双击事件 - 双击标题栏切换最大化/还原"""
        if event.button() == Qt.LeftButton and hasattr(self, 'titlebar_widget'):
            # 检查双击位置是否在标题栏范围内
            titlebar_pos = self.titlebar_widget.mapFrom(self, event.pos())
            if self.titlebar_widget.rect().contains(titlebar_pos):
                # 排除按钮区域（避免双击按钮触发最大化）
                clicked_widget = self.titlebar_widget.childAt(titlebar_pos)
                if clicked_widget is None or isinstance(clicked_widget, QLabel):
                    self.toggle_maximize()
                    event.accept()
                    return
        super().mouseDoubleClickEvent(event)
    
    def show_settings_menu(self):
        """显示设置对话框"""
        dlg = SettingsDialog(self.config, self)
        if dlg.exec_():
            # 获取新配置
            old_monitor = self.config.get("enable_explorer_monitor", True)
            old_interval = self.config.get("explorer_monitor_interval", 2.0)
            
            new_monitor = dlg.monitor_cb.isChecked()
            new_interval = dlg.interval_spinbox.value()
            
            # 更新配置
            self.config["enable_explorer_monitor"] = new_monitor
            self.config["explorer_monitor_interval"] = new_interval
            self.save_config()
            
            # 如果监听状态或间隔改变，重启监听
            if old_monitor != new_monitor or (new_monitor and old_interval != new_interval):
                if old_monitor:
                    self.stop_explorer_monitor()
                if new_monitor:
                    self.monitor_interval = new_interval
                    self.start_explorer_monitor()
                
                status = "已启用" if new_monitor else "已禁用"
                interval_info = f"（检查间隔: {new_interval}秒）" if new_monitor else ""
                QMessageBox.information(self, "设置已更新", f"Explorer窗口监听{status}{interval_info}")
        
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
        """保存固定标签页到config.json"""
        pinned_paths = []
        for i in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(i)
            if hasattr(tab, 'is_pinned') and tab.is_pinned:
                if hasattr(tab, 'current_path'):
                    pinned_paths.append(tab.current_path)
        
        # 更新config并保存
        self.config["pinned_tabs"] = pinned_paths
        self.save_config()
        
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
                        tab = FileExplorerTab(self, path, is_shell=is_shell)
                        tab.is_pinned = True
                        short = path[-12:] if len(path) > 12 else path
                        pin_prefix = "📌"
                        title = pin_prefix + short
                        self.tab_widget.addTab(tab, title)
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
        
        # 加载配置
        self.config = self.load_config()
        
        self.bookmark_manager = BookmarkManager()
        # 检查并自动添加常用书签
        self.ensure_default_bookmarks()
        
        # 窗口调整大小相关变量
        self.resizing = False
        self.resize_direction = None
        self.resize_margin = 10  # 边缘检测范围（像素），增加到10像素更容易触发
        self.cursor_overridden = False  # 通过QApplication是否已覆盖光标
        
        self.init_ui()

        # 使用应用图标作为窗口图标
        try:
            from PyQt5.QtWidgets import QApplication
            self.setWindowIcon(QApplication.windowIcon())
        except Exception:
            pass
    
    def load_config(self):
        """加载配置文件"""
        default_config = {
            "enable_explorer_monitor": True,  # 默认启用Explorer监听
            "pinned_tabs": []  # 默认没有固定标签页
        }
        
        try:
            if os.path.exists("config.json"):
                with open("config.json", "r", encoding="utf-8") as f:
                    config = json.load(f)
                    # 合并默认配置
                    for key, value in default_config.items():
                        if key not in config:
                            config[key] = value
                    return config
        except Exception as e:
            print(f"Failed to load config: {e}")
        
        return default_config
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save config: {e}")

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
        self.setGeometry(100, 100, 1200, 800)
        
        # 设置窗口最小尺寸，允许窗口缩小到很小
        self.setMinimumSize(400, 300)
        
        # 启用鼠标追踪，以便在边缘时显示调整大小光标
        self.setMouseTracking(True)
        
        # 隐藏默认标题栏
        self.setWindowFlags(Qt.FramelessWindowHint)
        
        # 创建主容器，蓝色背景作为边框
        main_container = QWidget()
        main_container.setStyleSheet("background: #2196F3;")
        main_container.setAttribute(Qt.WA_TransparentForMouseEvents)  # 让鼠标事件穿透到主窗口
        container_layout = QVBoxLayout(main_container)
        container_layout.setContentsMargins(4, 4, 4, 4)
        container_layout.setSpacing(0)
        
        # 创建内容容器，白色背景
        content_widget = QWidget()
        content_widget.setStyleSheet("background: white;")
        main_layout = QVBoxLayout(content_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 将内容容器添加到主容器
        container_layout.addWidget(content_widget)
        
        # 创建自定义标题栏
        self.create_custom_titlebar(main_layout)

        # 书签栏（菜单栏）
        menu_bar = self.menuBar()
        menu_bar.clear()
        menu_bar.setFixedHeight(28)  # 设置菜单栏高度
        # 设置菜单栏的大小策略，允许它被压缩
        menu_bar.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        menu_bar.setStyleSheet("""
            QMenuBar {
                background-color: #f5f5f5;
                border-bottom: 1px solid #ddd;
                padding: 2px;
            }
            QMenuBar::item {
                padding: 4px 8px;
                background: transparent;
                min-width: 0px;
            }
            QMenuBar::item:selected {
                background: #e0e0e0;
            }
            QMenuBar::item:pressed {
                background: #d0d0d0;
            }
        """)
        self.populate_bookmark_bar_menu()
        # 将菜单栏添加到主布局
        main_layout.addWidget(menu_bar)

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
        self.dir_tree.setMinimumWidth(150)
        # 移除最大宽度限制，允许自由拖动调整
        # self.dir_tree.setMaximumWidth(350)
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
        
        # 将分割器添加到主容器
        main_layout.addWidget(self.splitter)
        
        # 设置主容器为中心部件
        self.setCentralWidget(main_container)

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
        
        # 启动Explorer窗口监听
        self.start_explorer_monitor()

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
    
    def start_explorer_monitor(self):
        """启动Explorer窗口监听（优化版）"""
        # 检查配置是否启用
        if not self.config.get("enable_explorer_monitor", True):
            print("[Explorer Monitor] Monitoring disabled in config")
            return
        
        if not HAS_PYWIN:
            print("[Explorer Monitor] Windows API not available, monitoring disabled")
            return
        
        # 获取监听间隔配置（默认2秒，更轻量）
        self.monitor_interval = self.config.get("explorer_monitor_interval", 2.0)
        print(f"[Explorer Monitor] Will start monitoring in 3 seconds (interval: {self.monitor_interval}s)...")
        self.explorer_monitoring = False
        self.known_explorer_windows = set()  # 记录已知的Explorer窗口
        self.last_check_time = 0  # 上次检查时间
        
        # 延迟启动监听，确保主窗口完全初始化
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, self._start_monitor_thread)
    
    def _start_monitor_thread(self):
        """实际启动监听线程（延迟调用）"""
        try:
            self.monitor_our_window = int(self.winId())  # 记录我们自己的窗口句柄
            self.explorer_monitoring = True
            print("[Explorer Monitor] Starting Explorer window monitoring...")
            
            # 启动监听线程
            monitor_thread = threading.Thread(target=self._explorer_monitor_loop, daemon=True)
            monitor_thread.start()
        except Exception as e:
            print(f"[Explorer Monitor] Failed to start: {e}")
    
    def stop_explorer_monitor(self):
        """停止Explorer窗口监听"""
        self.explorer_monitoring = False
        print("[Explorer Monitor] Stopped")
    
    def _explorer_monitor_loop(self):
        """Explorer窗口监听循环（优化版 - 降低CPU占用）"""
        try:
            # 首先记录所有已存在的Explorer窗口
            def enum_windows_callback(hwnd, _):
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    # CabinetWClass: 标准Explorer窗口
                    # ExploreWClass: 另一种Explorer窗口类型（如通过"打开文件夹"打开的）
                    if class_name in ("CabinetWClass", "ExploreWClass"):
                        if win32gui.IsWindowVisible(hwnd):
                            self.known_explorer_windows.add(hwnd)
                except:
                    pass
                return True
            
            win32gui.EnumWindows(enum_windows_callback, None)
            print(f"[Explorer Monitor] Found {len(self.known_explorer_windows)} existing Explorer windows")
            print(f"[Explorer Monitor] Monitor interval: {self.monitor_interval}s (optimized for low CPU usage)")
            
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
                        # 调试：记录所有可见窗口的类名（仅在检测到新窗口时）
                        if win32gui.IsWindowVisible(hwnd):
                            title = win32gui.GetWindowText(hwnd)
                            # 记录可能相关的窗口信息
                            if title and len(title) > 0 and (':\\' in title or title.startswith('C:') or title.startswith('D:')):
                                print(f"[Explorer Monitor] Debug: Found window - Class: '{class_name}', Title: '{title}'")
                        
                        # CabinetWClass: 标准Explorer窗口
                        # ExploreWClass: 另一种Explorer窗口类型
                        if class_name in ("CabinetWClass", "ExploreWClass"):
                            if win32gui.IsWindowVisible(hwnd):
                                current_explorer_windows.add(hwnd)
                    except:
                        pass
                    return True
                
                win32gui.EnumWindows(check_windows_callback, None)
                
                # 找出新增的窗口
                new_windows = current_explorer_windows - self.known_explorer_windows
                
                # 优化：如果没有新窗口，直接跳过处理
                if not new_windows:
                    self.known_explorer_windows = current_explorer_windows
                    continue
                
                print(f"[Explorer Monitor] Detected {len(new_windows)} new Explorer window(s)")
                
                for hwnd in new_windows:
                    # 检查是否是我们自己的窗口（避免误捕获嵌入的Explorer控件）
                    try:
                        # 获取窗口标题
                        title = win32gui.GetWindowText(hwnd)
                        
                        # 只排除明确是我们应用的主窗口，不要误排除路径中包含TabEx的Explorer窗口
                        # 检查是否以"TabExplorer"开头（软件主窗口）或者窗口句柄是我们的主窗口
                        if title.startswith("TabExplorer"):
                            print(f"[Explorer Monitor] Skipping our main window: {title}")
                            continue
                        
                        # 获取窗口的父窗口，如果父窗口是我们的应用，则跳过
                        try:
                            parent = win32gui.GetParent(hwnd)
                            if parent == self.monitor_our_window:
                                print(f"[Explorer Monitor] Skipping child window")
                                continue
                        except:
                            pass
                        
                        print(f"[Explorer Monitor] New Explorer window detected: {hwnd} - {title}")
                        
                        # 尝试获取Explorer窗口的当前路径
                        path = self._get_explorer_path(hwnd)
                        
                        if path:
                            print(f"[Explorer Monitor] ✓ Path: {path}")
                            
                            # 在主线程中打开新标签页
                            self.open_path_signal.emit(path)
                            
                            # 优化：减少等待时间（从500ms到200ms）
                            time.sleep(0.2)
                            
                            # 关闭原Explorer窗口
                            try:
                                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                                print(f"[Explorer Monitor] ✓ Closed original Explorer (hwnd={hwnd})")
                            except Exception as e:
                                print(f"[Explorer Monitor] ✗ Failed to close: {e}")
                        else:
                            print(f"[Explorer Monitor] ✗ Could not get path from {hwnd}")
                    
                    except Exception as e:
                        print(f"[Explorer Monitor] Error processing window {hwnd}: {e}")
                
                # 更新已知窗口列表
                self.known_explorer_windows = current_explorer_windows
                
        except Exception as e:
            print(f"[Explorer Monitor] Monitor loop error: {e}")
            import traceback
            traceback.print_exc()
    
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
                                # 获取当前路径
                                location = window.LocationURL
                                
                                print(f"[Explorer Monitor] LocationURL: {location}")
                                
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
                                        print(f"[Explorer Monitor] Special shell path detected: {location}")
                                        
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
                                        print(f"[Explorer Monitor] Unknown CLSID, using default home path")
                                        return QDir.homePath()
                                    else:
                                        # 其他格式的路径
                                        return location
                                else:
                                    print(f"[Explorer Monitor] LocationURL is empty, trying alternative methods...")
                                    
                                    # 尝试获取 LocationName
                                    try:
                                        location_name = window.LocationName
                                        print(f"[Explorer Monitor] LocationName: {location_name}")
                                        
                                        # 根据位置名称推断路径
                                        if location_name in ['此电脑', 'This PC', 'My Computer']:
                                            return 'shell:MyComputerFolder'
                                        elif location_name in ['网络', 'Network']:
                                            return 'shell:NetworkPlacesFolder'
                                        elif location_name in ['回收站', 'Recycle Bin']:
                                            return 'shell:RecycleBinFolder'
                                    except:
                                        pass
                                    
                                    # 如果都失败了，返回用户主目录
                                    return QDir.homePath()
                        except Exception as e:
                            print(f"[Explorer Monitor] Error accessing window properties: {e}")
                            continue
                    
                    # 如果第一次没找到，等待一下再试
                    if attempt < 2:
                        time.sleep(0.2)
                        
                except Exception as e:
                    print(f"[Explorer Monitor] Attempt {attempt + 1} failed: {e}")
                    if attempt < 2:
                        time.sleep(0.2)
            
            return None
            
        except Exception as e:
            print(f"[Explorer Monitor] Error getting path: {e}")
            import traceback
            traceback.print_exc()
            return None

    def closeEvent(self, event):
        """窗口关闭时停止服务器和监听"""
        self.server_running = False
        if hasattr(self, 'server_socket'):
            try:
                self.server_socket.close()
            except:
                pass
        
        # 停止Explorer监听
        self.stop_explorer_monitor()
        
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
        # 仅显示书签内容，不在菜单栏添加“设置”或“书签管理”入口
    
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
        
        QMessageBox.information(
            self, 
            "设置已更新", 
            f"Explorer窗口监听已{'启用' if checked else '禁用'}\n{'新打开的文件管理器窗口将自动嵌入到标签页中' if checked else '新打开的文件管理器窗口将独立显示'}"
        )

    def show_bookmark_manager_dialog(self):
        dlg = BookmarkManagerDialog(self.bookmark_manager, self)
        dlg.exec_()

class SettingsDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.resize(500, 300)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 添加标题说明
        from PyQt5.QtWidgets import QDialogButtonBox, QLabel
        title_label = QLabel("应用设置")
        title_label.setStyleSheet("font-size: 14pt; font-weight: bold; color: #333;")
        layout.addWidget(title_label)
        
        # 添加分隔线
        from PyQt5.QtWidgets import QFrame
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)
        
        # 选项区域
        self.monitor_cb = QCheckBox("监听新Explorer窗口", self)
        self.monitor_cb.setChecked(config.get("enable_explorer_monitor", True))
        self.monitor_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        layout.addWidget(self.monitor_cb)
        
        # 监听间隔设置
        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel("监听间隔（秒）:"))
        from PyQt5.QtWidgets import QDoubleSpinBox
        self.interval_spinbox = QDoubleSpinBox()
        self.interval_spinbox.setRange(0.5, 10.0)
        self.interval_spinbox.setSingleStep(0.5)
        self.interval_spinbox.setValue(config.get("explorer_monitor_interval", 2.0))
        self.interval_spinbox.setToolTip("检查新Explorer窗口的时间间隔，更长的间隔降低CPU占用")
        interval_layout.addWidget(self.interval_spinbox)
        interval_layout.addWidget(QLabel("（推荐: 2.0秒，降低CPU占用）"))
        interval_layout.addStretch(1)
        layout.addLayout(interval_layout)
        
        # 提示信息
        tip_label = QLabel("💡 提示：\n• 较短间隔(0.5-1秒)响应更快但CPU占用高\n• 较长间隔(2-5秒)CPU占用低但响应稍慢\n• 推荐使用2秒平衡性能与响应")
        tip_label.setStyleSheet("QLabel { color: #666; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }")
        layout.addWidget(tip_label)
        
        # 添加弹性空间，将按钮推到底部
        layout.addStretch(1)
        
        # 按钮区域
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

# 书签管理对话框（初步框架，后续可扩展重命名/新建/删除等功能）
from PyQt5.QtWidgets import QDialog, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QPushButton, QHBoxLayout, QInputDialog, QMessageBox
class BookmarkManagerDialog(QDialog):

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
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
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
    # 创建并设置应用图标（用于任务栏与窗口）
    try:
        from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QIcon
        # 生成简单的蓝白主题图标，呼应应用配色
        pix = QPixmap(256, 256)
        pix.fill(Qt.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.Antialiasing)
        blue = QColor("#2196F3")
        white = QColor("white")
        # 外层蓝色圆角背景
        painter.setBrush(blue)
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(0, 0, 256, 256, 40, 40)
        # 内层白色圆角容器，形成蓝色边框效果
        margin = 28
        painter.setBrush(white)
        painter.drawRoundedRect(margin, margin, 256 - 2*margin, 256 - 2*margin, 24, 24)
        # 中央绘制 "TE" 字样（TabExplorer缩写）
        painter.setPen(blue)
        font = QFont()
        font.setBold(True)
        font.setPointSize(96)
        painter.setFont(font)
        painter.drawText(pix.rect(), Qt.AlignCenter, "TE")
        painter.end()
        icon = QIcon(pix)
        app.setWindowIcon(icon)
    except Exception as e:
        print(f"[Icon] Failed to create app icon: {e}")
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


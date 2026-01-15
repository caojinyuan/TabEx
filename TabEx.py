# 全局搜索缓存（LRU缓存，最多缓存50个搜索结果）
from collections import OrderedDict
import hashlib

# 性能优化配置常量
MAX_SEARCH_CACHE_SIZE = 50  # 搜索缓存最大数量
MAX_SEARCH_RESULTS = 1000000  # 单次搜索最大结果数
MAX_CLOSED_TABS_HISTORY = 20  # 关闭标签页历史最大数量（从10增加到20）
MAX_SEARCH_HISTORY = 30  # 搜索历史最大数量（从20增加到30）
MAX_NAVIGATION_HISTORY = 50  # 导航历史最大数量

# 大文件夹异步加载配置
LARGE_FOLDER_THRESHOLD = 1000  # 超过此数量文件视为大文件夹
FOLDER_CHECK_TIMEOUT = 500  # 文件夹检查超时时间(ms)
ASYNC_LOAD_ENABLED = True  # 是否启用异步加载

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

# 全局搜索缓存实例（使用配置常量）
_search_cache = SearchCache(max_size=MAX_SEARCH_CACHE_SIZE)

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
                show_toast(self, "路径错误", f"路径不存在: {local_path}", level="warning")

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

# 搜索对话框
from PyQt5.QtCore import pyqtSignal as _pyqtSignal
class SearchDialog(QDialog):    
    def __init__(self, search_path, parent=None, search_history=None):
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
        self.search_history = search_history or []  # 搜索历史列表
        
        # 检测Everything
        self.everything_path = detect_everything()
        debug_print(f"[Search] Everything detected: {self.everything_path}")
        
        # 线程安全的结果队列（限制大小防止内存溢出）
        import queue
        self.result_queue = queue.Queue(maxsize=1000)  # 最多缓存1000个待显示的结果
        self.ui_update_timer = None
        self.queue_overflow_count = 0  # 队列溢出计数
        
        # 结果限制配置（使用虚拟滚动优化，支持更多结果）
        self.max_results = 1000000  # 最多显示100万个结果（虚拟滚动优化）
        self.current_result_count = 0
        self.batch_insert_size = 500  # 批量插入大小
        
        layout = QVBoxLayout(self)
        
        # 搜索选项区域
        search_options = QHBoxLayout()
        search_options.setSpacing(5)  # 设置控件间距为5像素
        
        # 搜索关键词（改为QComboBox支持历史记录）
        search_label = QLabel("搜索:")
        search_label.setFixedWidth(40)  # 固定标签宽度
        search_options.addWidget(search_label)
        from PyQt5.QtWidgets import QComboBox
        self.search_input = QComboBox()
        self.search_input.setEditable(True)
        self.search_input.setInsertPolicy(QComboBox.NoInsert)  # 不自动插入新条目
        self.search_input.setMinimumWidth(300)  # 设置最小宽度300像素
        self.search_input.lineEdit().setPlaceholderText("输入搜索关键词...")
        self.search_input.lineEdit().returnPressed.connect(self.start_search)
        # 填充历史记录
        if self.search_history:
            self.search_input.addItems(self.search_history)
        search_options.addWidget(self.search_input, 1)  # 添加stretch factor，让搜索框可以拉伸
        
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
        
        # Everything搜索选项
        self.use_everything_cb = QCheckBox("使用 Everything (极速)")
        if self.everything_path:
            self.use_everything_cb.setChecked(True)  # 如果有Everything，默认启用
            self.use_everything_cb.setToolTip(f"使用Everything搜索引擎\n路径: {self.everything_path}\n只搜索文件名，速度极快")
        else:
            self.use_everything_cb.setEnabled(False)
            self.use_everything_cb.setToolTip("未检测到Everything，请从 https://www.voidtools.com/ 下载安装")
        self.use_everything_cb.stateChanged.connect(self.on_everything_toggled)
        type_options.addWidget(self.use_everything_cb)
        
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
        debug_print("[Search] 搜索缓存已清除")
    
    def start_search(self):
        keyword = self.search_input.currentText().strip()  # 改用currentText获取输入或选中的文本
        if not keyword:
            show_toast(self, "提示", "请输入搜索关键词", level="warning")
            return
        
        # 将搜索关键词添加到历史记录（通过主窗口）
        if self.main_window and hasattr(self.main_window, 'add_search_history'):
            self.main_window.add_search_history(keyword)
            # 更新下拉列表
            self.search_input.clear()
            if hasattr(self.main_window, 'search_history'):
                self.search_input.addItems(self.main_window.search_history)
            # 设置当前文本为刚刚搜索的关键词
            self.search_input.setCurrentText(keyword)
        
        if not self.search_filename_cb.isChecked() and not self.search_content_cb.isChecked():
            show_toast(self, "提示", "请至少选择一种搜索类型", level="warning")
            return
        
        # 获取并验证搜索路径
        search_path = self.path_input.text().strip()
        if not search_path:
            show_toast(self, "提示", "请输入搜索路径", level="warning")
            return
        
        # 检查路径是否存在
        if not os.path.exists(search_path):
            show_toast(self, "路径错误", f"路径不存在:\n{search_path}", level="warning")
            return
        
        # 检查是否是目录
        if not os.path.isdir(search_path):
            show_toast(self, "路径错误", f"路径不是文件夹:\n{search_path}", level="warning")
            return
        
        # 检查是否是特殊路径（不支持搜索）
        if search_path.startswith('shell:'):
            show_toast(self, "不支持", "不支持搜索特殊路径（shell:）", level="warning")
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
        use_everything = self.use_everything_cb.isChecked() if self.everything_path else False
        self.search_thread = threading.Thread(
            target=self.do_search,
            args=(keyword, self.search_filename_cb.isChecked(), self.search_content_cb.isChecked(), file_types, cache_key, use_everything)
        )
        self.search_thread.daemon = True
        self.search_thread.start()
    
    def stop_search(self):
        self.is_searching = False
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("已停止")
    
    def search_with_everything(self, keyword, search_path, file_types=""):
        """使用Everything进行搜索"""
        import subprocess
        
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
                    if line and os.path.exists(line):
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
    
    def do_search(self, keyword, search_filename, search_content, file_types="", cache_key=None, use_everything=False):
        # 如果使用Everything搜索
        if use_everything and self.everything_path:
            self.result_queue.put({'type': 'status', 'text': 'Using Everything搜索引擎...'})
            
            try:
                results = self.search_with_everything(keyword, self.search_path, file_types)
                
                if not self.is_searching:
                    return
                
                # 将Everything结果添加到显示队列
                for file_path in results:
                    if not self.is_searching:
                        break
                    
                    try:
                        self.add_search_result(file_path)
                    except:
                        pass
                
                # 搜索完成
                final_count = len(results)
                self.result_queue.put({
                    'type': 'complete',
                    'count': final_count,
                    'cache_key': cache_key,
                    'results': results[:1000] if cache_key else None  # 限制缓存大小
                })
                
            except Exception as e:
                self.result_queue.put({'type': 'error', 'text': f'Everything搜索错误: {str(e)}'})
            
            return
        
        # 原有的搜索逻辑
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
        """双击搜索结果，打开文件所在文件夹或文件夹本身，并选中文件"""
        # 从第一列获取存储的完整路径
        path_item = self.result_list.item(row, 0)
        if path_item:
            file_path = path_item.data(256)  # 获取存储的完整路径
            
            if os.path.exists(file_path):
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

import sys
import os
import json
import subprocess
import string
import time
import socket
import threading
import queue
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QListWidget, QLabel, QToolBar, QAction, QMenu, QInputDialog, QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView, QSizePolicy, QTreeView, QFileSystemModel, QSplitter, QProgressBar, QCompleter, QFrame)  # 添加QFrame
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import Qt, QDir, QUrl, pyqtSignal, pyqtSlot, Q_ARG, QObject, QSize, QFileSystemWatcher, QTimer, QThread, QMutex, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QMouseEvent, QCursor, QDrag
# from PyQt5.QtGui import QIcon  # unused

# 全局调试开关
_DEBUG_MODE = True
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


def set_explorer_monitor_debug(enabled):
    """设置 Explorer Monitor 日志开关"""
    global _EXPLORER_MONITOR_DEBUG
    _EXPLORER_MONITOR_DEBUG = enabled
    debug_print(f"[Config] Explorer Monitor debug output: {'enabled' if enabled else 'disabled'}")


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
    """以与主进程解耦的方式启动外部程序"""
    if os.name == 'nt':
        flags = _DETACHED_FLAGS | _NEW_PROCESS_GROUP | _BREAKAWAY_FROM_JOB | extra_creationflags
        return subprocess.Popen(cmd, cwd=cwd, creationflags=flags, close_fds=True)
    # 非 Windows 环境使用新 session，避免收到父进程信号
    return subprocess.Popen(cmd, cwd=cwd, start_new_session=True, close_fds=True)


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
    toast = ToastMessage(anchor, title, message, level=level, duration=duration)
    _active_toasts.append(toast)
    toast.show()

# ==================== 异步文件夹大小检查线程 ====================
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

def set_debug_mode(enabled):
    """设置全局调试模式"""
    global _DEBUG_MODE
    _DEBUG_MODE = enabled

def qt_message_handler(mode, context, message):
    """自定义 Qt 消息处理器，过滤 QAxBase 等不需要的警告"""
    # 只在调试模式下输出 Qt 警告
    if _DEBUG_MODE:
        # 如果是调试模式，输出所有消息
        print(f"Qt Message: {message}")
    else:
        # 非调试模式下，只输出严重错误（Critical 和 Fatal）
        from PyQt5.QtCore import QtCriticalMsg, QtFatalMsg
        if mode in (QtCriticalMsg, QtFatalMsg):
            print(f"Qt Error: {message}")
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
        print(f"Failed to setup Windows API monitoring: {e}")

# 面包屑导航路径栏
class BreadcrumbPathBar(QWidget):
    """类似Windows资源管理器的面包屑路径栏，支持点击层级跳转"""
    pathChanged = pyqtSignal(str)  # 当路径改变时发出信号
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_path = ""
        self.edit_mode = False
        # 绑定resize节流定时器，避免拖动卡顿
        try:
            from PyQt5.QtCore import QTimer
            self._resize_timer = QTimer(self)
            self._resize_timer.setSingleShot(True)
            self._resize_timer.setInterval(75)
            self._resize_timer.timeout.connect(self.update_breadcrumbs)
        except Exception:
            self._resize_timer = None
        self.init_ui()
    
    def init_ui(self):
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(2, 0, 2, 0)
        self.layout.setSpacing(0)
        
        # 路径编辑框（编辑模式时显示）
        self.path_edit = QLineEdit(self)
        path_edit_height = int(26 * getattr(self, 'parent_window', self).dpi_scale if hasattr(getattr(self, 'parent_window', self), 'dpi_scale') else 26)
        self.path_edit.setFixedHeight(path_edit_height)
        self.path_edit.setStyleSheet("QLineEdit { font-size: 10pt; padding: 2px; border: 1px solid #ccc; }")
        self.path_edit.hide()
        self.path_edit.returnPressed.connect(self.on_edit_finished)
        self.path_edit.editingFinished.connect(self.exit_edit_mode)
        
        # 设置路径自动补全
        self.setup_path_completer()
        
        # 面包屑容器（显示模式时显示）
        self.breadcrumb_widget = QWidget(self)
        self.breadcrumb_widget.setStyleSheet("QWidget { background: transparent; }")
        # 设置最小宽度为0，允许完全缩小
        self.breadcrumb_widget.setMinimumWidth(0)
        # 设置大小策略：宽度不强制扩展，允许压缩
        self.breadcrumb_widget.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        
        self.breadcrumb_layout = QHBoxLayout(self.breadcrumb_widget)
        self.breadcrumb_layout.setContentsMargins(0, 0, 0, 0)
        self.breadcrumb_layout.setSpacing(0)
        self.breadcrumb_layout.addStretch(1)
        
        self.layout.addWidget(self.breadcrumb_widget)
        self.layout.addWidget(self.path_edit)
        
        # 设置整体样式
        self.setStyleSheet("""
            BreadcrumbPathBar {
                background: transparent;
                border: none;
            }
        """)
        self.setFixedHeight(int(36 * getattr(self, 'parent_window', self).dpi_scale if hasattr(getattr(self, 'parent_window', self), 'dpi_scale') else 36))
        # 设置大小策略：宽度不强制扩展，高度固定；允许被压缩，避免影响左侧目录树
        self.setMinimumWidth(0)
        self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)

    def resizeEvent(self, event):
        """当路径栏宽度变化时，节流更新面包屑，减少重绘卡顿"""
        try:
            if getattr(self, '_resize_timer', None):
                # 重启定时器以合并连续的resize事件
                self._resize_timer.start()
            else:
                self.update_breadcrumbs()
        except Exception:
            pass
        super().resizeEvent(event)
    
    def setup_path_completer(self):
        """设置路径自动补全"""
        # 创建文件系统模型用于路径补全
        self.completer_model = QFileSystemModel()
        self.completer_model.setRootPath("")  # 设置为空以显示所有驱动器
        self.completer_model.setFilter(QDir.Drives | QDir.AllDirs | QDir.NoDotAndDotDot)
        
        # 创建补全器
        self.completer = QCompleter(self.completer_model, self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)  # 不区分大小写
        self.completer.setCompletionMode(QCompleter.PopupCompletion)  # 弹出补全列表
        self.completer.setMaxVisibleItems(10)  # 最多显示10个补全项
        
        # 设置补全器的弹出窗口样式
        popup = self.completer.popup()
        popup.setStyleSheet("""
            QListView {
                font-size: 11pt;
                border: 1px solid #ccc;
                background-color: white;
                selection-background-color: #0078d7;
                selection-color: white;
            }
        """)
        
        # 连接信号以动态更新补全
        self.path_edit.textChanged.connect(self.update_completer_prefix)
        
        # 将补全器设置到输入框
        self.path_edit.setCompleter(self.completer)
        
        debug_print("[PathCompleter] Path auto-completion initialized")
    
    def update_completer_prefix(self, text):
        """动态更新补全前缀"""
        if not text:
            return
        
        # 处理不同的路径格式
        if text.startswith('shell:') or text in ['cmd', '回收站', '此电脑', '我的电脑', '桌面', '网络']:
            # 特殊命令不需要补全
            return
        
        # 标准化路径分隔符
        normalized_path = text.replace('/', '\\') if os.name == 'nt' else text
        
        # 获取父目录路径
        parent_dir = os.path.dirname(normalized_path)
        
        # 如果有父目录，设置为补全的根路径
        if parent_dir and os.path.exists(parent_dir):
            self.completer_model.setRootPath(parent_dir)
            debug_print(f"[PathCompleter] Updated root path to: {parent_dir}")
    
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
        
        # 动态根据可用宽度决定显示方式
        # 先尝试完整显示，如果空间不足则逐步省略
        self._create_breadcrumb_widgets(parts)
    
    def _create_breadcrumb_widgets(self, parts):
        """创建面包屑widget，如果空间不足则显示省略版本"""
        # 获取可用宽度
        available_width = self.breadcrumb_widget.width() - 20  # 留一些边距
        
        # 如果宽度还没确定（初始化时），使用默认策略
        if available_width < 100:
            available_width = 800  # 假设一个合理的默认宽度
        
        # 计算完整路径需要的宽度
        from PyQt5.QtGui import QFontMetrics
        font_metrics = QFontMetrics(self.font())
        
        # 计算所有分段的总宽度
        total_width = 0
        for i, (name, _) in enumerate(parts):
            total_width += font_metrics.horizontalAdvance(name) + 20  # 20是padding
            if i < len(parts) - 1:
                total_width += font_metrics.horizontalAdvance(">") + 10  # 分隔符
        
        # 如果完整显示能放下，就完整显示
        if total_width <= available_width or len(parts) <= 3:
            visible_parts = parts
            show_ellipsis = False
            insert_ellipsis_at = -1
        else:
            # 空间不足，显示省略版本：前2级 + ... + 后面尽可能多的级
            num_start = 2
            
            # 计算前2级和省略号需要的宽度
            start_width = 0
            for i in range(min(num_start, len(parts))):
                start_width += font_metrics.horizontalAdvance(parts[i][0]) + 20
                start_width += font_metrics.horizontalAdvance(">") + 10  # 分隔符
            
            # 省略号的宽度
            ellipsis_width = font_metrics.horizontalAdvance("...") + 20
            ellipsis_width += font_metrics.horizontalAdvance(">") + 10  # 分隔符
            
            # 计算剩余可用宽度
            remaining_width = available_width - start_width - ellipsis_width
            
            # 从后往前计算能放下多少级
            num_end = 0
            end_width = 0
            for i in range(len(parts) - 1, num_start - 1, -1):
                segment_width = font_metrics.horizontalAdvance(parts[i][0]) + 20
                if i > num_start:  # 不是第一个后面的元素，需要加分隔符
                    segment_width += font_metrics.horizontalAdvance(">") + 10
                
                if end_width + segment_width <= remaining_width:
                    end_width += segment_width
                    num_end += 1
                else:
                    break
            
            # 至少显示最后1级
            if num_end == 0:
                num_end = 1
            
            if num_start + num_end >= len(parts):
                # 所有内容都能显示，不需要省略
                visible_parts = parts
                show_ellipsis = False
                insert_ellipsis_at = -1
            else:
                visible_parts = parts[:num_start] + parts[-num_end:]
                show_ellipsis = True
                insert_ellipsis_at = num_start
        
        # 创建面包屑标签
        widget_index = 0
        # 计算缩放后的最小宽度，避免完全压缩不可见
        try:
            from PyQt5.QtWidgets import QApplication
            dpi = QApplication.primaryScreen().logicalDotsPerInch()
            scale = dpi / 96.0
        except Exception:
            scale = 1.0
        min_label_w = int(40 * scale)
        min_sep_w = int(8 * scale)
        min_ellipsis_w = int(20 * scale)
        for i, (name, full_path) in enumerate(visible_parts):
            # 如果需要在这个位置插入省略号
            if show_ellipsis and i == insert_ellipsis_at:
                ellipsis = QLabel("...")
                ellipsis.setStyleSheet("QLabel { color: #888; font-size: 11pt; padding: 0; margin: 0 2px; }")
                # 允许在横向上被压缩，但保留最小可见宽度
                ellipsis.setMinimumWidth(min_ellipsis_w)
                ellipsis.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
                ellipsis.setToolTip(self.current_path)  # 鼠标悬停显示完整路径
                self.breadcrumb_layout.insertWidget(widget_index, ellipsis)
                widget_index += 1
                
                # 添加分隔符
                separator = QLabel(">")
                separator.setStyleSheet("QLabel { color: #888; font-size: 11pt; padding: 0; margin: 0 2px; }")
                separator.setMinimumWidth(min_sep_w)
                separator.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
                self.breadcrumb_layout.insertWidget(widget_index, separator)
                widget_index += 1
            
            # 创建可点击的标签
            label = ClickableLabel(name, full_path)
            # 允许横向压缩，但保留最小可见宽度；内部自动使用省略显示
            label.setMinimumWidth(min_label_w)
            label.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            label.clicked.connect(self.on_segment_clicked)
            self.breadcrumb_layout.insertWidget(widget_index, label)
            widget_index += 1
            
            # 添加分隔符（除了最后一个）
            if i < len(visible_parts) - 1:
                separator = QLabel(">")
                separator.setStyleSheet("QLabel { color: #888; font-size: 11pt; padding: 0; margin: 0 2px; }")
                separator.setMinimumWidth(min_sep_w)
                separator.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
                self.breadcrumb_layout.insertWidget(widget_index, separator)
                widget_index += 1
    
    def on_segment_clicked(self, path):
        """点击某个层级时触发"""
        self.current_path = path
        self.pathChanged.emit(path)
        # 不在这里立即更新面包屑，等待navigate_to完成后会自动调用set_path更新
        # self.update_breadcrumbs()  # 注释掉，避免竟态条件
    
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
            # 先发出信号，如果导航失败，在 FileExplorerTab 中会恢复路径
            self.pathChanged.emit(new_path)
        else:
            # 如果路径未改变或为空，直接退出编辑模式（保持当前路径）
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
        self.config_file = config_file
        self.bookmark_tree = self.load_bookmarks()
        # 优化：延迟保存机制，避免频繁写入磁盘
        self._save_timer = None
        self._pending_save = False

    def load_bookmarks(self):
        # 首先尝试加载主书签文件
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
                # 主文件损坏，尝试从备份恢复
                backup_file = self.config_file + ".bak"
                if os.path.exists(backup_file):
                    print(f"Attempting to restore bookmarks from backup: {backup_file}")
                    try:
                        with open(backup_file, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        # 恢复主文件
                        import shutil
                        shutil.copy2(backup_file, self.config_file)
                        print("Bookmarks restored from backup successfully")
                        if 'roots' in data:
                            return data['roots']
                        return data
                    except Exception as e2:
                        print(f"Failed to restore from backup: {e2}")
                return {}
        else:
            # 主文件不存在，检查是否有备份文件
            backup_file = self.config_file + ".bak"
            if os.path.exists(backup_file):
                print(f"Main bookmark file not found, restoring from backup: {backup_file}")
                try:
                    with open(backup_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    # 恢复主文件
                    import shutil
                    shutil.copy2(backup_file, self.config_file)
                    print("Bookmarks restored from backup successfully")
                    if 'roots' in data:
                        return data['roots']
                    return data
                except Exception as e:
                    print(f"Failed to restore from backup: {e}")
            else:
                print("No bookmark file or backup found, starting with empty bookmarks")
            return {}

    def save_bookmarks(self, immediate=False):
        # 优化：延迟保存，避免频繁操作时多次写入
        if immediate:
            # 立即保存
            try:
                # 先备份旧书签
                if os.path.exists(self.config_file):
                    import shutil
                    try:
                        shutil.copy2(self.config_file, self.config_file + ".bak")
                    except Exception as e:
                        print(f"Failed to backup bookmarks: {e}")
                
                # 保存新书签
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump({"roots": self.bookmark_tree}, f, ensure_ascii=False, indent=2)
                self._pending_save = False
            except Exception as e:
                print(f"Failed to save bookmarks: {e}")
                # 尝试从备份恢复
                backup_file = self.config_file + ".bak"
                if os.path.exists(backup_file):
                    try:
                        import shutil
                        shutil.copy2(backup_file, self.config_file)
                        print("Bookmarks restored from backup")
                    except Exception as e2:
                        print(f"Failed to restore bookmarks: {e2}")
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
            if self.main_window and hasattr(self.main_window, 'tab_widget'):
                # 因为 tab_widget 只包含占位符，需要在 content_stack 中查找索引
                idx = -1
                if hasattr(self.main_window, 'content_stack'):
                    idx = self.main_window.content_stack.indexOf(self)
                else:
                    idx = self.main_window.tab_widget.indexOf(self)
                
                if idx != -1:
                    self.main_window.tab_widget.setTabText(idx, title)
                    debug_print(f"DEBUG: Set tab {idx} text to '{title}'")

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
                    # 避免在导航过程中同步（减少干扰）
                    if getattr(self, '_navigating_folder', False):
                        return
                    
                    self.current_path = local_path
                    if hasattr(self, 'path_bar'):
                        self.path_bar.set_path(local_path)
                    self.update_tab_title()
                    # 只在非程序化导航时添加到历史记录
                    if not self._navigating_programmatically and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    # 同步左侧目录树（仅在非导航状态下）
                    if self.main_window and hasattr(self.main_window, 'expand_dir_tree_to_path'):
                        self.main_window.expand_dir_tree_to_path(local_path)
        except Exception:
            pass

    def setup_ui(self):
        from PyQt5.QtWidgets import QLabel
        # 设置FileExplorerTab背景为白色
        self.setStyleSheet("background: white;")
        self.setAutoFillBackground(True)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 面包屑路径栏
        self.path_bar = BreadcrumbPathBar(self)
        # 允许路径栏在横向上被压缩（不限制最小宽度），以避免影响splitter拖动
        try:
            self.path_bar.setMinimumWidth(0)
        except Exception:
            pass
        self.path_bar.pathChanged.connect(self.on_path_bar_changed)
        layout.addWidget(self.path_bar)
        
        # 添加负间距，让分隔线靠近路径栏
        layout.addSpacing(-10)
        
        # 添加分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #d0d0d0;")
        separator.setFixedHeight(1)
        layout.addWidget(separator)
        
        # 加载指示器（初始隐藏）
        self.loading_bar = QProgressBar(self)
        self.loading_bar.setMaximum(0)  # 不确定进度模式
        self.loading_bar.setTextVisible(True)
        self.loading_bar.setFormat("正在加载大文件夹...")
        self.loading_bar.setMaximumHeight(20)
        self.loading_bar.hide()
        layout.addWidget(self.loading_bar)

        # 嵌入Explorer控件 - 使用 Shell.Explorer（IE控件）
        # 注意：Shell.Explorer 不支持 TortoiseGit 图标覆盖层
        # ExplorerBrowser COM 接口过于复杂，无法通过 QAxWidget 简单集成
        self.explorer = QAxWidget(self)
        self.explorer.setControl("Shell.Explorer")
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
        self.status_bar = QLabel("就绪")
        self.status_bar.setFixedHeight(20)
        self.status_bar.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.status_bar.setStyleSheet(
            "QLabel { padding: 2px 8px; background: white; border-top: 1px solid #e0e0e0; font-size: 12px; color: #444; }"
        )
        layout.addWidget(self.status_bar)
        
        # 异步加载相关
        self.folder_checker = None  # 文件夹大小检查线程
        self.pending_navigation = None  # 待处理的导航请求
        
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
        
        # 初始设置路径栏（确保路径栏显示初始路径）
        if hasattr(self, 'path_bar'):
            self.path_bar.set_path(self.current_path)
        
        # 启动路径同步定时器
        self.start_path_sync_timer()

        # 启动状态栏定时更新（轮询选中状态与目录计数）
        self.status_timer = QTimer(self)
        self.status_timer.setInterval(800)  # 低频轮询，避免性能影响
        self.status_timer.timeout.connect(self.update_explorer_status)
        self.status_timer.start()
        self.update_explorer_status()
        
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
                        # 导航完成，延迟清除导航标志（等待双击检查完成，避免时序问题）
                        if hasattr(self, '_navigating_folder') and self._navigating_folder:
                            from PyQt5.QtCore import QTimer
                            # 延迟1500ms清除，确保700ms的双击最终确认已完成
                            QTimer.singleShot(1500, lambda: setattr(self, '_navigating_folder', False))
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

    def open_tortoisegit_log(self):
        """打开 TortoiseGit 日志查看器"""
        try:
            current_path = self.current_path
            if not current_path or not os.path.exists(current_path):
                show_toast(self, "提示", "当前路径无效", level="warning")
                return

            repo_root = self._find_git_root(current_path)
            if not repo_root:
                show_toast(self, "提示", "当前目录不是 Git 仓库，未找到 .git", level="warning")
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
                    "提示",
                    "未找到 TortoiseGit，请确认已安装 TortoiseGit\n下载地址: https://tortoisegit.org/download/",
                    level="warning",
                )
                return
            
            # 启动 TortoiseGit Log
            launch_detached([tortoisegit_exe, '/command:log', f'/path:{repo_root}'])
            debug_print(f"[TortoiseGit] Opened log for: {repo_root}")
            
        except Exception as e:
            show_toast(self, "错误", f"无法打开 TortoiseGit Log: {e}", level="error")
            debug_print(f"[TortoiseGit] Failed to open log: {e}")
    
    def open_tortoisegit_commit(self):
        """打开 TortoiseGit 提交窗口"""
        try:
            current_path = self.current_path
            if not current_path or not os.path.exists(current_path):
                show_toast(self, "提示", "当前路径无效", level="warning")
                return

            repo_root = self._find_git_root(current_path)
            if not repo_root:
                show_toast(self, "提示", "当前目录不是 Git 仓库，未找到 .git", level="warning")
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
                    "提示",
                    "未找到 TortoiseGit，请确认已安装 TortoiseGit\n下载地址: https://tortoisegit.org/download/",
                    level="warning",
                )
                return
            
            # 启动 TortoiseGit Commit
            launch_detached([tortoisegit_exe, '/command:commit', f'/path:{repo_root}'])
            debug_print(f"[TortoiseGit] Opened commit for: {repo_root}")
            
        except Exception as e:
            show_toast(self, "错误", f"无法打开 TortoiseGit Commit: {e}", level="error")
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

    def on_path_bar_changed(self, path):
        """处理面包屑路径栏的路径变化"""
        path = path.strip()
        # 处理cmd命令
        if path.lower() == 'cmd':
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    launch_detached(['cmd', '/K', 'cd', '/d', current_dir], extra_creationflags=subprocess.CREATE_NEW_CONSOLE)
                    # 恢复路径栏显示当前路径
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, "错误", "当前路径无效，无法打开命令行", level="error")
            except Exception as e:
                show_toast(self, "错误", f"无法打开命令行: {e}", level="error")
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
            show_toast(self, "路径错误", f"路径不存在: {path}", level="warning")
            # 恢复路径栏显示当前正确的路径
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
            sel = self.explorer.dynamicCall('SelectedItems()')
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
                elif event.type() == QEvent.MouseButtonDblClick:
                    # 取消所有之前待处理的双击检查
                    try:
                        for timer in getattr(self, '_pending_double_click_timers', []):
                            try:
                                timer.stop()
                            except Exception:
                                pass
                        self._pending_double_click_timers = []
                    except Exception:
                        self._pending_double_click_timers = []
                    
                    # 记录双击发生时的状态和鼠标位置
                    path_before = getattr(self, 'current_path', None)
                    selected_before = getattr(self, '_selected_before_click', None)
                    # 保存双击时的鼠标全局位置，用于后续hit-test
                    double_click_pos = QCursor.pos()
                    
                    # 立即检查当前选中项数量（双击时刻）
                    current_selection = self._get_selected_count_safe()
                    
                    debug_print(f"[DoubleClick] path_before='{path_before}', selected_before={selected_before}, current_selection={current_selection}")
                    
                    # 如果双击时有选中项，设置导航标志（防止误判为空白双击）
                    if current_selection is not None and int(current_selection) > 0:
                        self._navigating_folder = True
                        debug_print(f"[DoubleClick] Set navigation flag (has selection)")
                        # 暂停路径同步定时器，避免导航期间的干扰
                        if hasattr(self, '_path_sync_timer') and self._path_sync_timer:
                            self._path_sync_timer.stop()
                            debug_print(f"[DoubleClick] Path sync timer paused")

                    # 使用延迟检查，给足够时间让文件夹导航完成
                    def check_and_handle():
                        try:
                            # 优先检查导航标志
                            if getattr(self, '_navigating_folder', False):
                                debug_print(f"[DoubleClick] Navigation flag is set, skip go_up")
                                self._selected_before_click = None
                                return
                            
                            # 优先使用 native hit-test 检查双击位置是否有项目（最可靠）
                            hit_test_result = None
                            if HAS_PYWIN:
                                try:
                                    # 使用保存的双击位置进行检测
                                    hit_test_result = self._native_listview_hit_test(double_click_pos.x(), double_click_pos.y())
                                    if hit_test_result:
                                        debug_print(f"[DoubleClick] Hit test positive at double-click position, skip go_up")
                                        self._selected_before_click = None
                                        return
                                    else:
                                        debug_print(f"[DoubleClick] Hit test negative (no item at click position)")
                                except Exception as e:
                                    debug_print(f"[DoubleClick] Hit test exception: {e}")
                                    hit_test_result = None
                            
                            # 1. 检查是否点击前有选中项
                            if selected_before is not None and int(selected_before) > 0:
                                debug_print(f"[DoubleClick] Had selection before: {selected_before}, skip go_up")
                                self._selected_before_click = None
                                return
                            
                            # 2. 检查当前选中项数量
                            cnt = self._get_selected_count_safe()
                            
                            debug_print(f"[DoubleClick] SelectedItems count: {cnt}")
                            
                            if cnt is not None and int(cnt) > 0:
                                debug_print(f"[DoubleClick] Has selection: {cnt}, skip go_up")
                                self._selected_before_click = None
                                return
                            
                            # 3. 检查路径是否变化（进入了文件夹）
                            cur_path = getattr(self, 'current_path', None)
                            if path_before is not None and cur_path is not None and cur_path != path_before:
                                debug_print(f"[DoubleClick] Path changed: '{path_before}' -> '{cur_path}', skip go_up")
                                self._selected_before_click = None
                                return
                            
                            # 严格判断：只有在明确确认是空白区域时才执行 go_up
                            # 如果 SelectedItems 无法获取（cnt=None），检查其他信号
                            if cnt is None:
                                # 当 cnt=None 时，使用多重信号确认是否为空白双击：
                                # 1. hit-test 必须明确为 False（空白区域）
                                # 2. selected_before 必须为 None 或 0（双击前无选中）
                                # 3. 路径未改变（未进入文件夹）
                                cur_path_check = getattr(self, 'current_path', None)
                                if (hit_test_result is False and 
                                    (selected_before is None or selected_before == 0) and
                                    path_before is not None and cur_path_check == path_before):
                                    # 满足所有条件，执行最终确认后 go_up
                                    debug_print(f"[DoubleClick] cnt=None but all other signals confirm blank area, proceed to final check")
                                else:
                                    debug_print(f"[DoubleClick] Cannot determine selection state (cnt=None), insufficient signals for go_up")
                                    self._selected_before_click = None
                                    return
                            
                            # 如果 hit-test 明确返回 False (空白区域) 且 cnt=0 或满足上述fallback条件，执行最终确认后再 go_up
                            if hit_test_result is False and (cnt == 0 or cnt is None):
                                # 终次确认：延迟700ms确认路径仍未变化且无导航活动，再执行 go_up，避免误判
                                def final_confirm_and_go_up():
                                    try:
                                        # 首先检查导航标志
                                        if getattr(self, '_navigating_folder', False):
                                            debug_print(f"[DoubleClick] Final confirm: navigation flag is set, skip go_up")
                                            self._navigating_folder = False
                                            return
                                        
                                        cur_path3 = getattr(self, 'current_path', None)
                                        debug_print(f"[DoubleClick] Final confirm: path_before='{path_before}', cur_path='{cur_path3}'")
                                        
                                        # 检查路径是否变化（包括正在导航的情况）
                                        if path_before is not None and cur_path3 != path_before:
                                            debug_print(f"[DoubleClick] Final confirm: path changed to '{cur_path3}', skip go_up")
                                            return
                                        
                                        # 再次检查是否有选中项（可能导航过程中有选中）
                                        cnt_final = self._get_selected_count_safe()
                                        debug_print(f"[DoubleClick] Final confirm: selection count = {cnt_final}")
                                        if cnt_final is not None and int(cnt_final) > 0:
                                            debug_print(f"[DoubleClick] Final confirm: found selection ({cnt_final}), skip go_up")
                                            return
                                        
                                        # 再次执行 hit-test 以提高准确性（最关键的检查）
                                        r2 = hit_test_result
                                        if HAS_PYWIN:
                                            try:
                                                r2 = self._native_listview_hit_test(double_click_pos.x(), double_click_pos.y())
                                                debug_print(f"[DoubleClick] Final confirm: re-run hit-test result = {r2}")
                                            except Exception as e:
                                                debug_print(f"[DoubleClick] Final confirm: hit-test exception: {e}")
                                                pass
                                        
                                        # 只有当 hit-test 仍然确认是空白区域时才执行 go_up
                                        if r2 is False:
                                            debug_print(f"[DoubleClick] Execute go_up after final confirm (blank area)")
                                            self.go_up(force=True)
                                        elif r2 is True:
                                            debug_print(f"[DoubleClick] Final confirm: hit-test found item, skip go_up")
                                        else:
                                            debug_print(f"[DoubleClick] Final confirm: hit-test inconclusive (r2={r2}), skip go_up for safety")
                                    except Exception as e:
                                        debug_print(f"[DoubleClick] Final confirm exception: {e}")
                                        import traceback
                                        traceback.print_exc()
                                    finally:
                                        # 最终确认完成，清除导航标志（双保险）
                                        if hasattr(self, '_navigating_folder'):
                                            self._navigating_folder = False
                                        # 重启路径同步定时器
                                        if hasattr(self, '_path_sync_timer') and self._path_sync_timer and not self._path_sync_timer.isActive():
                                            self._path_sync_timer.start(1000)
                                            debug_print(f"[Navigation] Path sync timer restarted after navigation complete")
                                t2 = QTimer()
                                t2.setSingleShot(True)
                                t2.timeout.connect(final_confirm_and_go_up)
                                t2.start(700)  # 增加到 700ms 以确保导航有足够时间完成
                                if not hasattr(self, '_pending_double_click_timers'):
                                    self._pending_double_click_timers = []
                                self._pending_double_click_timers.append(t2)
                            else:
                                debug_print(f"[DoubleClick] Uncertain state (hit_test={hit_test_result}, cnt={cnt}), skip go_up for safety")
                            
                        except Exception as e:
                            debug_print(f"[DoubleClick] Exception in check_and_handle: {e}")
                        finally:
                            self._selected_before_click = None

                    # 使用 400ms 延迟，给更多时间让文件夹导航完成
                    timer = QTimer()
                    timer.setSingleShot(True)
                    timer.timeout.connect(check_and_handle)
                    timer.start(400)
                    if not hasattr(self, '_pending_double_click_timers'):
                        self._pending_double_click_timers = []
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

    def __init__(self, parent=None, path="", is_shell=False, select_file=None):
        super().__init__(parent)
        self.main_window = parent
        self.current_path = path if path else QDir.homePath()
        self.select_file = select_file  # 要选中的文件名
        # 浏览历史记录
        self.history = []
        self.history_index = -1
        # 标志：是否正在程序化导航（用于防止sync时重复添加历史）
        self._navigating_programmatically = False
        # 用于跟踪待处理的双击检查定时器
        self._pending_double_click_timers = []
        # 双击事件唯一ID，用于区分不同的双击操作
        self._double_click_id = 0
        
        # 文件系统监控（监控当前路径的变化）
        self.file_watcher = QFileSystemWatcher()
        self.file_watcher.directoryChanged.connect(self.on_directory_changed)
        # 延迟刷新定时器（避免频繁刷新）
        self.refresh_timer = QTimer()
        self.refresh_timer.setSingleShot(True)
        self.refresh_timer.timeout.connect(self.delayed_refresh)
        self.refresh_delay_ms = 500  # 500ms延迟
        # 防抖机制：记录最近处理的路径和时间，避免事件风暴
        self._last_watcher_event = {}  # {path: timestamp}
        self._watcher_debounce_ms = 1200  # 1.2s 内的重复事件会被忽略，进一步抑制风暴

        # 目录轮询刷新（已禁用，QFileSystemWatcher已足够，轮询会造成不必要的刷新）
        # self.dir_mtime = None
        # self.dir_poll_timer = QTimer()
        # self.dir_poll_timer.setInterval(2000)  # 2s 低频轮询，开销很小
        # self.dir_poll_timer.timeout.connect(self._poll_directory_changes)
        
        self.setup_ui()
        
        # 安装事件过滤器来处理快捷键（让Ctrl键能穿透到主窗口）
        self.installEventFilter(self)
        if hasattr(self, 'explorer'):
            self.explorer.installEventFilter(self)
        
        self.navigate_to(self.current_path, is_shell=is_shell)
        # 路径同步定时器已在setup_ui中启动，这里不需要重复启动
        
        # 如果指定了要选中的文件，延迟选中（等待导航完成）
        # 增加延迟时间确保文件夹完全加载
        if self.select_file:
            QTimer.singleShot(1500, lambda: self.select_file_in_explorer(self.select_file))

    # 移除重复的setup_ui，保留带路径栏的实现

    def get_selected_filenames(self):
        """获取选中的文件名列表（仅文件名，含后缀）"""
        filenames = []
        try:
            # 通过Document接口获取SelectedItems
            doc = self.explorer.querySubObject('Document')
            if doc:
                # 获取选中项集合
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
                                    debug_print(f"[get_selected_filenames] Found: {name}")
        except Exception as e:
            debug_print(f"[get_selected_filenames] Error: {e}")
        return filenames

    def navigate_to(self, path, is_shell=False, add_to_history=True, skip_async_check=False):
        debug_print(f"[navigate_to] To '{path}' (is_shell={is_shell}, skip_async={skip_async_check})")
        
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
            self.folder_checker.stop()
            self.folder_checker.wait(50)  # 减少等待时间从100ms到50ms
        
        # 支持本地路径和shell特殊路径
        if is_shell:
            self._hide_loading_indicator()
            self.explorer.dynamicCall("Navigate(const QString&)", path)
            self.current_path = path
            # shell 路径不做文件系统监控与轮询
            # self._update_dir_polling(None)  # 已禁用目录轮询
            if hasattr(self, 'path_bar'):
                self.path_bar.set_path(path)
            self.update_tab_title()
            # 添加到历史记录
            if add_to_history:
                self._add_to_history(path)
            self.update_explorer_status()
        elif os.path.exists(path):
            # 如果skip_async_check=True或禁用异步，直接导航
            if skip_async_check or not ASYNC_LOAD_ENABLED or not os.path.isdir(path):
                self._perform_navigation(path, add_to_history)
            else:
                # 异步检查文件夹大小
                self._check_folder_size_async(path, add_to_history)
        else:
            debug_print(f"[navigate_to] Path does not exist: {path}")
    
    def _check_folder_size_async(self, path, add_to_history):
        """异步检查文件夹大小并决定是否显示加载指示器"""
        # 显示加载指示器
        self._show_loading_indicator()
        
        # 创建并启动检查线程
        self.folder_checker = FolderSizeChecker(path, self)
        self.folder_checker.finished.connect(
            lambda p, count, is_large: self._on_folder_size_checked(p, count, is_large, add_to_history)
        )
        self.folder_checker.start()
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
        old_path = getattr(self, 'current_path', None)
        url = QDir.toNativeSeparators(path)
        
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
        except:
            # 回退到普通Navigate
            self.explorer.dynamicCall("Navigate(const QString&)", url)
        
        self.current_path = path

        # 不再启动目录轮询（QFileSystemWatcher已足够，轮询会造成不必要的刷新）
        # self._update_dir_polling(path)  # 已禁用

        # 更新状态栏
        self.update_explorer_status()
        
        # 更新文件系统监控（只监控真实文件系统路径）
        if hasattr(self, 'file_watcher'):
            # 移除旧路径的监控
            if old_path and os.path.exists(old_path) and os.path.isdir(old_path) and not old_path.startswith('shell:'):
                watched_dirs = self.file_watcher.directories()
                if old_path in watched_dirs:
                    self.file_watcher.removePath(old_path)
                    debug_print(f"[FileWatcher] Stopped watching: {old_path}")
            
            # 添加新路径的监控
            if os.path.isdir(path):
                if self.file_watcher.addPath(path):
                    debug_print(f"[FileWatcher] Now watching: {path}")
                else:
                    debug_print(f"[FileWatcher] Failed to watch: {path}")
        
        if hasattr(self, 'path_bar'):
            self.path_bar.set_path(path)
        self.update_tab_title()
        # 添加到历史记录
        if add_to_history:
            self._add_to_history(path)
        
        # 延迟3秒后允许自动刷新（避免刚导航完就因为文件监视器触发刷新）
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, lambda: setattr(self, '_suppress_auto_refresh', False))
    
    def _show_loading_indicator(self):
        """显示加载指示器"""
        if hasattr(self, 'loading_bar'):
            self.loading_bar.show()
            debug_print("[AsyncLoad] Loading indicator shown")
    
    def _hide_loading_indicator(self):
        """隐藏加载指示器"""
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
        """文件系统监控：目录内容发生变化（带防抖）"""
        # 防抖：检查是否在短时间内已处理过此路径
        import time
        current_time = time.time() * 1000  # 转为毫秒
        
        if hasattr(self, '_last_watcher_event'):
            last_time = self._last_watcher_event.get(path, 0)
            time_since_last = current_time - last_time

            # 如果已经有刷新定时器在跑且是当前路径，则直接忽略此次事件，避免重复排队
            if path == getattr(self, 'current_path', None) and self.refresh_timer.isActive():
                return

            if time_since_last < self._watcher_debounce_ms:
                # 忽略短时间内的重复事件，但不打印日志避免刷屏
                return

            self._last_watcher_event[path] = current_time
            # 清理旧记录（保留最近10个）
            if len(self._last_watcher_event) > 10:
                sorted_items = sorted(self._last_watcher_event.items(), key=lambda x: x[1], reverse=True)
                self._last_watcher_event = dict(sorted_items[:10])
        else:
            self._last_watcher_event = {path: current_time}
        
        debug_print(f"[FileWatcher] Directory changed: {path}")
        # 如果设置了抑制标志，不触发刷新
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[FileWatcher] Auto-refresh suppressed during navigation")
            return
        # 只在监控的是当前路径时才刷新
        if path == self.current_path:
            self._schedule_refresh(reason="watcher")

    def _schedule_refresh(self, reason="manual"):
        """统一的刷新调度，避免重复代码"""
        if not self.refresh_timer.isActive():
            debug_print(f"[AutoRefresh] Scheduling refresh in {self.refresh_delay_ms}ms (reason={reason})")
            self.refresh_timer.start(self.refresh_delay_ms)
        else:
            self.refresh_timer.stop()
            self.refresh_timer.start(self.refresh_delay_ms)
            debug_print(f"[AutoRefresh] Refresh timer restarted (reason={reason})")
    
    def delayed_refresh(self):
        """延迟刷新：避免频繁刷新"""
        debug_print(f"[FileWatcher] Auto-refreshing: {self.current_path}")
        if hasattr(self, 'explorer') and self.current_path:
            try:
                # 重新导航到当前路径以刷新
                is_shell = self.current_path.startswith('shell:')
                if is_shell:
                    self.explorer.dynamicCall('Navigate(const QString&)', self.current_path)
                else:
                    url = 'file:///' + self.current_path.replace('\\', '/')
                    self.explorer.dynamicCall('Navigate2(const QVariant&)', url)
                debug_print(f"[FileWatcher] Refresh completed")
            except Exception as e:
                debug_print(f"[FileWatcher] Refresh error: {e}")
        # 刷新后更新状态栏
        self.update_explorer_status()

    def on_dir_model_directory_loaded(self, path):
        """目录树模型加载完成时，若当前显示路径匹配，则触发一次刷新"""
        try:
            if not path or not self.current_path:
                return
            # QFileSystemModel 传入的 path 可能不带末尾分隔符，统一规范后比较
            norm_loaded = os.path.abspath(path)
            norm_current = os.path.abspath(self.current_path)
            if norm_loaded == norm_current:
                self._schedule_refresh(reason="dir_tree_loaded")
        except Exception as e:
            debug_print(f"[DirModel] directoryLoaded handler error: {e}")

    def _poll_directory_changes(self):
        """兜底轮询目录修改时间，解决部分环境下 watcher 不触发的问题"""
        # 如果设置了抑制标志，不触发刷新
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[DirPoll] Auto-refresh suppressed during navigation")
            return
            
        path = self.current_path
        if not path or not os.path.isdir(path):
            return
        current_mtime = self._get_dir_mtime(path)
        if current_mtime is None:
            return
        if self.dir_mtime is None:
            self.dir_mtime = current_mtime
            return
        if current_mtime != self.dir_mtime:
            debug_print(f"[DirPoll] Detected mtime change for {path}, scheduling refresh")
            self.dir_mtime = current_mtime
            self._schedule_refresh(reason="poll")

    def _update_dir_polling(self, path):
        """根据当前路径启动或停止兜底轮询"""
        if not hasattr(self, 'dir_poll_timer'):
            return
        if path and os.path.isdir(path):
            # 立即更新mtime，避免误判为变化（特别是刚导航到新目录时）
            self.dir_mtime = self._get_dir_mtime(path)
            debug_print(f"[DirPoll] Updated mtime for {path}: {self.dir_mtime}")
            if not self.dir_poll_timer.isActive():
                self.dir_poll_timer.start()
                debug_print(f"[DirPoll] Started polling {path}")
        else:
            if self.dir_poll_timer.isActive():
                self.dir_poll_timer.stop()
                debug_print(f"[DirPoll] Stopped polling")
            self.dir_mtime = None

    def _get_dir_mtime(self, path):
        """安全获取目录修改时间"""
        try:
            return os.stat(path).st_mtime
        except Exception:
            return None

    def update_explorer_status(self):
        """更新嵌入 Explorer 下方状态栏"""
        if not hasattr(self, 'status_bar'):
            return
        path = getattr(self, 'current_path', None)
        if not path:
            self.status_bar.setText("就绪")
            return

        # shell 路径直接展示提示
        if path.startswith('shell:') or '::' in path:
            self.status_bar.setText(f"{path}")
            return

        # 目录总项目数始终显示
        total = self._count_dir_entries(path)
        total_text = f"（共 {total} 项）" if total is not None else ""

        # 统计选中项
        selection = self._get_selection_entries()
        if selection is None:
            text = f"{total} 项" if total is not None else "就绪"
            self.status_bar.setText(text)
            return

        sel_count = len(selection)
        if sel_count == 0:
            text = f"{total} 项" if total is not None else "就绪"
            self.status_bar.setText(text)
            return

        # 选中状态
        text = f"已选 {sel_count} 项{total_text}"
        if all(entry['is_file'] for entry in selection):
            size_sum = sum(entry['size'] for entry in selection if entry['size'] is not None)
            text += f"，总大小 {self._format_size(size_sum)}"
        if sel_count == 1:
            entry = selection[0]
            mtime = self._get_mtime(entry['path'])
            if mtime:
                text += f"，修改时间 {mtime}"
        self.status_bar.setText(text)

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
            entries = []
            for i in range(int(count)):
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
                if is_file:
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

    def _count_dir_entries(self, path):
        """计算目录下的项目数"""
        try:
            with os.scandir(path) as it:
                return sum(1 for _ in it)
        except Exception:
            return None

    def _format_size(self, size):
        """格式化文件大小"""
        try:
            units = ['B', 'KB', 'MB', 'GB', 'TB']
            s = float(size)
            idx = 0
            while s >= 1024 and idx < len(units) - 1:
                s /= 1024
                idx += 1
            if idx == 0:
                return f"{int(s)} {units[idx]}"
            return f"{s:.2f} {units[idx]}"
        except Exception:
            return "-"

    def _get_mtime(self, path):
        """获取文件/文件夹修改时间字符串"""
        try:
            ts = os.path.getmtime(path)
            return self._format_mtime(ts)
        except Exception:
            return None

    def _format_mtime(self, ts):
        try:
            import datetime
            dt = datetime.datetime.fromtimestamp(ts)
            return dt.strftime('%Y-%m-%d %H:%M')
        except Exception:
            return None
    
    def select_file_in_explorer(self, filename):
        """在Explorer控件中选中指定的文件"""
        try:
            debug_print(f"[SelectFile] Attempting to select file: {filename}")
            
            # 构建完整路径
            full_path = os.path.join(self.current_path, filename)
            if not os.path.exists(full_path):
                debug_print(f"[SelectFile] File not found: {full_path}")
                return
            
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
                    return
                
                # 获取ListView中的项目数
                item_count = user32.SendMessageW(listview_hwnd, LVM_GETITEMCOUNT, 0, 0)
                debug_print(f"[SelectFile] ListView has {item_count} items")
                
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
                                            user32.SetFocus(listview_hwnd)
                                            
                                            debug_print(f"[SelectFile] Successfully selected file via API: {filename}")
                                            return
                    except Exception as e:
                        debug_print(f"[SelectFile] Error matching item {i}: {e}")
                        continue
                
                debug_print(f"[SelectFile] File not found in ListView: {filename}")
                
            except Exception as e:
                debug_print(f"[SelectFile] Windows API method failed: {e}")
                import traceback
                traceback.print_exc()
        
        except Exception as e:
            debug_print(f"[SelectFile] Error: {e}")
            import traceback
            traceback.print_exc()


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
                # 空白区域，打开新标签页
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    debug_print(f"[DEBUG] Opening new tab from TabBar blank area...")
                    self.main_window.add_new_tab()
                    return
        else:
            # 不在 TabBar 内，检查是否在标签页头部区域（TabBar 右侧的空白）
            # 获取 TabWidget 的 TabBar 所在的区域高度
            if event.pos().y() < tabbar.height():
                debug_print(f"[DEBUG] Click is in tab header area but outside TabBar")
                # 这是标签头和按钮之间的空白区域，打开新标签页
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    debug_print(f"[DEBUG] Opening new tab from header blank area...")
                    self.main_window.add_new_tab()
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



# 自定义委托，用于在目录树中显示箭头
from PyQt5.QtWidgets import QStyledItemDelegate
from PyQt5.QtGui import QPainter, QColor, QPen, QFont
from PyQt5.QtCore import Qt, QRect

class TreeItemDelegate(QStyledItemDelegate):
    """自定义目录树项委托，用于绘制展开/折叠箭头"""
    def paint(self, painter, option, index):
        # 先绘制默认内容
        super().paint(painter, option, index)
        
        # 如果有子项，绘制箭头
        if index.model().hasChildren(index):
            painter.save()
            
            # 设置字体和颜色
            font = QFont("Segoe UI", 10)
            font.setBold(True)
            painter.setFont(font)
            painter.setPen(QColor("#505050"))
            
            # 计算箭头位置（在项目左侧）
            arrow_rect = QRect(option.rect.left() - 16, option.rect.top(), 16, option.rect.height())
            
            # 根据展开状态选择箭头符号
            tree_view = option.widget
            if tree_view and isinstance(tree_view, QTreeView):
                if tree_view.isExpanded(index):
                    arrow = "˅"  # 展开 - 向下的尖角符号
                else:
                    arrow = "˃"  # 折叠 - 向右的尖角符号
                
                painter.drawText(arrow_rect, Qt.AlignCenter, arrow)
            
            painter.restore()

# 自定义支持拖拽的目录树
class DragDropTreeView(QTreeView):
    """支持拖拽打开文件夹的自定义QTreeView"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)  # 启用拖拽
        self.setDragDropMode(QTreeView.DragDrop)  # 支持拖入和拖出
        self.main_window = parent
    
    def dragEnterEvent(self, event):
        """目录树拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            debug_print("[DEBUG] DirTree: Drag enter accepted")
        else:
            event.ignore()
    
    def dragMoveEvent(self, event):
        """目录树拖拽移动事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def dropEvent(self, event):
        """目录树拖拽释放事件"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            debug_print(f"[DEBUG] DirTree: Drop event, urls count: {len(urls)}")
            
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
                    debug_print(f"[DEBUG] DirTree: Processing dropped path: {path}")
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
        self.hovered_tab = -1  # 当前鼠标悬停的标签页索引
        self.setMouseTracking(True)  # 启用鼠标追踪
        self.setMovable(True)  # 启用标签页拖拽排序
        # 连接标签移动信号
        self.tabMoved.connect(self.on_tab_moved)
    
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
            # 点击在空白区域，打开新标签页
            if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                debug_print(f"[DEBUG] Opening new tab from TabBar...")
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
                tab_close_btn_size = int(16 * getattr(self, 'parent_window', self).dpi_scale if hasattr(getattr(self, 'parent_window', self), 'dpi_scale') else 16)
                close_btn.setFixedSize(tab_close_btn_size, tab_close_btn_size)
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
        """标签页移动后的处理，确保固定标签页始终在左侧，并同步content_stack"""
        if not self.main_window:
            return
        
        debug_print(f"[TabMoved] Moving tab from {from_index} to {to_index}")
        
        # 同步移动content_stack中的对应内容
        if hasattr(self.main_window, 'content_stack'):
            content_stack = self.main_window.content_stack
            # 从content_stack中取出对应的widget
            moved_widget = content_stack.widget(from_index)
            if moved_widget:
                # 先移除
                content_stack.removeWidget(moved_widget)
                # 插入到新位置
                content_stack.insertWidget(to_index, moved_widget)
                debug_print(f"[TabMoved] Synced content_stack: moved widget from {from_index} to {to_index}")
        
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
    
    def get_tab_widget(self, index):
        """获取指定索引的实际标签页内容（从content_stack）"""
        if hasattr(self, 'content_stack') and index >= 0 and index < self.content_stack.count():
            return self.content_stack.widget(index)
        return self.tab_widget.widget(index)
    
    def get_current_tab_widget(self):
        """获取当前标签页的实际内容（从content_stack）"""
        current_index = self.tab_widget.currentIndex()
        return self.get_tab_widget(current_index)
    
    def _on_splitter_moved(self, pos, index):
        """当splitter移动时，保存目录树宽度（节流）"""
        if hasattr(self, 'splitter'):
            sizes = self.splitter.sizes()
            if len(sizes) >= 2:
                # 暂存宽度，使用定时器节流写入，减少频繁日志和重算
                self._pending_dir_tree_width = sizes[0]
                try:
                    if not hasattr(self, '_splitter_save_timer') or self._splitter_save_timer is None:
                        from PyQt5.QtCore import QTimer
                        self._splitter_save_timer = QTimer(self)
                        self._splitter_save_timer.setSingleShot(True)
                        self._splitter_save_timer.setInterval(100)
                        self._splitter_save_timer.timeout.connect(self._commit_splitter_size)
                    # 重启定时器以合并连续的拖动事件
                    self._splitter_save_timer.start()
                except Exception:
                    # 兜底：无定时器则直接保存
                    self._saved_dir_tree_width = self._pending_dir_tree_width
                    debug_print(f"[Splitter] Dir tree width saved: {self._saved_dir_tree_width}px")

    def _commit_splitter_size(self):
        """实际保存分割器左侧宽度并记录日志（被节流调用）"""
        try:
            if hasattr(self, '_pending_dir_tree_width'):
                self._saved_dir_tree_width = self._pending_dir_tree_width
                debug_print(f"[Splitter] Dir tree width saved: {self._saved_dir_tree_width}px")
        except Exception:
            pass


    def go_up_current_tab(self):
        current_tab = self.get_current_tab_widget()
        if hasattr(current_tab, 'go_up'):
            current_tab.go_up(force=True)
    
    def go_back_current_tab(self):
        """后退当前标签页"""
        current_tab = self.get_current_tab_widget()
        if current_tab and hasattr(current_tab, 'go_back'):
            current_tab.go_back()
    
    def go_forward_current_tab(self):
        """前进当前标签页"""
        current_tab = self.get_current_tab_widget()
        if current_tab and hasattr(current_tab, 'go_forward'):
            current_tab.go_forward()
    
    def update_navigation_buttons(self):
        """更新前进后退按钮状态"""
        current_tab = self.get_current_tab_widget()
        if current_tab and hasattr(current_tab, 'can_go_back'):
            self.back_button.setEnabled(current_tab.can_go_back())
        else:
            self.back_button.setEnabled(False)
        
        if current_tab and hasattr(current_tab, 'can_go_forward'):
            self.forward_button.setEnabled(current_tab.can_go_forward())
        else:
            self.forward_button.setEnabled(False)
    
    def open_tortoisegit_log_current_tab(self):
        """打开当前标签页的 TortoiseGit 日志"""
        current_tab = self.get_current_tab_widget()
        if current_tab and hasattr(current_tab, 'open_tortoisegit_log'):
            current_tab.open_tortoisegit_log()
    
    def open_tortoisegit_commit_current_tab(self):
        """打开当前标签页的 TortoiseGit 提交窗口"""
        current_tab = self.get_current_tab_widget()
        if current_tab and hasattr(current_tab, 'open_tortoisegit_commit'):
            current_tab.open_tortoisegit_commit()
    
    def apply_tortoisegit_buttons_config(self):
        """根据配置显示/隐藏 TortoiseGit 按钮"""
        enable = self.config.get("enable_tortoisegit_buttons", False)
        if hasattr(self, 'git_log_button'):
            self.git_log_button.setVisible(enable)
        if hasattr(self, 'git_commit_button'):
            self.git_commit_button.setVisible(enable)

    def _update_window_title(self, current_path: str = None):
        """根据当前状态更新窗口标题和自定义标题栏文本。
        当处于恢复状态时，在标题中附加“窗口恢复中”。
        """
        base_title = "TabExplorer"
        # 更新自定义标题栏文本
        if hasattr(self, 'title_label'):
            if self.is_restoring:
                self.title_label.setText(f"{base_title} - 窗口恢复中")
            else:
                self.title_label.setText(base_title)

        # 组合窗口标题
        if current_path:
            title = f"{base_title} - {current_path}"
        else:
            title = base_title
        if self.is_restoring:
            title = f"{title} - 窗口恢复中"
        try:
            self.setWindowTitle(title)
        except Exception:
            pass


    @pyqtSlot()
    @pyqtSlot(str)
    @pyqtSlot(str, bool)
    def add_new_tab(self, path="", is_shell=False, select_file=None):
        # 默认新建标签页为“此电脑”
        if not path:
            path = 'shell:MyComputerFolder'
            is_shell = True
        
        # 保存当前splitter尺寸
        saved_sizes = None
        if hasattr(self, 'splitter'):
            saved_sizes = self.splitter.sizes()
            if len(saved_sizes) >= 2:
                self._saved_dir_tree_width = saved_sizes[0]  # 更新保存的目录树宽度
        
        tab = FileExplorerTab(self, path, is_shell=is_shell, select_file=select_file)
        tab.is_pinned = False
        short = path[-16:] if len(path) > 16 else path
        
        # 同时添加到 tab_widget 和 content_stack
        tab_index = self.tab_widget.addTab(QWidget(), short)  # tab_widget 只显示标签，内容用占位widget
        self.content_stack.addWidget(tab)  # 实际内容添加到 content_stack
        
        self.tab_widget.setCurrentIndex(tab_index)
        
        # 强制恢复目录树宽度，防止长路径标签页影响布局
        if hasattr(self, '_saved_dir_tree_width') and hasattr(self, 'splitter'):
            from PyQt5.QtCore import QTimer
            # 延迟执行，确保布局已完成
            def restore_width():
                current_sizes = self.splitter.sizes()
                if len(current_sizes) >= 2:
                    total_width = sum(current_sizes)
                    self.splitter.setSizes([self._saved_dir_tree_width, total_width - self._saved_dir_tree_width])
                    debug_print(f"[Splitter] Restored dir tree width to {self._saved_dir_tree_width}px")
            QTimer.singleShot(0, restore_width)
        
        # 更新导航按钮状态（确保新标签页的按钮状态正确）
        self.update_navigation_buttons()
        
        # 激活窗口（当从其他实例接收到路径时）
        self.activateWindow()
        self.raise_()
        
        return tab_index


    def close_tab(self, index):
        tab = self.get_tab_widget(index)
        
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
                'title': self.tab_widget.tabText(index),
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
        if self.tab_widget.count() > 1:
            self.tab_widget.removeTab(index)
            # 同时从 content_stack 移除
            if hasattr(self, 'content_stack') and index < self.content_stack.count():
                widget = self.content_stack.widget(index)
                self.content_stack.removeWidget(widget)
                if widget:
                    widget.deleteLater()
        else:
            self.close()


    def close_current_tab(self):
        current_index = self.tab_widget.currentIndex()
        self.close_tab(current_index)
    
    def reopen_closed_tab(self):
        """恢复最近关闭的标签页"""
        if not self.closed_tabs_history:
            debug_print("[ClosedTabs] No closed tabs to restore")
            return
        
        # 取出最近关闭的标签页
        tab_info = self.closed_tabs_history.pop(0)
        debug_print(f"[ClosedTabs] Restoring tab: {tab_info['path']}, remaining history: {len(self.closed_tabs_history)}")
        
        # 重新打开标签页
        self.add_new_tab(tab_info['path'], is_shell=tab_info.get('is_shell', False))
        
        # 更新恢复按钮状态
        if hasattr(self, 'reopen_tab_button'):
            self.reopen_tab_button.setEnabled(len(self.closed_tabs_history) > 0)

    def on_tab_changed(self, index):
        if index >= 0:
            # 调试信息：检查同步状态
            print(f"[TabSwitch] Tab changed to index {index}")
            if hasattr(self, 'content_stack'):
                print(f"[TabSwitch] content_stack has {self.content_stack.count()} widgets, tab_widget has {self.tab_widget.count()} tabs")
            
            # 同步 content_stack 的显示
            if hasattr(self, 'content_stack') and index < self.content_stack.count():
                self.content_stack.setCurrentIndex(index)
                print(f"[TabSwitch] Set content_stack to index {index}")
            else:
                print(f"[TabSwitch] WARNING: Cannot sync - content_stack count is {self.content_stack.count() if hasattr(self, 'content_stack') else 'N/A'}")
            
            # 从 content_stack 获取实际的标签页内容
            tab = self.content_stack.widget(index) if hasattr(self, 'content_stack') else self.tab_widget.widget(index)
            if hasattr(tab, 'current_path'):
                # 统一通过内部方法更新窗口标题（可带“窗口恢复中”标记）
                self._update_window_title(tab.current_path)
                # 展开并选中左侧目录树到当前目录
                self.expand_dir_tree_to_path(tab.current_path)
            # 更新导航按钮状态
            self.update_navigation_buttons()
        
        # 更新所有标签页的字体：选中的加粗，未选中的正常
        from PyQt5.QtGui import QFont
        for i in range(self.tab_widget.count()):
            font = QFont()
            if i == index:
                # 选中的标签页：加粗
                font.setBold(True)
            else:
                # 未选中的标签页：正常
                font.setBold(False)
            self.tab_widget.tabBar().setTabTextColor(i, self.tab_widget.tabBar().tabTextColor(i))  # 保持颜色不变
            # 设置字体
            tab_bar = self.tab_widget.tabBar()
            # 通过样式表设置字体粗细
            if i == index:
                # 当前标签加粗
                current_text = self.tab_widget.tabText(i)
                self.tab_widget.setTabText(i, current_text)  # 触发重绘
            
        # 使用样式表设置选中标签的字体加粗（保持与初始样式一致）
        self.tab_widget.tabBar().setStyleSheet("""
            QTabBar::tab {
                border: 1px solid #b0b0b0;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 2px 4px;
                height: 24px;
                width: 120px;
                min-width: 120px;
                max-width: 120px;
                font-family: 'Microsoft YaHei UI', 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
                margin-top: 2px;
                text-align: center;
            }
            QTabBar::tab:selected {
                background: #FFF9CC;
                border: 1px solid #999;
                border-bottom: none;
                margin-top: 0px;
                padding-top: 3px;
            }
            QTabBar::tab:!selected {
                font-weight: normal;
            }
        """)
        # 设置标签文本省略模式 - 左边省略，保留右侧文件/文件夹名称
        self.tab_widget.tabBar().setElideMode(Qt.ElideLeft)

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
        """创建自定义标题栏，包含窗口控制按钮和功能按钮"""
        # 根据DPI调整标题栏高度
        titlebar_height = int(32 * getattr(self, 'dpi_scale', 1.0))
        titlebar = QWidget()
        titlebar.setFixedHeight(titlebar_height)
        titlebar.setStyleSheet("background-color: #f3f3f3;")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(10, 0, 0, 0)
        titlebar_layout.setSpacing(0)
        
        # 窗口标题 - 使用图标 + 文字的组合
        title_container = QWidget()
        title_layout = QHBoxLayout(title_container)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(6)
        
        # 应用图标
        icon_label = QLabel("📂")
        icon_label.setStyleSheet("font-size: 14pt; padding: 0; margin: 0;")
        title_layout.addWidget(icon_label)
        
        # 应用名称
        self.title_label = QLabel("Tab Explorer")
        title_font_size = int(12 * getattr(self, 'dpi_scale', 1.0))
        self.title_label.setStyleSheet(f"""
            font-family: 'Segoe UI Semibold', 'Segoe UI', 'Microsoft YaHei UI', sans-serif;
            font-weight: 600;
            font-size: 10.5pt;
            color: #303030;
            letter-spacing: 0.3px;
            padding: 0;
            margin: 0;
        """)
        title_layout.addWidget(self.title_label)
        title_layout.addStretch()
        
        titlebar_layout.addWidget(title_container)
        
        # 用于拖动窗口
        self.titlebar_widget = titlebar
        self.drag_position = None
        
        titlebar_layout.addStretch()
        
        # TortoiseGit 按钮（可在设置中启用/禁用）
        btn_size = int(32 * getattr(self, 'dpi_scale', 1.0))
        btn_font_size = int(14 * getattr(self, 'dpi_scale', 1.0))
        btn_radius = int(4 * getattr(self, 'dpi_scale', 1.0))
        
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
        
        # Git Log 按钮
        self.git_log_button = QPushButton("🐢")
        self.git_log_button.setToolTip("打开 TortoiseGit 日志")
        self.git_log_button.setFixedSize(btn_size, btn_size)
        self.git_log_button.setStyleSheet(git_btn_style)
        self.git_log_button.clicked.connect(self.open_tortoisegit_log_current_tab)
        titlebar_layout.addWidget(self.git_log_button)
        
        # Git Commit 按钮
        self.git_commit_button = QPushButton("📤")
        self.git_commit_button.setToolTip("打开 TortoiseGit 提交窗口")
        self.git_commit_button.setFixedSize(btn_size, btn_size)
        self.git_commit_button.setStyleSheet(git_btn_style)
        self.git_commit_button.clicked.connect(self.open_tortoisegit_commit_current_tab)
        titlebar_layout.addWidget(self.git_commit_button)
        
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
        self.back_button.setToolTip("后退 (Alt+←)")
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
        self.forward_button.setToolTip("前进 (Alt+→)")
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
        self.add_tab_button.setToolTip("新建标签页 (Ctrl+T)")
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
        self.reopen_tab_button.setToolTip("恢复关闭的标签页 (Ctrl+Shift+T)")
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
        self.search_button.setToolTip("搜索当前文件夹 (Ctrl+F)")
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
        
        # 添加竖杠分隔符
        separator = QLabel("|")
        sep_font_size = int(16 * getattr(self, 'dpi_scale', 1.0))
        sep_padding = int(8 * getattr(self, 'dpi_scale', 1.0))
        separator.setStyleSheet(f"color: #d0d0d0; font-size: {sep_font_size}pt; padding: 0px {sep_padding}px; font-weight: 300;")
        separator.setFixedHeight(titlebar_height)
        titlebar_layout.addWidget(separator)
        
        # 书签管理按钮
        bookmark_btn = QPushButton("★")
        bookmark_btn.setToolTip("书签管理")
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
        settings_btn.setToolTip("设置")
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
        
        # 最小化按钮
        min_btn = QPushButton("─")
        win_btn_width = int(45 * getattr(self, 'dpi_scale', 1.0))
        min_btn.setFixedSize(win_btn_width, titlebar_height)
        min_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: 4px;
                font-size: {btn_font_size}pt;
                font-weight: bold;
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
        min_btn.clicked.connect(self.showMinimized)
        titlebar_layout.addWidget(min_btn)
        
        # 最大化/还原按钮
        self.max_btn = QPushButton("☐")
        self.max_btn.setFixedSize(win_btn_width, titlebar_height)
        self.max_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: 4px;
                font-size: {btn_font_size}pt;
                font-weight: bold;
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
        self.max_btn.clicked.connect(self.toggle_maximize)
        titlebar_layout.addWidget(self.max_btn)
        
        # 关闭按钮
        close_btn = QPushButton("✖")
        close_btn.setFixedSize(win_btn_width, titlebar_height)
        close_btn_font_size = int(16 * getattr(self, 'dpi_scale', 1.0))
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: 4px;
                font-size: {close_btn_font_size}pt;
                color: #202020;
            }}
            QPushButton:hover {{
                background: #e81123;
                color: white;
            }}
            QPushButton:pressed {{
                background: #c50f1f;
                color: white;
            }}
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
            self.max_btn.setText("◱")
    
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
    
    def mousePressEvent(self, event):
        """鼠标按下事件 - 用于拖动窗口或调整大小"""
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
        
        if event.button() == Qt.LeftButton:
            # 检查点击位置是否在标题栏内（拖动窗口）- 优先检查以避免误触发边缘调整
            if hasattr(self, 'titlebar_widget'):
                titlebar_rect = self.titlebar_widget.geometry()
                if titlebar_rect.contains(event.pos()):
                    self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
                    event.accept()
                    return
            
            # 检测是否在边缘（调整大小）- 排除标题栏+标签栏区域
            titlebar_exclusion_height = int(70 * getattr(self, 'dpi_scale', 1.0))
            if event.pos().y() > titlebar_exclusion_height:  # 只在标题栏+标签栏下方才检测边缘
                edge = self.detect_edge(event.pos())
                if edge and not self.isMaximized():
                    self.resizing = True
                    self.resize_direction = edge
                    self.resize_start_pos = event.globalPos()
                    self.resize_start_geometry = self.geometry()
                    event.accept()
                    return
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
            if self.drag_position is not None:
                if self.isMaximized():
                    # 最大化状态下拖动：先恢复窗口，然后继续拖动
                    # 计算恢复后窗口的位置，使鼠标保持在标题栏的相对位置
                    mouse_global_pos = event.globalPos()
                    
                    # 恢复窗口
                    self.showNormal()
                    
                    # 计算新的拖动位置，使鼠标在窗口宽度的相对位置保持一致
                    # 假设鼠标在标题栏的相对位置是 drag_position 的 x 坐标
                    window_width = self.width()
                    # 将拖动位置调整为窗口中心附近，让拖动更自然
                    new_drag_x = window_width // 2
                    self.drag_position.setX(new_drag_x)
                    
                    # 移动窗口，使鼠标位置正确
                    new_pos = mouse_global_pos - self.drag_position
                    self.move(new_pos)
                else:
                    # 正常拖动
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
                else:
                    # 不在边缘，恢复默认光标
                    self.update_cursor(None)
            else:
                # 最大化状态下确保恢复默认光标
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
        margin = max(self.resize_margin, 15)
        
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

    def detect_edge_global(self, global_pos):
        """基于全局坐标检测边缘命中（包含阴影区域）"""
        margin = max(self.resize_margin, 15)
        frame_rect = self.frameGeometry()

        # 检测是否在标题栏+标签栏区域内，如果是则排除上边缘检测
        # 标题栏32px + 标签栏32px + 额外缓冲6px = 70px
        titlebar_exclusion_height = int(70 * getattr(self, 'dpi_scale', 1.0))
        local_pos = self.mapFromGlobal(global_pos)
        in_titlebar = (hasattr(self, 'titlebar_widget') and 
                      local_pos.y() >= 0 and 
                      local_pos.y() <= titlebar_exclusion_height and
                      local_pos.x() >= 0 and 
                      local_pos.x() <= self.width())

        left = global_pos.x() <= frame_rect.left() + margin
        right = global_pos.x() >= frame_rect.right() - margin
        top = global_pos.y() <= frame_rect.top() + margin and not in_titlebar  # 标题栏内不触发上边缘
        bottom = global_pos.y() >= frame_rect.bottom() - margin

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
    
    def nativeEvent(self, eventType, message):
        """处理 Windows 原生事件，实现任务栏点击切换"""
        try:
            if eventType == b"windows_generic_MSG" or eventType == "windows_generic_MSG":
                from ctypes import wintypes, cast, POINTER
                import ctypes
                
                # 解析消息结构
                msg = cast(int(message), POINTER(wintypes.MSG)).contents

                # WM_NCHITTEST - 让 Windows 自己处理无边框窗口的拖动调整大小
                if msg.message == 0x0084:  # WM_NCHITTEST
                    from PyQt5.QtCore import QPoint

                    # 最大化时不允许调整大小
                    if self.isMaximized():
                        return False, 0

                    # 从 lParam 解析全局坐标
                    x = ctypes.c_short(msg.lParam & 0xFFFF).value
                    y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value
                    global_pos = QPoint(x, y)

                    margin = max(self.resize_margin, 15)  # 阴影+命中容差

                    # 先用客户端坐标检测（无标题栏区域）
                    client_pos = self.mapFromGlobal(global_pos)
                    client_rect = self.rect()

                    left = client_pos.x() <= margin
                    right = client_pos.x() >= client_rect.width() - margin
                    top = client_pos.y() <= margin
                    bottom = client_pos.y() >= client_rect.height() - margin

                    # 如果坐标在客户端之外，使用 frameGeometry 兜底（阴影区域）
                    if not client_rect.contains(client_pos):
                        frame_rect = self.frameGeometry()
                        left = global_pos.x() <= frame_rect.left() + margin
                        right = global_pos.x() >= frame_rect.right() - margin
                        top = global_pos.y() <= frame_rect.top() + margin
                        bottom = global_pos.y() >= frame_rect.bottom() - margin

                    # 命中测试结果常量
                    HTLEFT = 10
                    HTRIGHT = 11
                    HTTOP = 12
                    HTTOPLEFT = 13
                    HTTOPRIGHT = 14
                    HTBOTTOM = 15
                    HTBOTTOMLEFT = 16
                    HTBOTTOMRIGHT = 17

                    if top and left:
                        return True, HTTOPLEFT
                    if top and right:
                        return True, HTTOPRIGHT
                    if bottom and left:
                        return True, HTBOTTOMLEFT
                    if bottom and right:
                        return True, HTBOTTOMRIGHT
                    if top:
                        return True, HTTOP
                    if bottom:
                        return True, HTBOTTOM
                    if left:
                        return True, HTLEFT
                    if right:
                        return True, HTRIGHT
                    # 非边缘区域，交回默认处理
                    return False, 0
                
                # WM_SYSCOMMAND = 0x0112
                if msg.message == 0x0112:  # WM_SYSCOMMAND
                    command = msg.wParam & 0xFFF0
                    
                    # SC_RESTORE = 0xF120 - 通过任务栏点击时发送
                    if command == 0xF120:  # SC_RESTORE
                        # 如果窗口当前可见（包括正常、最大化状态），则最小化
                        # 只有在最小化时，才执行默认的恢复行为
                        if not self.isMinimized():
                            self.showMinimized()
                            return True, 0
                        # 如果已经最小化，返回 False 让系统执行默认恢复

                # WM_SETCURSOR - 在命中边缘时主动切换调整大小光标
                if msg.message == 0x0020:  # WM_SETCURSOR
                    from PyQt5.QtCore import QPoint

                    if self.isMaximized():
                        return False, 0

                    # 获取当前全局鼠标位置
                    import ctypes
                    from ctypes import wintypes
                    pt = wintypes.POINT()
                    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
                    global_pos = QPoint(pt.x, pt.y)

                    margin = max(self.resize_margin, 15)

                    # 首先用客户端坐标检测
                    client_pos = self.mapFromGlobal(global_pos)
                    client_rect = self.rect()

                    left = client_pos.x() <= margin
                    right = client_pos.x() >= client_rect.width() - margin
                    top = client_pos.y() <= margin
                    bottom = client_pos.y() >= client_rect.height() - margin

                    # 客户端外部则使用 frame 几何（包含阴影）
                    if not client_rect.contains(client_pos):
                        frame_rect = self.frameGeometry()
                        left = global_pos.x() <= frame_rect.left() + margin
                        right = global_pos.x() >= frame_rect.right() - margin
                        top = global_pos.y() <= frame_rect.top() + margin
                        bottom = global_pos.y() >= frame_rect.bottom() - margin

                    edge = None
                    if top and left:
                        edge = 'top-left'
                    elif top and right:
                        edge = 'top-right'
                    elif bottom and left:
                        edge = 'bottom-left'
                    elif bottom and right:
                        edge = 'bottom-right'
                    elif top:
                        edge = 'top'
                    elif bottom:
                        edge = 'bottom'
                    elif left:
                        edge = 'left'
                    elif right:
                        edge = 'right'

                    # 主动设置光标；若无边缘命中，交给默认处理
                    if edge:
                        self.update_cursor(edge)
                        return True, 1  # 告诉系统已处理，防止重置光标
                    else:
                        self.update_cursor(None)
                        return False, 0
        except Exception:
            pass
        
        return super().nativeEvent(eventType, message)
    
    def changeEvent(self, event):
        """窗口状态变化事件"""
        if event.type() == event.WindowStateChange:
            # 保存最后的非最小化状态，用于任务栏点击判断
            if not (self.windowState() & Qt.WindowMinimized):
                self._last_active_state = self.windowState()
            
            # 根据窗口状态切换阴影显示
            self._update_window_shadow()
        
        super().changeEvent(event)
    
    def event(self, event):
        """处理窗口事件，实现任务栏点击切换"""
        if event.type() == event.WindowActivate:
            # 窗口激活事件 - 用于检测任务栏点击
            # 使用定时器短暂延迟检测，避免初始激活时触发
            if hasattr(self, '_activation_timer'):
                # 如果窗口已经激活，再次点击任务栏应该最小化
                if self.isActiveWindow() and not self.isMinimized():
                    from PyQt5.QtCore import QTimer
                    QTimer.singleShot(50, self._check_taskbar_click)
            else:
                # 首次激活，标记定时器已设置
                self._activation_timer = True
        
        return super().event(event)
    
    def _check_taskbar_click(self):
        """检查是否是任务栏点击（通过短时间内重复激活判断）"""
        # 这个方法可以进一步优化，目前主要依赖 nativeEvent 的实现
        pass
    
    def setup_shortcuts(self):
        """设置全局快捷键（现在使用轮询方式，不再使用QShortcut）"""
        # QShortcut 被 QAxWidget 拦截，所以现在使用定时器轮询方式
        # 保留此方法以便将来扩展或备用
        self.shortcuts = []
    def setup_shortcuts(self):
        """设置全局快捷键（现在使用轮询方式，不再使用QShortcut）"""
        # QShortcut 被 QAxWidget 拦截，所以现在使用定时器轮询方式
        # 保留此方法以便将来扩展或备用
        self.shortcuts = []
    
    def refresh_current_tab(self):
        """刷新当前标签页"""
        current_tab = self.get_current_tab_widget()
        if hasattr(current_tab, 'current_path'):
            current_tab.navigate_to(current_tab.current_path, 
                                  is_shell=current_tab.current_path.startswith('shell:'))
    
    def add_current_tab_bookmark(self):
        """添加当前标签页到书签"""
        current_tab = self.get_current_tab_widget()
        if current_tab:
            self.add_tab_bookmark(current_tab)

    def copy_selected_filename(self):
        """复制当前选中文件名（含后缀）到剪贴板并提示"""
        debug_print("[copy_selected_filename] Called")
        current_tab = self.get_current_tab_widget()
        debug_print(f"[copy_selected_filename] current_tab: {current_tab}")
        if not current_tab or not hasattr(current_tab, 'get_selected_filenames'):
            debug_print("[copy_selected_filename] No tab or no get_selected_filenames method")
            show_toast(self, "提示", "请先选择一个文件", level="warning")
            return
        names = current_tab.get_selected_filenames()
        debug_print(f"[copy_selected_filename] Selected names: {names}")
        if not names:
            debug_print("[copy_selected_filename] No files selected")
            show_toast(self, "提示", "请先选择一个文件", level="warning")
            return
        from PyQt5.QtWidgets import QApplication
        # 将所有选中的文件名用", "连接
        filenames_text = ", ".join(names)
        QApplication.clipboard().setText(filenames_text)
        debug_print(f"[copy_selected_filename] Copied to clipboard: {filenames_text}")
        # 根据文件数量显示不同的提示
        if len(names) == 1:
            show_toast(self, "复制成功", f"文件名: {names[0]}", level="info")
        else:
            show_toast(self, "复制成功", f"已复制 {len(names)} 个文件名", level="info")
    
    def keyPressEvent(self, event):
        """处理快捷键（备用方案，主要使用QShortcut）"""
        # 保留此方法以防QShortcut在某些情况下不工作
        super().keyPressEvent(event)
    
    def eventFilter(self, obj, event):
        """应用级别的事件过滤器（暂时不使用，因为被QAxWidget拦截）"""
        # 由于QAxWidget在底层拦截事件，eventFilter接收不到事件
        # 现在使用定时器轮询方式处理快捷键
        return super().eventFilter(obj, event)
    
    def _check_shortcuts(self):
        """定时检查快捷键状态（用于检测被QAxWidget拦截的快捷键）"""
        try:
            # 严格检查窗口是否激活 - 使用多重检查确保只在TabEx激活时响应
            # 1. 检查Qt窗口是否激活
            if not self.isActiveWindow():
                self._last_keys_state.clear()
                return
            
            # 2. 检查应用程序焦点窗口
            from PyQt5.QtWidgets import QApplication
            if QApplication.activeWindow() != self:
                self._last_keys_state.clear()
                return
            
            # 3. 使用Windows API检查前台窗口是否是当前窗口
            import ctypes
            from ctypes import wintypes
            
            foreground_hwnd = ctypes.windll.user32.GetForegroundWindow()
            current_hwnd = int(self.winId())
            
            # 如果前台窗口不是TabEx，则不响应快捷键
            if foreground_hwnd != current_hwnd:
                self._last_keys_state.clear()
                return
            
            # Windows虚拟键码
            VK_CONTROL = 0x11
            VK_SHIFT = 0x10
            VK_MENU = 0x12  # Alt键
            VK_F5 = 0x74
            
            # 获取键盘状态
            def is_key_pressed(vk_code):
                return ctypes.windll.user32.GetAsyncKeyState(vk_code) & 0x8000 != 0
            
            hotkeys = self.config.get("hotkeys", {})
            
            # 首先检查Ctrl+Alt组合键（优先级最高，避免被其他快捷键拦截）
            if is_key_pressed(VK_CONTROL) and is_key_pressed(VK_MENU) and hotkeys.get("copy_filename", True):
                key_combo = "Ctrl+Alt"
                if not self._last_keys_state.get(key_combo, False):
                    debug_print("[Shortcut Poll] Detected Ctrl+Alt")
                    self.copy_selected_filename()
                    self._last_keys_state[key_combo] = True
                return
            else:
                self._last_keys_state["Ctrl+Alt"] = False
            
            # 检查Ctrl组合键
            if is_key_pressed(VK_CONTROL):
                # Ctrl+Shift+T (0x54) - 恢复关闭的标签页（必须在Ctrl+T之前检测）
                if is_key_pressed(VK_SHIFT) and is_key_pressed(0x54) and hotkeys.get("reopen_tab", True):
                    key_combo = "Ctrl+Shift+T"
                    if not self._last_keys_state.get(key_combo, False):
                        print("[Shortcut Poll] Detected Ctrl+Shift+T")
                        self.reopen_closed_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+Shift+T"] = False
                
                # Ctrl+T (0x54) - 新建标签页（不包含Shift）
                if is_key_pressed(0x54) and not is_key_pressed(VK_SHIFT) and hotkeys.get("new_tab", True):
                    key_combo = "Ctrl+T"
                    if not self._last_keys_state.get(key_combo, False):
                        print("[Shortcut Poll] Detected Ctrl+T")
                        self.add_new_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+T"] = False
                
                # Ctrl+W (0x57)
                if is_key_pressed(0x57) and hotkeys.get("close_tab", True):
                    key_combo = "Ctrl+W"
                    if not self._last_keys_state.get(key_combo, False):
                        print("[Shortcut Poll] Detected Ctrl+W")
                        self.close_current_tab()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+W"] = False
                
                # Ctrl+F (0x46)
                if is_key_pressed(0x46) and hotkeys.get("search", True):
                    key_combo = "Ctrl+F"
                    if not self._last_keys_state.get(key_combo, False):
                        print("[Shortcut Poll] Detected Ctrl+F")
                        self.show_search_dialog()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+F"] = False
                
                # Ctrl+D (0x44)
                if is_key_pressed(0x44) and hotkeys.get("add_bookmark", True):
                    key_combo = "Ctrl+D"
                    if not self._last_keys_state.get(key_combo, False):
                        print("[Shortcut Poll] Detected Ctrl+D")
                        self.add_current_tab_bookmark()
                        self._last_keys_state[key_combo] = True
                    return
                else:
                    self._last_keys_state["Ctrl+D"] = False
                
                # Ctrl+Tab (0x09)
                if is_key_pressed(0x09) and hotkeys.get("switch_tab", True):
                    if is_key_pressed(VK_SHIFT):
                        # Ctrl+Shift+Tab
                        key_combo = "Ctrl+Shift+Tab"
                        if not self._last_keys_state.get(key_combo, False):
                            print("[Shortcut Poll] Detected Ctrl+Shift+Tab")
                            self.tab_widget.setCurrentIndex(
                                (self.tab_widget.currentIndex() - 1) % self.tab_widget.count())
                            self._last_keys_state[key_combo] = True
                        return
                    else:
                        # Ctrl+Tab
                        key_combo = "Ctrl+Tab"
                        if not self._last_keys_state.get(key_combo, False):
                            print("[Shortcut Poll] Detected Ctrl+Tab")
                            self.tab_widget.setCurrentIndex(
                                (self.tab_widget.currentIndex() + 1) % self.tab_widget.count())
                            self._last_keys_state[key_combo] = True
                        return
                else:
                    self._last_keys_state["Ctrl+Tab"] = False
                    self._last_keys_state["Ctrl+Shift+Tab"] = False
            
            # 检查Alt组合键
            if is_key_pressed(VK_MENU):
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
                
        except Exception as e:
            # 如果轮询出错，不影响程序运行
            pass
    
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
            self.config["debug_mode"] = dlg.debug_mode_cb.isChecked()
            self.config["explorer_monitor_interval"] = new_interval
            self.config["enable_cache_tabs"] = dlg.cache_tabs_cb.isChecked()
            
            # 更新全局调试开关
            set_debug_mode(self.config["debug_mode"])
            
            # 更新快捷键配置
            if "hotkeys" not in self.config:
                self.config["hotkeys"] = {}
            self.config["hotkeys"]["new_tab"] = dlg.hotkey_new_tab.isChecked()
            self.config["hotkeys"]["close_tab"] = dlg.hotkey_close_tab.isChecked()
            self.config["hotkeys"]["reopen_tab"] = dlg.hotkey_reopen_tab.isChecked()
            self.config["hotkeys"]["switch_tab"] = dlg.hotkey_switch_tab.isChecked()
            self.config["hotkeys"]["search"] = dlg.hotkey_search.isChecked()
            self.config["hotkeys"]["navigate"] = dlg.hotkey_navigate.isChecked()
            self.config["hotkeys"]["go_up"] = dlg.hotkey_go_up.isChecked()
            self.config["hotkeys"]["refresh"] = dlg.hotkey_refresh.isChecked()
            self.config["hotkeys"]["add_bookmark"] = dlg.hotkey_add_bookmark.isChecked()
            self.config["hotkeys"]["copy_filename"] = dlg.hotkey_copy_filename.isChecked()
            
            self.save_config()
            
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
            show_toast(self, "检查更新", "已在浏览器中打开更新页面", level="info")
        except Exception as e:
            show_toast(self, "错误", f"无法打开浏览器: {e}", level="error")
    
    def show_search_dialog(self):
        """显示搜索对话框（非模态）"""
        current_tab = self.get_current_tab_widget()
        if not current_tab or not hasattr(current_tab, 'current_path'):
            show_toast(self, "提示", "请先打开一个文件夹", level="warning")
            self.setFocus()
            return
        
        search_path = current_tab.current_path
        
        # 不支持搜索特殊路径
        if search_path.startswith('shell:'):
            show_toast(self, "提示", "不支持搜索特殊路径（shell:）", level="warning")
            self.setFocus()
            return
        
        if not os.path.exists(search_path):
            show_toast(self, "提示", f"路径不存在: {search_path}", level="warning")
            self.setFocus()
            return
        
        # 创建非模态对话框，传入搜索历史
        dlg = SearchDialog(search_path, self, self.search_history)
        # 保存对话框引用，防止被垃圾回收
        if not hasattr(self, 'search_dialogs'):
            self.search_dialogs = []
        self.search_dialogs.append(dlg)
        
        # 对话框关闭时从列表中移除
        dlg.finished.connect(lambda: self.search_dialogs.remove(dlg) if dlg in self.search_dialogs else None)
        
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
        
        debug_print(f"[Search History] Added '{keyword}', total: {len(self.search_history)}")

    def tab_context_menu(self, pos):
        tab_index = self.tab_widget.tabBar().tabAt(pos)
        if tab_index < 0:
            return
        tab = self.get_tab_widget(tab_index)
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
            show_toast(self, "无可用书签文件夹", "请先在 bookmarks.json 中添加至少一个文件夹。", level="warning")
            return
        # 选择父文件夹
        folder_names = [f"{name} (id:{fid})" for fid, name in folder_list]
        from PyQt5.QtWidgets import QInputDialog
        idx, ok = QInputDialog.getItem(self, "选择书签文件夹", "请选择父文件夹：", folder_names, 0, False)
        # 对话框关闭后，强制将焦点设回主窗口，防止QAxWidget拦截快捷键
        self.setFocus()
        if not ok:
            return
        folder_id = folder_list[folder_names.index(idx)][0]
        # 输入书签名称
        name, ok = QInputDialog.getText(self, "书签名称", "请输入书签名称：", text=os.path.basename(tab.current_path))
        # 对话框关闭后，强制将焦点设回主窗口
        self.setFocus()
        if not ok or not name:
            return
        # 保存到 bookmarks.json
        url = "file:///" + tab.current_path.replace("\\", "/")
        if bm.add_bookmark(folder_id, name, url):
            self.populate_bookmark_bar_menu()
        else:
            show_toast(self, "添加失败", "未能添加书签，请检查父文件夹。", level="warning")

    def pin_tab(self, tab_index):
        tab = self.get_tab_widget(tab_index)
        tab.is_pinned = True
        # 重新排序：所有固定的在最左侧
        self.sort_tabs_by_pinned()
        self.save_pinned_tabs()

    def unpin_tab(self, tab_index):
        tab = self.get_tab_widget(tab_index)
        tab.is_pinned = False
        self.sort_tabs_by_pinned()
        self.save_pinned_tabs()

    def sort_tabs_by_pinned(self):
        pinned = []
        unpinned = []
        # 记录当前tab对象
        current_index = self.tab_widget.currentIndex()
        current_tab = self.get_tab_widget(current_index) if current_index >= 0 else None
        for i in range(self.tab_widget.count()):
            tab = self.get_tab_widget(i)
            if hasattr(tab, 'is_pinned') and tab.is_pinned:
                pinned.append(tab)
            else:
                unpinned.append(tab)
        self.tab_widget.clear()
        if hasattr(self, 'content_stack'):
            # 清空 content_stack
            while self.content_stack.count() > 0:
                widget = self.content_stack.widget(0)
                self.content_stack.removeWidget(widget)
        new_tabs = pinned + unpinned
        for tab in new_tabs:
            # 先添加标签页（临时标题）- 占位widget
            self.tab_widget.addTab(QWidget(), "")
            # 将实际内容添加到 content_stack
            if hasattr(self, 'content_stack'):
                self.content_stack.addWidget(tab)
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
            tab = self.get_tab_widget(i)
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
        
        # 窗口调整大小相关变量
        self.resizing = False
        self.resize_direction = None
        self.resize_margin = 15  # 边缘检测范围（像素），与阴影边距一致以便更容易触发
        self.cursor_overridden = False  # 通过QApplication是否已覆盖光标
        
        # 搜索历史（内存中，软件关闭后自动清除）- 使用常量限制大小
        self.search_history = []
        self.max_search_history = MAX_SEARCH_HISTORY
        
        # 关闭标签页历史 - 使用常量限制大小
        self.closed_tabs_history = []  # 每项格式: {'path': str, 'title': str, 'is_shell': bool}
        self.max_closed_tabs_history = MAX_CLOSED_TABS_HISTORY

        # 窗口恢复状态标记（用于在标题上显示“窗口恢复中”）
        self.is_restoring = False
        self._restore_title_suffix = " - 窗口恢复中"
        
        # 性能优化：延迟初始化UI（先显示基本界面）
        self.init_ui()
        
        # 根据配置显示/隐藏 TortoiseGit 按钮
        self.apply_tortoisegit_buttons_config()
        
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
        self._shortcut_timer.start(50)  # 每50ms检查一次
        
        # 性能优化：延迟加载非关键功能（100ms后加载）
        QTimer.singleShot(100, self._delayed_initialization)

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
            "debug_mode": False,  # 默认关闭调试输出
            "explorer_monitor_debug": False,  # 默认关闭Explorer Monitor调试输出
            "pinned_tabs": [],  # 默认没有固定标签页
            "enable_cache_tabs": True,  # 默认启用缓存标签功能
            "cached_tabs": [],  # 缓存的非固定标签页
            "enable_tortoisegit_buttons": False,  # 默认关闭TortoiseGit按钮
            # 快捷键配置
            "hotkeys": {
                "new_tab": True,           # Ctrl+T
                "close_tab": True,         # Ctrl+W
                "reopen_tab": True,        # Ctrl+Shift+T
                "switch_tab": True,        # Ctrl+Tab / Ctrl+Shift+Tab
                "search": True,            # Ctrl+F
                "navigate": True,          # Alt+Left/Right
                "go_up": True,             # Alt+Up
                "refresh": True,           # F5
                "add_bookmark": True,      # Ctrl+D
                "copy_filename": True      # Ctrl+Alt - 复制选中文件名
            }
        }
        
        try:
            # 首先尝试加载主配置文件
            if os.path.exists("config.json"):
                with open("config.json", "r", encoding="utf-8") as f:
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
                    return config
            else:
                # 主配置文件不存在，检查是否有备份文件
                backup_file = "config.json.bak"
                if os.path.exists(backup_file):
                    print(f"Main config file not found, restoring from backup: {backup_file}")
                    try:
                        with open(backup_file, 'r', encoding='utf-8') as f:
                            config = json.load(f)
                        # 恢复主配置文件
                        import shutil
                        shutil.copy2(backup_file, "config.json")
                        print("Config restored from backup successfully")
                        # 合并默认配置
                        for key, value in default_config.items():
                            if key not in config:
                                config[key] = value
                        if "hotkeys" in config:
                            for key, value in default_config["hotkeys"].items():
                                if key not in config["hotkeys"]:
                                    config["hotkeys"][key] = value
                        return config
                    except Exception as e:
                        print(f"Failed to restore from backup: {e}")
                else:
                    print("No config file or backup found, starting with default config")
        except Exception as e:
            print(f"Failed to load config: {e}")
            # 主文件损坏，尝试从备份恢复
            backup_file = "config.json.bak"
            if os.path.exists(backup_file):
                print(f"Attempting to restore config from backup: {backup_file}")
                try:
                    with open(backup_file, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    # 恢复主文件
                    import shutil
                    shutil.copy2(backup_file, "config.json")
                    print("Config restored from backup successfully")
                    # 合并默认配置
                    for key, value in default_config.items():
                        if key not in config:
                            config[key] = value
                    if "hotkeys" in config:
                        for key, value in default_config["hotkeys"].items():
                            if key not in config["hotkeys"]:
                                config["hotkeys"][key] = value
                    return config
                except Exception as e2:
                    print(f"Failed to restore from backup: {e2}")
        
        return default_config
    
    def save_config(self):
        """保存配置文件"""
        try:
            # 保存新配置
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            
            # 保存成功后创建备份
            import shutil
            try:
                shutil.copy2("config.json", "config.json.bak")
            except Exception as e:
                print(f"Failed to backup config: {e}")
                
        except Exception as e:
            print(f"Failed to save config: {e}")
            # 尝试从备份恢复
            if os.path.exists("config.json.bak"):
                try:
                    import shutil
                    shutil.copy2("config.json.bak", "config.json")
                    print("Config restored from backup")
                except Exception as e2:
                    print(f"Failed to restore config: {e2}")

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
            ]
            bm.save_bookmarks()

    def _setup_window_shadow(self):
        """设置窗口阴影效果"""
        # 初始化阴影状态标志
        self._shadow_enabled = False
    
    def _update_window_shadow(self):
        """根据窗口状态更新阴影显示 - 使用 Windows DWM 原生阴影"""
        # 只在窗口化状态（非最大化、非最小化、非全屏）时显示阴影
        is_normal = not (self.windowState() & (Qt.WindowMaximized | Qt.WindowMinimized | Qt.WindowFullScreen))
        
        if is_normal and not self._shadow_enabled:
            # 需要显示阴影且当前未启用，使用 Windows DWM 阴影
            if HAS_PYWIN:
                try:
                    import win32gui
                    import win32con
                    hwnd = int(self.winId())
                    import ctypes
                    from ctypes import wintypes
                    dwmapi = ctypes.windll.dwmapi
                    
                    # 启用窗口阴影
                    DWMWA_NCRENDERING_POLICY = 2
                    DWMNCRP_ENABLED = 2
                    dwmapi.DwmSetWindowAttribute(
                        hwnd,
                        DWMWA_NCRENDERING_POLICY,
                        ctypes.byref(wintypes.DWORD(DWMNCRP_ENABLED)),
                        ctypes.sizeof(wintypes.DWORD)
                    )
                    
                    # 不扩展玻璃框架到客户区，避免产生半透明缝隙
                    # WS_THICKFRAME 本身会提供系统阴影，无需调用 DwmExtendFrameIntoClientArea
                    
                    # 设置窗口圆角（Windows 11 风格）
                    try:
                        DWM_WINDOW_CORNER_PREFERENCE = 33
                        DWMWCP_ROUND = 2  # 圆角
                        dwmapi.DwmSetWindowAttribute(
                            hwnd,
                            DWM_WINDOW_CORNER_PREFERENCE,
                            ctypes.byref(wintypes.DWORD(DWMWCP_ROUND)),
                            ctypes.sizeof(wintypes.DWORD)
                        )
                    except:
                        pass  # Windows 10 不支持此属性，忽略错误
                    
                    # 添加可见边框和可调整大小的样式（WS_THICKFRAME/WS_SIZEBOX）
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    # WS_THICKFRAME 允许通过命中测试调整大小
                    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style | win32con.WS_BORDER | win32con.WS_THICKFRAME)
                    
                    self._shadow_enabled = True
                    debug_print("[Shadow] Windows shadow enabled with enhanced border")
                except Exception as e:
                    debug_print(f"[Shadow] Failed to enable Windows shadow: {e}")
        elif not is_normal and self._shadow_enabled:
            # 最大化时移除阴影和边框
            if HAS_PYWIN:
                try:
                    import win32gui
                    import win32con
                    hwnd = int(self.winId())
                    import ctypes
                    from ctypes import wintypes
                    dwmapi = ctypes.windll.dwmapi
                    
                    # 禁用窗口阴影
                    DWMWA_NCRENDERING_POLICY = 2
                    DWMNCRP_DISABLED = 1
                    dwmapi.DwmSetWindowAttribute(
                        hwnd,
                        DWMWA_NCRENDERING_POLICY,
                        ctypes.byref(wintypes.DWORD(DWMNCRP_DISABLED)),
                        ctypes.sizeof(wintypes.DWORD)
                    )
                    
                    # 移除边框和可调整大小样式
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style & ~win32con.WS_BORDER & ~win32con.WS_THICKFRAME)
                    
                    self._shadow_enabled = False
                    debug_print("[Shadow] Windows shadow disabled")
                except Exception as e:
                    debug_print(f"[Shadow] Failed to disable Windows shadow: {e}")

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
        
        # 启用鼠标追踪，以便在边缘时显示调整大小光标
        self.setMouseTracking(True)
        
        # 隐藏默认标题栏
        self.setWindowFlags(Qt.FramelessWindowHint)

        # 填充窗口背景，避免边框与内容之间出现半透明缝隙
        self.setAutoFillBackground(True)
        from PyQt5.QtGui import QPalette, QColor
        pal = self.palette()
        pal.setColor(QPalette.Window, QColor(255, 255, 255))
        self.setPalette(pal)
        # 再次用样式表确保非客户区也以白色填充
        self.setStyleSheet("QMainWindow { background: white; }")
        
        # 添加窗口阴影效果（仅在非最大化时显示）
        self._setup_window_shadow()
        
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
                margin-right: 2px;
                text-align: center;
                color: #505050;
            }}
            QTabBar::tab:hover:!selected {{
                background: #e8e8e8;
                border-color: #c0c0c0;
            }}
            QTabBar::tab:selected {{
                background: #ffffff;
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
        
        # 将标签栏容器添加到主布局
        main_layout.addWidget(tab_bar_container)
        
        # 右键标签页支持固定/取消固定
        tabbar.setContextMenuPolicy(Qt.CustomContextMenu)
        tabbar.customContextMenuRequested.connect(self.tab_context_menu)

        # 书签栏（使用自定义菜单栏）
        self.menu_bar = CustomMenuBar(self)
        menu_bar_height = int(28 * getattr(self, 'dpi_scale', 1.0))
        self.menu_bar.setFixedHeight(menu_bar_height)  # 设置菜单栏高度
        # 设置菜单栏的大小策略，允许它被压缩
        self.menu_bar.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self.menu_bar.setStyleSheet("""
            QMenuBar {
                background-color: #f3f3f3;
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
        self.splitter = QSplitter()
        self.splitter.setOrientation(Qt.Horizontal)
        # 设置分割条宽度（必须在设置样式之前）
        self.splitter.setHandleWidth(5)
        # 允许左侧目录树折叠（往左拖动时隐藏）
        self.splitter.setCollapsible(0, True)  # 索引0是左侧目录树，允许折叠
        self.splitter.setCollapsible(1, False)  # 索引1是右侧标签页，不允许折叠
        # 设置分割条样式
        self.splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: transparent;
            }
            QSplitter::handle:hover {
                background-color: #e0e0e0;
            }
            QSplitter::handle:pressed {
                background-color: #d0d0d0;
            }
        """)
        # 设置子控件的拉伸因子（左侧0，右侧1，右侧会占据剩余空间）
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        
        # 保存目录树宽度，用于在添加新标签页时强制保持
        self._saved_dir_tree_width = 240  # 初始宽度 (约占总宽度的 20%)
        
        # 监听splitter移动事件，保存用户设置的目录树宽度
        self.splitter.splitterMoved.connect(self._on_splitter_moved)

        # 左侧目录树（应用性能优化）
        self.dir_model = QFileSystemModel()
        # 性能优化：启用延迟加载和过滤器
        self.dir_model.setFilter(QDir.Dirs | QDir.NoDotAndDotDot)  # 只显示文件夹
        # 当目录模型加载完成时，如果正显示的路径有更新，主动刷新右侧Explorer
        self.dir_model.directoryLoaded.connect(self.on_dir_model_directory_loaded)
        # 设置根为计算机根目录（显示所有盘符）
        root_path = QDir.rootPath()  # 通常为C:/
        # 性能优化：在后台线程中加载文件系统
        # setRootPath 本身是异步的，会在后台线程中填充
        self.dir_model.setRootPath(root_path)
        self.dir_tree = DragDropTreeView(self)
        self.dir_tree.setModel(self.dir_model)
        # 性能优化：统一排序，减少渲染开销
        self.dir_tree.setSortingEnabled(True)
        self.dir_tree.sortByColumn(0, Qt.AscendingOrder)
        # 设置缩进，为箭头留出空间
        self.dir_tree.setIndentation(16)
        # 不设置setRootIndex，或者设置为index("")，这样能显示所有盘符
        # self.dir_tree.setRootIndex(self.dir_model.index(""))
        self.dir_tree.setHeaderHidden(True)
        self.dir_tree.setColumnHidden(1, True)
        self.dir_tree.setColumnHidden(2, True)
        self.dir_tree.setColumnHidden(3, True)
        
        # 设置自定义委托以显示箭头
        self.dir_tree.setItemDelegate(TreeItemDelegate(self.dir_tree))
        
        # 设置目录树样式
        self.dir_tree.setStyleSheet("""
            QTreeView {
                background-color: #f9f9f9;
                border: none;
                border-right: 1px solid #e5e5e5;
                font-family: 'Microsoft YaHei UI', 'Segoe UI', Arial, sans-serif;
                font-size: 9pt;
                outline: none;
            }
            QTreeView::item {
                padding: 3px 6px;
                border: none;
                height: 24px;
            }
            QTreeView::item:hover {
                background-color: #f3f3f3;
                border-radius: 4px;
            }
            QTreeView::item:selected {
                background-color: #e3f2fd;
                color: #000;
                border-radius: 4px;
            }
            QTreeView::item:selected:active {
                background-color: #bbdefb;
                border-radius: 4px;
            }
            QTreeView::branch:has-siblings:!adjoins-item {
                border-image: none;
            }
            QTreeView::branch:has-siblings:adjoins-item {
                border-image: none;
            }
            QTreeView::branch:!has-children:!has-siblings:adjoins-item {
                border-image: none;
            }
            QTreeView::branch:has-children:!has-siblings:closed,
            QTreeView::branch:closed:has-children:has-siblings {
                border-image: none;
                image: none;
                background: transparent;
            }
            QTreeView::branch:open:has-children:!has-siblings,
            QTreeView::branch:open:has-children:has-siblings {
                border-image: none;
                image: none;
                background: transparent;
            }
        """)
        # 移除最小宽度限制，允许完全隐藏（往左拖动时）
        self.dir_tree.setMinimumWidth(0)
        # 移除最大宽度限制，允许用户根据需要拖动到任意宽度
        # self.dir_tree.setMaximumWidth(1800)  # 移除限制
        # 设置目录树的大小策略：宽度固定（用户手动拖动），高度自适应
        self.dir_tree.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
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

        # 右侧标签页内容区域（使用 StackedWidget 独立显示，不依赖 tab_widget）
        from PyQt5.QtWidgets import QStackedWidget
        self.content_stack = QStackedWidget()
        self.content_stack.setStyleSheet("background: white;")
        self.content_stack.setAutoFillBackground(True)
        # 允许右侧内容在窗口缩小时被压缩，避免阻止左侧目录树向左拖动
        self.content_stack.setMinimumWidth(0)
        
        self.splitter.addWidget(self.content_stack)
        
        # 设置左侧目录树和右侧内容的初始宽度比例（左:右 = 2:8，符合 Windows 11 风格）
        # 假设窗口总宽度1200px，左侧240px，右侧960px，根据DPI缩放
        left_width = int(240 * self.dpi_scale)
        right_width = int(960 * self.dpi_scale)
        self.splitter.setSizes([left_width, right_width])
        
        # 将分割器添加到主容器
        main_layout.addWidget(self.splitter)
        
        # 将内容布局添加到主容器
        container_layout.addLayout(main_layout)
        
        # 设置主容器为中心部件
        self.setCentralWidget(main_container)

        # 性能优化：延迟加载固定标签页（移到 _delayed_initialization）
        # 先添加一个默认标签页，避免窗口空白
        self.add_new_tab(QDir.homePath())
        
        # 连接信号
        self.open_path_signal.connect(self.handle_open_path_from_instance)
        
        # 性能优化：单实例服务器和Explorer监听移到延迟初始化
    
    def resizeEvent(self, event):
        """窗口大小改变事件"""
        super().resizeEvent(event)

    def handle_open_path_from_instance(self, path):
        """处理从其他实例接收到的路径（在主线程中）"""
        print(f"[MainWindow] Opening path in new tab: {path}")
        self.add_new_tab(path)
        # 激活并置顶窗口
        self.activateWindow()
        self.raise_()
        # 只在窗口最小化时恢复，保持最大化状态不变
        if self.isMinimized():
            self.showNormal()
    
    def _delayed_initialization(self):
        """延迟初始化非关键功能（性能优化）"""
        debug_print("[Performance] Starting delayed initialization...")
        # 进入恢复状态：在标题上显示提示
        self.is_restoring = True
        # 初次更新标题（无路径）
        self._update_window_title()
        
        # 延迟加载固定标签页
        try:
            has_pinned = self.load_pinned_tabs()
            # 如果有固定标签页，关闭默认的主目录标签
            if has_pinned and self.tab_widget.count() > 0:
                # 检查第一个标签是否是默认的主目录
                first_tab = self.get_tab_widget(0)
                if first_tab and hasattr(first_tab, 'current_path'):
                    if first_tab.current_path == QDir.homePath():
                        self.close_tab(0)
        except Exception as e:
            debug_print(f"[Performance] Failed to load pinned tabs: {e}")
        
        # 恢复缓存的标签页
        try:
            if self.config.get("enable_cache_tabs", True):
                cached_tabs = self.config.get("cached_tabs", [])
                if cached_tabs:
                    debug_print(f"[App] 恢复 {len(cached_tabs)} 个缓存标签页")
                    for tab_info in cached_tabs:
                        path = tab_info.get('path', '')
                        if path:
                            self.add_new_tab(path)
                    # 恢复后清除缓存
                    self.config['cached_tabs'] = []
                    self.save_config()
                    debug_print("[App] 缓存标签页已恢复并清除")
        except Exception as e:
            debug_print(f"[Performance] Failed to restore cached tabs: {e}")
        
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
        
        self.server_running = True
        server_thread_obj = threading.Thread(target=server_thread, daemon=True)
        server_thread_obj.start()
        # 等待服务器启动
        time.sleep(0.2)
    
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
        try:
            self.monitor_our_window = int(self.winId())  # 记录我们自己的窗口句柄
            self.explorer_monitoring = True
            debug_print("[Explorer Monitor] Starting Explorer window monitoring...")
            
            # 启动监听线程
            monitor_thread = threading.Thread(target=self._explorer_monitor_loop, daemon=True)
            monitor_thread.start()
        except Exception as e:
            debug_print(f"[Explorer Monitor] Failed to start: {e}")
    
    def stop_explorer_monitor(self):
        """停止Explorer窗口监听"""
        self.explorer_monitoring = False
        debug_print("[Explorer Monitor] Stopped")
    
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
                        # 调试：记录所有可见窗口的类名（仅在检测到新窗口时）
                        if win32gui.IsWindowVisible(hwnd):
                            title = win32gui.GetWindowText(hwnd)
                            # 记录可能相关的窗口信息
                            if title and len(title) > 0 and (':\\' in title or title.startswith('C:') or title.startswith('D:')):
                                debug_print(f"[Explorer Monitor] Debug: Found window - Class: '{class_name}', Title: '{title}'")
                        
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
                
                debug_print(f"[Explorer Monitor] Detected {len(new_windows)} new Explorer window(s)")
                
                for hwnd in new_windows:
                    # 检查是否是我们自己的窗口（避免误捕获嵌入的Explorer控件）
                    try:
                        # 获取窗口标题
                        title = win32gui.GetWindowText(hwnd)
                        
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
                        except:
                            pass
                        
                        debug_print(f"[Explorer Monitor] New Explorer window detected: {hwnd} - {title}")
                        
                        # 尝试获取Explorer窗口的当前路径
                        path = self._get_explorer_path(hwnd)
                        
                        if path:
                            debug_print(f"[Explorer Monitor] ✓ Path: {path}")
                            
                            # 在主线程中打开新标签页
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
                                
                                debug_print(f"[Explorer Monitor] LocationURL: {location}")
                                
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
        # 缓存非固定标签（如果启用）
        try:
            if self.config.get("enable_cache_tabs", True):
                cached_tabs = []
                for i in range(self.tab_widget.count()):
                    tab = self.get_tab_widget(i)
                    if tab and hasattr(tab, 'current_path'):
                        # 只缓存非固定标签
                        if not getattr(tab, 'is_pinned', False):
                            cached_tabs.append({
                                'path': tab.current_path,
                                'is_shell': tab.current_path.startswith('shell:') if tab.current_path else False
                            })
                
                if cached_tabs:
                    self.config['cached_tabs'] = cached_tabs
                    self.save_config()
                    debug_print(f"[App] 缓存了 {len(cached_tabs)} 个非固定标签页")
        except Exception as e:
            print(f"Error caching tabs: {e}")
        
        # 清除搜索缓存
        try:
            global _search_cache
            _search_cache.clear()
            debug_print("[App] 程序关闭，已清除搜索缓存")
        except Exception as e:
            print(f"Error clearing search cache: {e}")
        
        # 停止服务器
        self.server_running = False
        if hasattr(self, 'server_socket'):
            try:
                self.server_socket.close()
            except Exception as e:
                print(f"Error closing server socket: {e}")
        
        # 停止Explorer监听
        try:
            self.stop_explorer_monitor()
        except Exception as e:
            print(f"Error stopping explorer monitor: {e}")
        
        # 停止所有标签页中的定时器和COM对象
        try:
            for i in range(self.tab_widget.count()):
                tab = self.get_tab_widget(i)
                if hasattr(tab, '_path_sync_timer') and tab._path_sync_timer:
                    tab._path_sync_timer.stop()
                    tab._path_sync_timer.deleteLater()
                # 清理COM对象
                if hasattr(tab, 'explorer'):
                    try:
                        tab.explorer.clear()
                    except:
                        pass
        except Exception as e:
            print(f"Error stopping timers: {e}")
        
        super().closeEvent(event)


    def on_dir_tree_clicked(self, index):
        # 目录树点击，右侧当前标签页跳转
        if not index.isValid():
            return
        path = self.dir_model.filePath(index)
        current_tab = self.get_current_tab_widget()
        if hasattr(current_tab, 'navigate_to'):
            current_tab.navigate_to(path)

    def on_dir_model_directory_loaded(self, path):
        """目录树模型加载完成时，如果当前标签页正在显示该路径，则触发刷新"""
        try:
            tab = self.get_current_tab_widget()
            if not tab or not getattr(tab, 'current_path', None):
                return
            norm_loaded = os.path.abspath(path) if path else None
            norm_tab = os.path.abspath(tab.current_path) if tab.current_path else None
            if norm_loaded and norm_tab and norm_loaded == norm_tab:
                if hasattr(tab, '_schedule_refresh'):
                    tab._schedule_refresh(reason="dir_tree_loaded")
        except Exception as e:
            debug_print(f"[DirModel] directoryLoaded handler error: {e}")

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
                show_toast(self, "路径错误", f"路径不存在: {local_path}", level="warning")
        elif url.startswith('shell:'):
            self.add_new_tab(url, is_shell=True)
        elif os.path.isabs(url) and os.path.exists(url):
            self.add_new_tab(url)
        else:
            show_toast(self, "不支持的书签", f"暂不支持打开此类型书签: {url}", level="warning")

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
        
        delete_action = menu.addAction("🗑️ 删除书签")
        delete_action.triggered.connect(lambda: self.confirm_delete_bookmark(bookmark_id, bookmark_name))
        
        debug_print(f"[DEBUG] Showing menu...")
        menu.exec_(pos)
        debug_print(f"[DEBUG] Menu closed")
    
    def confirm_delete_bookmark(self, bookmark_id, bookmark_name):
        """直接删除书签并给出轻量提示"""
        self.delete_bookmark_by_id(bookmark_id)
        show_toast(self, "已删除", f"书签 '{bookmark_name}' 已删除", level="info")

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
        
        # 边框命中优先：在全局范围内拦截鼠标用于调整大小，避免子控件抢占事件
        if event.type() in (QEvent.MouseMove, QEvent.MouseButtonPress, QEvent.MouseButtonRelease):
            if isinstance(event, QMouseEvent):
                # 最大化时不处理边框
                if not self.isMaximized():
                    global_pos = event.globalPos()
                    local_pos = self.mapFromGlobal(global_pos)
                    
                    # 检查是否在标题栏+标签栏区域，如果是则跳过边缘检测（允许拖动窗口）
                    # 标题栏32px + 标签栏32px + 额外缓冲6px = 70px
                    titlebar_exclusion_height = int(70 * getattr(self, 'dpi_scale', 1.0))
                    in_titlebar = (hasattr(self, 'titlebar_widget') and 
                                  local_pos.y() >= 0 and 
                                  local_pos.y() <= titlebar_exclusion_height and
                                  local_pos.x() >= 0 and 
                                  local_pos.x() <= self.width())
                    
                    edge = self.detect_edge_global(global_pos) if not in_titlebar else None

                    if event.type() == QEvent.MouseMove:
                        if self.resizing and self.resize_direction:
                            self.resize_window(global_pos)
                            self.update_cursor(self.resize_direction)
                            return True  # 吃掉事件，避免子控件触发
                        if edge:
                            self.update_cursor(edge)
                        else:
                            self.update_cursor(None)
                        # 鼠标移动仅更新光标，不吃掉事件
                        return False

                    if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                        if edge:
                            self.resizing = True
                            self.resize_direction = edge
                            self.resize_start_pos = global_pos
                            self.resize_start_geometry = self.geometry()
                            self.update_cursor(edge)
                            return True  # 吃掉事件，防止按钮/目录树收到点击

                    if event.type() == QEvent.MouseButtonRelease:
                        if self.resizing:
                            self.resizing = False
                            self.resize_direction = None
                            self.update_cursor(None)
                            return True

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
            "设置已更新",
            f"Explorer窗口监听已{'启用' if checked else '禁用'}\n{'新打开的文件管理器窗口将自动嵌入到标签页中' if checked else '新打开的文件管理器窗口将独立显示'}",
            level="info",
        )

    def show_bookmark_manager_dialog(self):
        dlg = BookmarkManagerDialog(self.bookmark_manager, self)
        dlg.exec_()

class SettingsDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.resize(600, 500)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        from PyQt5.QtWidgets import QDialogButtonBox, QLabel, QGroupBox
        # Explorer监听设置组
        monitor_group = QGroupBox("Explorer监听设置")
        monitor_layout = QVBoxLayout()
        
        self.monitor_cb = QCheckBox("监听新Explorer窗口", self)
        self.monitor_cb.setChecked(config.get("enable_explorer_monitor", True))
        self.monitor_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        monitor_layout.addWidget(self.monitor_cb)
        
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
        interval_layout.addWidget(QLabel("（推荐: 2.0秒）"))
        interval_layout.addStretch(1)
        monitor_layout.addLayout(interval_layout)
        
        monitor_group.setLayout(monitor_layout)
        layout.addWidget(monitor_group)
        
        # 调试设置组
        debug_group = QGroupBox("调试设置")
        debug_layout = QVBoxLayout()
        
        self.debug_mode_cb = QCheckBox("启用调试输出（输出到终端）", self)
        self.debug_mode_cb.setChecked(config.get("debug_mode", False))
        self.debug_mode_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.debug_mode_cb.setToolTip("启用后将在终端输出调试信息，用于开发和问题排查")
        debug_layout.addWidget(self.debug_mode_cb)
        
        self.explorer_monitor_debug_cb = QCheckBox("启用 Explorer Monitor 调试输出", self)
        self.explorer_monitor_debug_cb.setChecked(config.get("explorer_monitor_debug", False))
        self.explorer_monitor_debug_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.explorer_monitor_debug_cb.setToolTip("单独控制 Explorer Monitor 的日志输出（需要先启用调试输出）")
        debug_layout.addWidget(self.explorer_monitor_debug_cb)
        
        debug_group.setLayout(debug_layout)
        layout.addWidget(debug_group)
        
        # 标签页设置组
        tabs_group = QGroupBox("标签页设置")
        tabs_layout = QVBoxLayout()
        
        self.cache_tabs_cb = QCheckBox("关闭时缓存当前标签页，下次启动时恢复", self)
        self.cache_tabs_cb.setChecked(config.get("enable_cache_tabs", True))
        self.cache_tabs_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.cache_tabs_cb.setToolTip("关闭软件时保存非固定标签，下次启动时自动恢复（不包括固定标签）")
        tabs_layout.addWidget(self.cache_tabs_cb)
        
        tabs_group.setLayout(tabs_layout)
        layout.addWidget(tabs_group)
        
        # 开机启动设置组
        startup_group = QGroupBox("开机启动设置")
        startup_layout = QVBoxLayout()
        
        self.auto_startup_cb = QCheckBox("开机自动启动 TabExplorer", self)
        self.auto_startup_cb.setChecked(self._is_auto_startup_enabled())
        self.auto_startup_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.auto_startup_cb.setToolTip("在 Windows 启动时自动运行 TabExplorer.exe")
        startup_layout.addWidget(self.auto_startup_cb)
        
        startup_group.setLayout(startup_layout)
        layout.addWidget(startup_group)
        
        # Git 工具设置组
        git_group = QGroupBox("Git 工具设置")
        git_layout = QVBoxLayout()
        
        self.tortoisegit_buttons_cb = QCheckBox("显示 TortoiseGit 快捷按钮（标题栏）", self)
        self.tortoisegit_buttons_cb.setChecked(config.get("enable_tortoisegit_buttons", False))
        self.tortoisegit_buttons_cb.setStyleSheet("font-size: 11pt; padding: 5px;")
        self.tortoisegit_buttons_cb.setToolTip("在标题栏显示 Git Log 和 Git Commit 快捷按钮")
        git_layout.addWidget(self.tortoisegit_buttons_cb)
        
        git_group.setLayout(git_layout)
        layout.addWidget(git_group)
        
        # 快捷键设置组
        hotkey_group = QGroupBox("快捷键设置")
        hotkey_layout = QVBoxLayout()
        
        hotkeys = config.get("hotkeys", {})
        
        self.hotkey_new_tab = QCheckBox("Ctrl+T - 新建标签页")
        self.hotkey_new_tab.setChecked(hotkeys.get("new_tab", True))
        hotkey_layout.addWidget(self.hotkey_new_tab)
        
        self.hotkey_close_tab = QCheckBox("Ctrl+W - 关闭当前标签页")
        self.hotkey_close_tab.setChecked(hotkeys.get("close_tab", True))
        hotkey_layout.addWidget(self.hotkey_close_tab)
        
        self.hotkey_reopen_tab = QCheckBox("Ctrl+Shift+T - 恢复关闭的标签页")
        self.hotkey_reopen_tab.setChecked(hotkeys.get("reopen_tab", True))
        hotkey_layout.addWidget(self.hotkey_reopen_tab)
        
        self.hotkey_switch_tab = QCheckBox("Ctrl+Tab / Ctrl+Shift+Tab - 切换标签页")
        self.hotkey_switch_tab.setChecked(hotkeys.get("switch_tab", True))
        hotkey_layout.addWidget(self.hotkey_switch_tab)
        
        self.hotkey_search = QCheckBox("Ctrl+F - 打开搜索对话框")
        self.hotkey_search.setChecked(hotkeys.get("search", True))
        hotkey_layout.addWidget(self.hotkey_search)
        
        self.hotkey_navigate = QCheckBox("Alt+Left/Right - 前进/后退")
        self.hotkey_navigate.setChecked(hotkeys.get("navigate", True))
        hotkey_layout.addWidget(self.hotkey_navigate)
        
        self.hotkey_go_up = QCheckBox("Alt+Up - 返回上级目录")
        self.hotkey_go_up.setChecked(hotkeys.get("go_up", True))
        hotkey_layout.addWidget(self.hotkey_go_up)
        
        self.hotkey_refresh = QCheckBox("F5 - 刷新当前路径")
        self.hotkey_refresh.setChecked(hotkeys.get("refresh", True))
        hotkey_layout.addWidget(self.hotkey_refresh)
        
        self.hotkey_add_bookmark = QCheckBox("Ctrl+D - 添加当前路径到书签")
        self.hotkey_add_bookmark.setChecked(hotkeys.get("add_bookmark", True))
        hotkey_layout.addWidget(self.hotkey_add_bookmark)

        self.hotkey_copy_filename = QCheckBox("Ctrl+Alt - 复制选中文件名（含后缀）")
        self.hotkey_copy_filename.setChecked(hotkeys.get("copy_filename", True))
        hotkey_layout.addWidget(self.hotkey_copy_filename)
        
        # 提示信息（放在快捷键设置框内）
        tip_label = QLabel("💡 提示：取消勾选可禁用对应的快捷键")
        tip_label.setStyleSheet("QLabel { color: #666; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }")
        hotkey_layout.addWidget(tip_label)
        
        hotkey_group.setLayout(hotkey_layout)
        layout.addWidget(hotkey_group)
        
        # 添加弹性空间，将按钮推到底部
        layout.addStretch(1)
        
        # 检查更新链接
        update_link = QLabel()
        update_link.setText('检查更新: <a href="https://github.com/caojinyuan/TabEx/releases">https://github.com/caojinyuan/TabEx/releases</a>')
        update_link.setOpenExternalLinks(True)
        update_link.setStyleSheet("QLabel { padding: 10px; font-size: 10pt; }")
        update_link.setTextFormat(Qt.RichText)
        update_link.setToolTip("点击链接在浏览器中打开 GitHub Releases 页面")
        layout.addWidget(update_link)
        
        # 按钮区域
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    
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
                    show_toast(self.parent(), "错误", "未找到 TabExplorer.exe，请确保程序已正确安装", level="error")
                    return False
                
                # 创建快捷方式（使用 win32com）
                from win32com.client import Dispatch
                
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = exe_path
                shortcut.WorkingDirectory = os.path.dirname(exe_path)
                shortcut.IconLocation = exe_path
                shortcut.save()
                
                show_toast(self.parent(), "成功", "已启用开机自动启动", level="success")
                return True
            else:
                # 删除快捷方式
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                    show_toast(self.parent(), "成功", "已禁用开机自动启动", level="success")
                return True
        except Exception as e:
            show_toast(self.parent(), "错误", f"设置开机启动失败: {e}", level="error")
            return False
    
    def _get_exe_path(self):
        """获取 TabExplorer.exe 路径"""
        # 如果是打包的exe，使用sys.executable
        import sys
        if getattr(sys, 'frozen', False):
            return sys.executable
        
        # 如果是开发环境，尝试查找同目录下的 TabExplorer.exe
        script_dir = os.path.dirname(os.path.abspath(__file__))
        exe_path = os.path.join(script_dir, "TabExplorer.exe")
        if os.path.exists(exe_path):
            return exe_path
        
        # 查找上级目录
        parent_dir = os.path.dirname(script_dir)
        exe_path = os.path.join(parent_dir, "TabExplorer.exe")
        if os.path.exists(exe_path):
            return exe_path
        
        return None
    
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
            self.parent().config["enable_cache_tabs"] = self.cache_tabs_cb.isChecked()
            self.parent().config["enable_tortoisegit_buttons"] = self.tortoisegit_buttons_cb.isChecked()
            
            # 保存快捷键配置
            self.parent().config["hotkeys"] = {
                "new_tab": self.hotkey_new_tab.isChecked(),
                "close_tab": self.hotkey_close_tab.isChecked(),
                "reopen_tab": self.hotkey_reopen_tab.isChecked(),
                "switch_tab": self.hotkey_switch_tab.isChecked(),
                "search": self.hotkey_search.isChecked(),
                "navigate": self.hotkey_navigate.isChecked(),
                "go_up": self.hotkey_go_up.isChecked(),
                "refresh": self.hotkey_refresh.isChecked(),
                "add_bookmark": self.hotkey_add_bookmark.isChecked()
            }
            
            # 保存到文件
            self.parent().save_config()
            
            # 应用设置
            set_debug_mode(self.parent().config.get("debug_mode", False))
            set_explorer_monitor_debug(self.parent().config.get("explorer_monitor_debug", False))
            self.parent().apply_tortoisegit_buttons_config()
            
            # 重新设置快捷键
            self.parent().setup_shortcuts()
        
        super().accept()

# 书签管理对话框（初步框架，后续可扩展重命名/新建/删除等功能）
from PyQt5.QtWidgets import QDialog, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QPushButton, QHBoxLayout, QInputDialog, QLabel
class BookmarkManagerDialog(QDialog):

    def move_item_up(self):
        item = self.tree.currentItem()
        if not item:
            show_toast(self, "未选择", "请先选择要上移的书签或文件夹。", level="warning")
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
            show_toast(self, "未选择", "请先选择要下移的书签或文件夹。", level="warning")
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
                show_toast(self, "已保存", "书签已保存", level="success")
        except Exception as e:
            debug_print(f"[BookmarkDrag] Error updating structure: {e}")
            import traceback
            traceback.print_exc()
            show_toast(self, "保存失败", f"拖拽保存失败: {e}", level="error")
    
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
            
            if node_type == '文件夹':
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
            elif node_type == '书签':
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
        self.setWindowTitle("书签管理")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.resize(600, 500)
        self.bookmark_manager = bookmark_manager
        layout = QVBoxLayout(self)
        
        # 使用标准树形控件（拖拽不自动保存）
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["名称", "类型", "路径"])
        self.tree.setColumnWidth(0, 250)  # 第一列宽一些
        
        # 启用拖拽
        self.tree.setDragEnabled(True)
        self.tree.setAcceptDrops(True)
        self.tree.setDropIndicatorShown(True)
        self.tree.setDragDropMode(QTreeWidget.InternalMove)
        self.tree.setSelectionMode(QTreeWidget.SingleSelection)
        
        layout.addWidget(self.tree)
        
        # 添加拖拽提示
        drag_hint = QLabel("💡 提示：可以拖动书签和文件夹调整顺序和层级，调整后点击【保存】按钮保存更改")
        drag_hint.setStyleSheet("QLabel { color: #666; background: #f0f0f0; padding: 8px; border-radius: 4px; font-size: 10pt; }")
        layout.addWidget(drag_hint)
        
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
        
        # 添加导入/导出按钮
        self.export_btn = QPushButton("📤 导出")
        self.export_btn.setToolTip("导出书签到JSON文件")
        self.export_btn.clicked.connect(self.export_bookmarks)
        btn_layout.addWidget(self.export_btn)
        
        self.import_btn = QPushButton("📥 导入")
        self.import_btn.setToolTip("从JSON文件导入书签")
        self.import_btn.clicked.connect(self.import_bookmarks)
        btn_layout.addWidget(self.import_btn)
        
        # 添加手动保存按钮
        self.save_btn = QPushButton("💾 保存")
        self.save_btn.setToolTip("保存当前书签顺序和层级")
        self.save_btn.clicked.connect(self.manual_save)
        btn_layout.addWidget(self.save_btn)
        
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)
    
    def manual_save(self):
        """手动保存书签"""
        try:
            self.on_items_moved()
        except Exception as e:
            show_toast(self, "保存失败", f"保存失败: {e}", level="error")
    
    def edit_item(self):
        item = self.tree.currentItem()
        if not item:
            show_toast(self, "未选择", "请先选择要编辑的书签或文件夹。", level="warning")
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
            show_toast(self, "未选择", "请先选择要删除的书签或文件夹。", level="warning")
            return
        node_id = item.data(0, 1)
        # 直接执行删除并给出提示，避免阻塞
        show_toast(self, "已删除", "选中的书签/文件夹已删除", level="info")
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
            show_toast(self, "未选择", "请先选择要重命名的书签或文件夹。", level="warning")
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
            "导出书签",
            default_name,
            "JSON Files (*.json);;All Files (*)"
        )
        
        if file_path:
            try:
                # 复制当前的bookmarks.json到目标位置
                shutil.copy2("bookmarks.json", file_path)
                show_toast(self, "导出成功", f"书签已成功导出到:\n{file_path}", level="success")
                print(f"[Bookmark Export] Successfully exported to: {file_path}")
            except Exception as e:
                show_toast(self, "导出失败", f"导出书签时出错:\n{str(e)}", level="error")
                print(f"[Bookmark Export] Error: {e}")
    
    def import_bookmarks(self):
        """从JSON文件导入书签"""
        from PyQt5.QtWidgets import QFileDialog
        import json
        
        # 打开文件选择对话框
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "导入书签",
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
                show_toast(self, "格式错误", "导入的文件格式不正确，必须包含 'bookmark_bar' 节点", level="warning")
                return
            
            # 默认选择更安全的“合并”模式，避免阻塞确认
            show_toast(self, "导入方式", "已自动选择合并模式，将导入内容追加到现有书签。", level="info")
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
                show_toast(self, "导入成功", f"成功导入 {count} 个书签项", level="success")
                print(f"[Bookmark Import] Merged {count} items from: {file_path}")
            else:
                show_toast(self, "提示", "导入的文件中没有书签内容", level="info")
            
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
            show_toast(self, "格式错误", "导入的文件不是有效的JSON格式", level="error")
            print(f"[Bookmark Import] Invalid JSON format: {file_path}")
        except Exception as e:
            show_toast(self, "导入失败", f"导入书签时出错:\n{str(e)}", level="error")
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
    import os
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
        debug_print(f"[Icon] Failed to create app icon: {e}")
    window = MainWindow()
    
    # 如果有路径参数，在新窗口中打开
    # 注意：MainWindow.__init__() 已经调用了 load_pinned_tabs()
    if path_to_open:
        # 无论是否有固定标签页，都添加目标路径作为新标签
        window.add_new_tab(path_to_open)
    
    # 启动时最大化显示
    window.showMaximized()
    
    # 确保启动时阴影状态正确（最大化时无阴影）
    window._update_window_shadow()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()


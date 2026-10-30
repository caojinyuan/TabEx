"""文件搜索窗口、搜索任务与结果缓存。"""

import base64
import json
import os
import queue
import re
import stat
import threading
import time
from collections import OrderedDict

from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt
from PyQt5.QtWidgets import (
    QAbstractItemView, QCheckBox, QDialog, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMenu, QPushButton,
    QStyledItemDelegate, QTableView, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from . import debuglog as _debuglog
from . import theme as _theme
from .i18n import tr
from .debuglog import debug_print
from .system import detect_notepad_plus_plus, format_file_size, launch_detached, launch_detached_async
from .widgets import _set_tool_icon, show_toast


# 性能优化配置常量
MAX_SEARCH_CACHE_SIZE = 50  # 搜索缓存最大数量
MAX_SEARCH_RESULTS = 1000000  # 单次搜索最大结果数
MAX_CACHED_RESULTS_PER_QUERY = 5000  # 单次搜索最多缓存结果数（控制内存占用）
CONTENT_SEARCH_CHUNK_SIZE = 10 * 1024 * 1024  # 内容搜索分块大小（10MB）
CONTENT_SEARCH_MAX_BYTES_PER_FILE = 64 * 1024 * 1024  # 单文件最多扫描64MB，避免超大文件拖慢整体
CONTENT_SEARCH_IN_MEMORY_THRESHOLD = 2 * 1024 * 1024  # 小文件（<=2MB）一次性读入内存后多编码匹配
CONTENT_SEARCH_ENCODINGS = ['utf-8', 'utf-8-sig', 'utf-16', 'utf-16-le', 'utf-16-be', 'gbk', 'gb2312', 'latin-1']
SEARCH_RESULT_QUEUE_MAXSIZE = 3000  # 搜索结果队列容量（降低高吞吐时溢出概率）
SEARCH_RESULT_BATCH_BASE = 100  # 搜索线程默认批量发送大小
SEARCH_RESULT_BATCH_MIN = 50  # 低积压时最小批量发送大小
SEARCH_RESULT_BATCH_MAX = 400  # 高积压时最大批量发送大小
SEARCH_METADATA_DEGRADE_ENABLED = True  # 队列高压时降级元数据(stat)获取
SEARCH_METADATA_DEGRADE_QUEUE_RATIO = 0.75  # 触发降级的队列占用比例
SEARCH_RESULT_TYPE_COL_WIDTH = 90
SEARCH_RESULT_DATE_COL_WIDTH = 155
SEARCH_RESULT_SIZE_COL_WIDTH = 100
# 修改时间筛选：(显示名, 天数)；0 不限，-1 从今天 0 点起
SEARCH_MODIFIED_RANGES = (('不限', 0), ('今天', -1), ('最近 7 天', 7), ('最近 30 天', 30), ('最近一年', 365))
_SIZE_UNITS = {'': 1, 'b': 1, 'k': 1024, 'kb': 1024, 'm': 1024 ** 2, 'mb': 1024 ** 2, 'g': 1024 ** 3, 'gb': 1024 ** 3}


def parse_size_text(text):
    """'10MB'、'1.5 g'、'500' → 字节数；空文本返回 None，无法识别时抛出 ValueError。"""
    value = str(text or '').strip().lower().replace(' ', '')
    if not value:
        return None
    match = re.fullmatch(r'(\d+(?:\.\d+)?)([kmg]?b?)', value)
    if not match:
        raise ValueError(text)
    return int(float(match.group(1)) * _SIZE_UNITS[match.group(2)])


def modified_after_timestamp(days, now=None):
    """修改时间筛选的起始时间戳；0 表示不限（返回 None），-1 表示今天 0 点。"""
    if not days:
        return None
    now = time.time() if now is None else now
    if days < 0:
        local = time.localtime(now)
        return time.mktime((local.tm_year, local.tm_mon, local.tm_mday, 0, 0, 0, 0, 0, -1))
    return now - days * 86400


def compile_search_pattern(keyword, match_case=False, match_whole_word=False, use_regex=False):
    """正则或全词匹配时返回编译后的模式；普通子串匹配返回 None。正则无效时抛出 re.error。"""
    if not use_regex and not match_whole_word:
        return None
    body = keyword if use_regex else re.escape(keyword)
    if match_whole_word:
        body = rf'\b(?:{body})\b'
    return re.compile(body, 0 if match_case else re.IGNORECASE)


def passes_search_filters(filters, size, mtime, is_dir):
    """大小筛选只保留文件（设置后排除文件夹）；修改时间筛选对文件和文件夹都生效。"""
    if not filters:
        return True
    min_size, max_size = filters.get('min_size'), filters.get('max_size')
    if (min_size is not None or max_size is not None) and is_dir:
        return False
    if min_size is not None and size < min_size:
        return False
    if max_size is not None and size > max_size:
        return False
    after = filters.get('modified_after')
    return after is None or mtime >= after


def apply_runtime_performance_config(perf_cfg=None):
    """将配置中的性能参数应用到运行时常量（带边界校验）。"""
    global CONTENT_SEARCH_CHUNK_SIZE
    global CONTENT_SEARCH_MAX_BYTES_PER_FILE
    global CONTENT_SEARCH_IN_MEMORY_THRESHOLD
    global SEARCH_RESULT_QUEUE_MAXSIZE
    global SEARCH_RESULT_BATCH_BASE
    global SEARCH_RESULT_BATCH_MIN
    global SEARCH_RESULT_BATCH_MAX
    global SEARCH_METADATA_DEGRADE_ENABLED
    global SEARCH_METADATA_DEGRADE_QUEUE_RATIO

    if not isinstance(perf_cfg, dict):
        return

    def _clamp_int(val, default_val, min_val, max_val):
        try:
            iv = int(val)
        except Exception:
            return default_val
        if iv < min_val:
            return min_val
        if iv > max_val:
            return max_val
        return iv

    def _clamp_float(val, default_val, min_val, max_val):
        try:
            fv = float(val)
        except Exception:
            return default_val
        if fv < min_val:
            return min_val
        if fv > max_val:
            return max_val
        return fv

    def _to_bool(val, default_val):
        if isinstance(val, bool):
            return val
        if isinstance(val, (int, float)):
            return bool(val)
        if isinstance(val, str):
            v = val.strip().lower()
            if v in ('1', 'true', 'yes', 'on'):
                return True
            if v in ('0', 'false', 'no', 'off'):
                return False
        return default_val

    CONTENT_SEARCH_CHUNK_SIZE = _clamp_int(
        perf_cfg.get("content_search_chunk_size", CONTENT_SEARCH_CHUNK_SIZE),
        CONTENT_SEARCH_CHUNK_SIZE,
        256 * 1024,
        64 * 1024 * 1024,
    )
    CONTENT_SEARCH_MAX_BYTES_PER_FILE = _clamp_int(
        perf_cfg.get("content_search_max_bytes_per_file", CONTENT_SEARCH_MAX_BYTES_PER_FILE),
        CONTENT_SEARCH_MAX_BYTES_PER_FILE,
        1 * 1024 * 1024,
        1024 * 1024 * 1024,
    )
    CONTENT_SEARCH_IN_MEMORY_THRESHOLD = _clamp_int(
        perf_cfg.get("content_search_in_memory_threshold", CONTENT_SEARCH_IN_MEMORY_THRESHOLD),
        CONTENT_SEARCH_IN_MEMORY_THRESHOLD,
        64 * 1024,
        16 * 1024 * 1024,
    )
    SEARCH_RESULT_QUEUE_MAXSIZE = _clamp_int(
        perf_cfg.get("search_result_queue_maxsize", SEARCH_RESULT_QUEUE_MAXSIZE),
        SEARCH_RESULT_QUEUE_MAXSIZE,
        200,
        100000,
    )
    SEARCH_RESULT_BATCH_BASE = _clamp_int(
        perf_cfg.get("search_result_batch_base", SEARCH_RESULT_BATCH_BASE),
        SEARCH_RESULT_BATCH_BASE,
        20,
        2000,
    )
    SEARCH_RESULT_BATCH_MIN = _clamp_int(
        perf_cfg.get("search_result_batch_min", SEARCH_RESULT_BATCH_MIN),
        SEARCH_RESULT_BATCH_MIN,
        10,
        1000,
    )
    SEARCH_RESULT_BATCH_MAX = _clamp_int(
        perf_cfg.get("search_result_batch_max", SEARCH_RESULT_BATCH_MAX),
        SEARCH_RESULT_BATCH_MAX,
        20,
        5000,
    )
    SEARCH_METADATA_DEGRADE_ENABLED = _to_bool(
        perf_cfg.get("search_metadata_degrade_enabled", SEARCH_METADATA_DEGRADE_ENABLED),
        SEARCH_METADATA_DEGRADE_ENABLED,
    )
    SEARCH_METADATA_DEGRADE_QUEUE_RATIO = _clamp_float(
        perf_cfg.get("search_metadata_degrade_queue_ratio", SEARCH_METADATA_DEGRADE_QUEUE_RATIO),
        SEARCH_METADATA_DEGRADE_QUEUE_RATIO,
        0.1,
        0.98,
    )

    # 保证内存阈值不大于单文件扫描上限
    if CONTENT_SEARCH_IN_MEMORY_THRESHOLD > CONTENT_SEARCH_MAX_BYTES_PER_FILE:
        CONTENT_SEARCH_IN_MEMORY_THRESHOLD = CONTENT_SEARCH_MAX_BYTES_PER_FILE
    if SEARCH_RESULT_BATCH_MIN > SEARCH_RESULT_BATCH_MAX:
        SEARCH_RESULT_BATCH_MIN = SEARCH_RESULT_BATCH_MAX
    if SEARCH_RESULT_BATCH_BASE < SEARCH_RESULT_BATCH_MIN:
        SEARCH_RESULT_BATCH_BASE = SEARCH_RESULT_BATCH_MIN
    elif SEARCH_RESULT_BATCH_BASE > SEARCH_RESULT_BATCH_MAX:
        SEARCH_RESULT_BATCH_BASE = SEARCH_RESULT_BATCH_MAX

    debug_print(
        "[Config] Performance applied:",
        f"chunk={CONTENT_SEARCH_CHUNK_SIZE}",
        f"max_file={CONTENT_SEARCH_MAX_BYTES_PER_FILE}",
        f"in_memory={CONTENT_SEARCH_IN_MEMORY_THRESHOLD}",
        f"queue={SEARCH_RESULT_QUEUE_MAXSIZE}",
        f"batch_base={SEARCH_RESULT_BATCH_BASE}",
        f"batch_min={SEARCH_RESULT_BATCH_MIN}",
        f"batch_max={SEARCH_RESULT_BATCH_MAX}",
        f"meta_degrade={SEARCH_METADATA_DEGRADE_ENABLED}",
        f"meta_ratio={SEARCH_METADATA_DEGRADE_QUEUE_RATIO}",
    )


class SearchCache:
    """搜索结果缓存，使用LRU策略"""
    def __init__(self, max_size=50):
        self.cache = OrderedDict()
        self.max_size = max_size
        self._lock = threading.RLock()
        self._timestamps = {}
    
    def get_key(self, search_path, keyword, search_filename, search_content, file_types, match_case=False, match_whole_word=False, use_everything=False, extra=''):
        """生成缓存键"""
        import hashlib
        key_str = f"{search_path}|{keyword}|{search_filename}|{search_content}|{file_types}|{match_case}|{match_whole_word}|{use_everything}|{extra}"
        return hashlib.md5(key_str.encode()).hexdigest()
    
    def get(self, key):
        """获取缓存结果"""
        with self._lock:
            if key in self.cache:
                if time.monotonic() - self._timestamps.get(key, 0) > 5:
                    self.cache.pop(key, None)
                    self._timestamps.pop(key, None)
                    return None
                self.cache.move_to_end(key)
                return self.cache[key]
        return None
    
    def put(self, key, value):
        """存储缓存结果"""
        with self._lock:
            self.cache[key] = value
            self._timestamps[key] = time.monotonic()
            self.cache.move_to_end(key)
            if len(self.cache) > self.max_size:
                expired_key, _value = self.cache.popitem(last=False)
                self._timestamps.pop(expired_key, None)
    
    def clear(self):
        """清空缓存"""
        with self._lock:
            self.cache.clear()
            self._timestamps.clear()


# 全局搜索缓存实例（使用配置常量）
_search_cache = SearchCache(max_size=MAX_SEARCH_CACHE_SIZE)


class _SearchCancelled(Exception):
    pass


def _rg_text(value):
    if 'text' in value:
        return value['text']
    return base64.b64decode(value['bytes']).decode('utf-8', errors='replace')


def iter_ripgrep_matches(task, keyword, search_path, extensions=(), match_case=False,
                        match_whole_word=False, use_regex=False, flush_callback=None):
    """Stream one match per file; bound pipe buffering and stop the process on cancellation."""
    import subprocess
    import tempfile

    command = [task.ripgrep_path, '--no-config', '--json', '--line-buffered', '--max-count', '1',
               '-uu', '-s' if match_case else '-i']
    if not use_regex:
        command.append('-F')
    if match_whole_word:
        command.append('-w')
    for extension in sorted(extensions):
        command.extend(['--iglob', f'*.{extension}'])
    passes = [[]] if keyword.isascii() else [[], ['-E', 'gbk']]
    for encoding in passes:
        with tempfile.TemporaryFile() as errors:
            process = subprocess.Popen(
                command + encoding + ['-e', keyword, '--', search_path],
                stdout=subprocess.PIPE, stderr=errors,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            messages = queue.Queue(maxsize=8)
            stop_reader = threading.Event()

            def enqueue(value):
                while not stop_reader.is_set():
                    try:
                        messages.put(value, timeout=.05)
                        return
                    except queue.Full:
                        continue

            def read_output():
                try:
                    while not stop_reader.is_set():
                        line = process.stdout.readline(1024 * 1024 + 1)
                        if not line:
                            break
                        if not line.endswith(b'\n') and len(line) > 1024 * 1024:
                            # rg JSON includes the entire matching line, even with --only-matching.
                            prefix = line.decode('utf-8', errors='replace')
                            if not prefix.startswith('{"type":"match",'):
                                raise ValueError(tr("搜索引擎输出过长"))
                            path, _end = json.JSONDecoder().raw_decode(prefix.split('"path":', 1)[1])
                            while line and not line.endswith(b'\n') and not stop_reader.is_set():
                                line = process.stdout.readline(65536)
                            enqueue({'type': 'match', 'data': {'path': path}, 'preview_unavailable': True})
                        else:
                            enqueue(json.loads(line))
                except (OSError, ValueError, KeyError, IndexError) as error:
                    enqueue(error)
                finally:
                    enqueue(None)

            reader = threading.Thread(target=read_output, name='rg-output', daemon=True)
            try:
                reader.start()
                while True:
                    if not task.is_searching:
                        raise _SearchCancelled()
                    try:
                        event = messages.get(timeout=.05)
                    except queue.Empty:
                        if flush_callback is not None:
                            flush_callback()
                        continue
                    if event is None:
                        break
                    if isinstance(event, Exception):
                        raise event
                    if event.get('type') != 'match':
                        continue
                    data = event['data']
                    result = {'path': os.path.normpath(_rg_text(data['path']))}
                    if event.get('preview_unavailable'):
                        result['preview'] = tr("匹配行过长，仅显示文件；请在编辑器中搜索")
                    else:
                        lines = data['lines']
                        raw_bytes = (lines['text'].encode('utf-8') if 'text' in lines
                                     else base64.b64decode(lines['bytes']))
                        encoding = 'utf-8'
                        try:
                            raw = raw_bytes.decode(encoding)
                        except UnicodeDecodeError:
                            encoding = 'gbk'
                            raw = raw_bytes.decode(encoding, errors='replace')
                        offsets = data.get('submatches', [])
                        offset = offsets[0]['start'] if offsets else 0
                        column = len(raw_bytes[:offset].decode(encoding, errors='replace'))
                        start = max(0, column - 120)
                        result.update(line=data['line_number'], column=column + 1,
                                      preview=('…' if start else '') + raw[start:start + 500].rstrip('\r\n'))
                    yield result
                while process.poll() is None:
                    if not task.is_searching:
                        raise _SearchCancelled()
                    stop_reader.wait(.05)
                if process.returncode not in (0, 1):
                    errors.seek(0)
                    raise OSError(errors.read(4096).decode('utf-8', errors='replace')
                                  or tr("搜索引擎退出码: {}").format(process.returncode))
            finally:
                stop_reader.set()
                if process.poll() is None:
                    process.kill()
                process.wait()
                if reader.ident is not None:
                    reader.join()
                process.stdout.close()


def _content_result(match, info):
    path = match['path']
    return dict(match, name=os.path.splitext(path)[0], full_path=path,
                file_type=os.path.splitext(path)[1][1:].upper() or tr("无"),
                date=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(info.st_mtime)),
                size=format_file_size(info.st_size), sort_date_ts=info.st_mtime,
                sort_size_bytes=info.st_size)


class _SearchResultQueue(queue.Queue):
    def __init__(self, cancel_event):
        super().__init__(maxsize=SEARCH_RESULT_QUEUE_MAXSIZE)
        self.cancel_event = cancel_event

    def put(self, item, block=True, timeout=None):
        while not self.cancel_event.is_set():
            try:
                return super().put(item, block=block, timeout=0.1 if block else None)
            except queue.Full:
                if not block:
                    raise
        raise _SearchCancelled()


class _SearchTask:
    def __init__(self, search_path, everything_path, max_results, ripgrep_path=None):
        self.search_path = search_path
        self.everything_path = everything_path
        self.ripgrep_path = ripgrep_path
        self.max_results = max_results
        self.cancel_event = threading.Event()
        self.result_queue = _SearchResultQueue(self.cancel_event)
        self.queue_overflow_count = 0
        self._running = True

    @property
    def is_searching(self):
        return self._running and not self.cancel_event.is_set()

    @is_searching.setter
    def is_searching(self, value):
        self._running = bool(value)

    def add_search_results_batch(self, items, timeout=0.5):
        self.result_queue.put({'type': 'result_batch', 'items': items})

    def search_with_everything(self, *args, **kwargs):
        return SearchDialog.search_with_everything(self, *args, **kwargs)

    def run(self, *args):
        try:
            if not os.path.isdir(self.search_path):
                raise ValueError(tr("路径不是文件夹:") + self.search_path)
            SearchDialog.do_search(self, *args)
        except _SearchCancelled:
            pass
        except Exception as error:
            if not self.cancel_event.is_set():
                self.result_queue.put({'type': 'status', 'text': str(error)})
        finally:
            self._running = False
            if not self.cancel_event.is_set():
                try:
                    self.result_queue.put({'type': 'finished'})
                except _SearchCancelled:
                    pass


class QuickFindResultsDialog(QDialog):
    def __init__(self, matched_paths, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr("选择匹配项"))
        self.resize(680, 420)
        self.selected_path = None

        layout = QVBoxLayout(self)

        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels([tr("名称"), tr("类型"), tr("完整路径")])
        self.table.setRowCount(len(matched_paths))
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.itemDoubleClicked.connect(self._accept_current_selection)

        for row, path in enumerate(matched_paths):
            name_item = QTableWidgetItem(os.path.basename(path))
            type_item = QTableWidgetItem(tr("文件夹") if os.path.isdir(path) else tr("文件"))
            path_item = QTableWidgetItem(path)
            name_item.setData(Qt.UserRole, path)
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, type_item)
            self.table.setItem(row, 2, path_item)

        if matched_paths:
            self.table.selectRow(0)

        layout.addWidget(self.table)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        ok_btn = QPushButton(tr("确定"), self)
        cancel_btn = QPushButton(tr("取消"), self)
        ok_btn.clicked.connect(self._accept_current_selection)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

    def _accept_current_selection(self):
        current_row = self.table.currentRow()
        if current_row < 0:
            return
        current_item = self.table.item(current_row, 0)
        if current_item is None:
            return
        self.selected_path = current_item.data(Qt.UserRole)
        if self.selected_path:
            self.accept()


# 自定义委托：在文件名列实现省略号在开头

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


class SearchResultsTableModel(QAbstractTableModel):
    _HEADER_KEYS = ["文件名", "类型", "修改日期", "大小"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows = []

    @property
    def HEADERS(self):
        return [tr(k) for k in self._HEADER_KEYS]

    def rowCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self._rows)

    def columnCount(self, parent=QModelIndex()):
        if parent.isValid():
            return 0
        return len(self.HEADERS)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row = self._rows[index.row()]
        column = index.column()

        if role == Qt.DisplayRole:
            if column == 0:
                return row.get('name', '')
            if column == 1:
                return row.get('file_type', '')
            if column == 2:
                return row.get('date', '')
            if column == 3:
                return row.get('size', '')
        elif role == Qt.ToolTipRole:
            if column == 0:
                return row.get('full_path') or row.get('path', '')
            return row.get('path', '')
        elif role == Qt.UserRole:
            return row.get('path', '')
        elif role == Qt.TextAlignmentRole and column == 0:
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal and 0 <= section < len(self.HEADERS):
            return self.HEADERS[section]
        return super().headerData(section, orientation, role)

    def sort(self, column, order=Qt.AscendingOrder):
        key_map = {
            0: lambda row: str(row.get('name', '')).lower(),
            1: lambda row: str(row.get('file_type', '')).lower(),
            2: lambda row: (row.get('sort_date_ts') is None, row.get('sort_date_ts') if row.get('sort_date_ts') is not None else 0, str(row.get('date', ''))),
            3: lambda row: (row.get('sort_size_bytes') is None, row.get('sort_size_bytes') if row.get('sort_size_bytes') is not None else -1, str(row.get('size', ''))),
        }
        key_fn = key_map.get(column)
        if key_fn is None or len(self._rows) <= 1:
            return
        self.layoutAboutToBeChanged.emit()
        self._rows.sort(key=key_fn, reverse=(order == Qt.DescendingOrder))
        self.layoutChanged.emit()

    def clear(self):
        self.beginResetModel()
        self._rows = []
        self.endResetModel()

    def append_results(self, rows):
        if not rows:
            return 0
        start_row = len(self._rows)
        end_row = start_row + len(rows) - 1
        self.beginInsertRows(QModelIndex(), start_row, end_row)
        normalized_rows = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            row_copy = dict(row)
            row_copy.setdefault('sort_date_ts', None)
            row_copy.setdefault('sort_size_bytes', None)
            normalized_rows.append(row_copy)
        self._rows.extend(normalized_rows)
        self.endInsertRows()
        return len(normalized_rows)

    def path_for_row(self, row):
        if 0 <= row < len(self._rows):
            return self._rows[row].get('path', '')
        return ''

    def result_for_row(self, row):
        return dict(self._rows[row]) if 0 <= row < len(self._rows) else {}

    def snapshot_rows(self, limit=None):
        rows = list(self._rows)
        if isinstance(limit, int) and limit > 0:
            return rows[:limit]
        return rows


# Everything 搜索引擎集成
def _bundled_tool_path(file_name):
    from .paths import get_bundled_path
    return get_bundled_path(file_name)


def get_bundled_es_path():
    return _bundled_tool_path('es.exe')


def detect_ripgrep():
    """获取 rg.exe：优先使用内置版本，其次回退到 PATH 中已有的 rg.exe。"""
    import shutil
    return _bundled_tool_path('rg.exe') or shutil.which('rg.exe')


def detect_everything():
    """获取 es.exe：优先使用内置版本，其次回退到系统中已有的 es.exe。"""
    import shutil
    bundled = get_bundled_es_path()
    if bundled:
        return bundled
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


def is_text_file(file_path, sample_size=1024):
    """智能检测文件是否为文本文件（读取前N字节检测）。

    模块级实现：避免在每次搜索时重新定义闭包，且可独立复用/测试。
    """
    try:
        with open(file_path, 'rb') as f:
            sample = f.read(sample_size)
            if not sample:
                return True  # 空文件视为文本文件

            # UTF-16/UTF-32 BOM 视为文本，避免被 NULL 字节规则误判
            if sample.startswith((b'\xff\xfe', b'\xfe\xff', b'\xff\xfe\x00\x00', b'\x00\x00\xfe\xff')):
                return True

            # NULL 字节几乎总是二进制特征
            if b'\x00' in sample:
                return False

            # 仅统计 ASCII 控制字符（排除 \t\n\r），避免把 UTF-8 非 ASCII 文本误判为二进制
            control_count = 0
            for byte in sample:
                if byte < 9 or (13 < byte < 32):
                    control_count += 1

            # 控制字符占比过高，判定为二进制
            if control_count / len(sample) > 0.1:
                return False

            return True
    except Exception:
        return False  # 无法读取则视为二进制


# 搜索对话框
class SearchDialog(QDialog):    
    def __init__(self, search_path, parent=None, search_history=None):
        super().__init__(parent)
        self.setWindowTitle(tr("搜索 - {}").format(search_path))
        # 设置为可调整大小，并显示最小化/最大化按钮
        self.setWindowFlags(
            Qt.Dialog
            | Qt.WindowTitleHint
            | Qt.WindowCloseButtonHint
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
            | Qt.WindowSystemMenuHint
        )
        # 关闭时立即销毁 C++ 对象（释放所有 Qt 子控件占用的内存）
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.resize(800, 500)  # 初始大小，但允许调整
        self.search_path = search_path
        self.main_window = parent
        self.search_thread = None
        self.is_searching = False
        self.search_history = search_history or []  # 搜索历史列表
        
        # 检测Everything
        self.everything_path = detect_everything()
        self.ripgrep_path = detect_ripgrep()
        self.notepad_plus_plus_path = detect_notepad_plus_plus()
        debug_print(f"[Search] Everything detected: {self.everything_path}")
        
        # 线程安全的结果队列（限制大小防止内存溢出）
        import queue
        self.result_queue = queue.Queue(maxsize=SEARCH_RESULT_QUEUE_MAXSIZE)  # 结果队列容量
        self.ui_update_timer = None
        self.queue_overflow_count = 0  # 队列溢出计数
        self._queue_idle_ticks = 0
        
        # 结果限制配置（使用虚拟滚动优化，支持更多结果）
        self.max_results = 1000000  # 最多显示100万个结果（虚拟滚动优化）
        self.current_result_count = 0
        self.batch_insert_size = 500  # 批量插入大小
        
        layout = QVBoxLayout(self)
        
        # 搜索选项区域
        search_options = QHBoxLayout()
        search_options.setSpacing(5)  # 设置控件间距为5像素
        
        # 搜索关键词（改为QComboBox支持历史记录）
        search_label = QLabel(tr("搜索:"))
        search_options.addWidget(search_label)
        from PyQt5.QtWidgets import QComboBox, QGridLayout, QToolButton, QStyle
        self.search_input = QComboBox()
        self.search_input.setEditable(True)
        self.search_input.setInsertPolicy(QComboBox.NoInsert)  # 不自动插入新条目
        self.search_input.setMinimumWidth(140)
        self.search_input.lineEdit().setPlaceholderText(tr("输入搜索关键词..."))
        self.search_input.lineEdit().returnPressed.connect(self.start_search)
        # 填充历史记录
        if self.search_history:
            self.search_input.addItems(self.search_history)
        search_options.addWidget(self.search_input, 1)  # 添加stretch factor，让搜索框可以拉伸
        
        # 搜索按钮
        self.search_btn = QPushButton(tr("搜索"))
        _set_tool_icon(self.search_btn, 'edit-find', QStyle.SP_FileDialogContentsView)
        self.search_btn.clicked.connect(self.start_search)
        search_options.addWidget(self.search_btn)
        
        # 停止按钮
        self.stop_btn = QPushButton(tr("停止"))
        _set_tool_icon(self.stop_btn, 'process-stop', QStyle.SP_BrowserStop)
        self.stop_btn.clicked.connect(self.stop_search)
        self.stop_btn.setEnabled(False)
        search_options.addWidget(self.stop_btn)

        # AI 总结按钮
        self.ai_summary_btn = QPushButton(tr("AI总结结果"))
        _set_tool_icon(self.ai_summary_btn, 'help-contents', QStyle.SP_FileDialogInfoView)
        self.ai_summary_btn.clicked.connect(self.request_ai_search_summary)
        
        layout.addLayout(search_options)
        
        # 搜索路径输入框（可编辑）
        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel(tr("搜索路径:")))
        self.path_input = QLineEdit(search_path)
        self.path_input.setClearButtonEnabled(True)
        _theme.bind_style(self.path_input, "QLineEdit { color: #0066cc; font-weight: bold; padding: 5px; }")
        self.path_input.setPlaceholderText(tr("输入要搜索的文件夹路径..."))
        path_layout.addWidget(self.path_input)
        layout.addLayout(path_layout)
        
        # 搜索类型选择
        type_options = QHBoxLayout()
        type_options.setSpacing(10)
        self.search_filename_cb = QCheckBox(tr("搜索文件名"))
        self.search_filename_cb.setChecked(True)
        type_options.addWidget(self.search_filename_cb)
        
        self.search_content_cb = QCheckBox(tr("搜索文件内容"))
        self.search_content_cb.setChecked(True)  # 默认也选中
        type_options.addWidget(self.search_content_cb)

        self.match_case_cb = QCheckBox(tr("区分大小写"))
        self.match_case_cb.setToolTip(tr("区分大小写匹配文件名和文件内容"))
        type_options.addWidget(self.match_case_cb)

        self.match_whole_word_cb = QCheckBox(tr("全词匹配"))
        self.match_whole_word_cb.setToolTip(tr("仅匹配完整单词，避免命中更长字符串的一部分"))
        type_options.addWidget(self.match_whole_word_cb)

        self.use_regex_cb = QCheckBox(tr("正则"))
        self.use_regex_cb.setToolTip(tr("按正则表达式匹配文件名和文件内容（Python 语法）"))
        type_options.addWidget(self.use_regex_cb)
        
        # Everything搜索选项
        self.use_everything_cb = QCheckBox(tr("使用 Everything（不支持搜索文件内容）"))
        if self.everything_path:
            self.use_everything_cb.setChecked(True)  # 如果有Everything，默认启用
            self.use_everything_cb.setToolTip(tr("使用Everything搜索引擎\n路径: {}\n仅支持搜索文件名，不支持搜索文件内容").format(self.everything_path))
        else:
            self.use_everything_cb.setEnabled(False)
            self.use_everything_cb.setToolTip(tr("未检测到Everything，请从 https://www.voidtools.com/ 下载安装"))
        self.use_everything_cb.stateChanged.connect(self.on_everything_toggled)
        self.on_everything_toggled(self.use_everything_cb.isChecked())
        type_options.addWidget(self.use_everything_cb)

        # 手动轻量模式已移除；队列高压时仍自动降级元数据
        type_options.addStretch(1)
        # 窄窗口下允许复选框文字被压缩，而不是撑宽整个对话框
        for checkbox in (self.search_filename_cb, self.search_content_cb, self.match_case_cb,
                         self.match_whole_word_cb, self.use_regex_cb, self.use_everything_cb):
            checkbox.setMinimumWidth(24)
        layout.addLayout(type_options)
        
        # 文件类型过滤
        file_type_layout = QHBoxLayout()
        file_type_layout.addWidget(QLabel(tr("文件类型:")))
        self.file_type_input = QLineEdit()
        self.file_type_input.setClearButtonEnabled(True)
        self.file_type_input.setPlaceholderText(tr("例如: *.c,*.h,*.xml (留空表示搜索所有类型)"))
        self.file_type_input.setText("*.c,*.h,*.xdm,*.arxml,*.xml")  # 默认值
        self.file_type_input.setStyleSheet("QLineEdit { padding: 5px; }")
        file_type_layout.addWidget(self.file_type_input)
        layout.addLayout(file_type_layout)

        # 大小 / 修改时间筛选与常用搜索
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel(tr("大小:")))
        self.min_size_input = QLineEdit()
        self.max_size_input = QLineEdit()
        for edit, hint in ((self.min_size_input, tr("最小，如 10KB")), (self.max_size_input, tr("最大，如 5MB"))):
            edit.setPlaceholderText(hint)
            edit.setClearButtonEnabled(True)
            edit.setToolTip(tr("支持 B、KB、MB、GB；设置大小后只显示文件"))
            edit.setMaximumWidth(130)
        filter_layout.addWidget(self.min_size_input)
        filter_layout.addWidget(QLabel("~"))
        filter_layout.addWidget(self.max_size_input)
        filter_layout.addSpacing(10)
        filter_layout.addWidget(QLabel(tr("修改时间:")))
        self.modified_combo = QComboBox()
        for label, days in SEARCH_MODIFIED_RANGES:
            self.modified_combo.addItem(tr(label), days)
        filter_layout.addWidget(self.modified_combo)
        filter_layout.addStretch(1)
        self.presets_btn = QToolButton()
        self.presets_btn.setText(tr("常用搜索"))
        self.presets_btn.setToolTip(tr("保存或套用搜索条件（不含搜索路径）"))
        self.presets_btn.setPopupMode(QToolButton.InstantPopup)
        self.presets_menu = QMenu(self.presets_btn)
        self.presets_menu.aboutToShow.connect(self._rebuild_presets_menu)
        self.presets_btn.setMenu(self.presets_menu)
        filter_layout.addWidget(self.presets_btn)
        layout.addLayout(filter_layout)
        
        # 状态标签
        self.status_label = QLabel(tr("就绪"))
        self.status_label.setWordWrap(True)
        self.status_label.setMinimumWidth(0)
        status_row = QHBoxLayout()
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.ai_summary_btn)
        layout.addLayout(status_row)
        
        # 结果表格
        self.result_model = SearchResultsTableModel(self)
        self.result_list = QTableView()
        self.result_list.setModel(self.result_model)
        self.result_list.horizontalHeader().setStretchLastSection(False)
        self.result_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.result_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.result_list.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.result_list.horizontalHeader().setSectionResizeMode(3, QHeaderView.Fixed)
        self.result_list.setColumnWidth(1, SEARCH_RESULT_TYPE_COL_WIDTH)
        self.result_list.setColumnWidth(2, SEARCH_RESULT_DATE_COL_WIDTH)
        self.result_list.setColumnWidth(3, SEARCH_RESULT_SIZE_COL_WIDTH)
        self.result_list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.result_list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.result_list.setWordWrap(False)
        self.result_list.doubleClicked.connect(self.on_result_double_clicked)
        self.result_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_list.customContextMenuRequested.connect(self.show_result_context_menu)
        # 启用排序功能
        self.result_list.setSortingEnabled(True)
        # 设置自定义委托，让文件名列的省略号显示在开头
        self.result_list.setItemDelegateForColumn(0, ElideLeftDelegate(self.result_list))
        # 设置行高和网格线
        self.result_list.verticalHeader().setDefaultSectionSize(max(28, self.fontMetrics().height() + 10))
        self.result_list.setShowGrid(False)
        self.result_list.setAlternatingRowColors(True)  # 启用交替行颜色
        # 设置表头样式
        _theme.bind_style(self.result_list, """
            QHeaderView::section {
                background-color: #E0E0E0;
                padding: 4px;
                border: 1px solid #C0C0C0;
                font-weight: bold;
            }
        """)
        layout.addWidget(self.result_list)

        from PyQt5.QtWidgets import QPlainTextEdit
        self.match_preview = QPlainTextEdit(self)
        self.match_preview.setReadOnly(True)
        self.match_preview.setMaximumHeight(110)
        self.match_preview.setPlaceholderText(tr("选择内容搜索结果查看匹配行"))
        layout.addWidget(self.match_preview)
        self.open_match_button = QPushButton(tr("在 VS Code 中打开匹配行"), self)
        self.open_match_button.setEnabled(False)
        self.open_match_button.clicked.connect(self.open_match_in_editor)
        layout.addWidget(self.open_match_button, 0, Qt.AlignRight)
        self.result_list.selectionModel().currentRowChanged.connect(self.update_match_preview)
        self.result_model.modelReset.connect(self.update_match_preview)
        self.result_model.layoutChanged.connect(self.update_match_preview)
        
        # 启动UI更新定时器
        from PyQt5.QtCore import QTimer
        self.ui_update_timer = QTimer(self)
        self.ui_update_timer.timeout.connect(self.update_ui_from_queue)
        self.ui_update_timer.start(80)  # 搜索时动态提速，空闲时降频

    def _drain_result_queue(self):
        if not self.result_queue:
            return
        while True:
            try:
                self.result_queue.get_nowait()
            except Exception:
                break

    def update_match_preview(self, *args):
        result = self.result_model.result_for_row(self.result_list.currentIndex().row())
        line = result.get('line')
        preview = result.get('preview', '')
        self.match_preview.setPlainText(
            (tr("第 {} 行").format(line) + '\n' if line else '') + preview)
        self.open_match_button.setEnabled(bool(line))

    def open_match_in_editor(self):
        from .context_tools import ToolContext, configured_tools, launch_context_tool
        result = self.result_model.result_for_row(self.result_list.currentIndex().row())
        if not result.get('line'):
            return
        config = getattr(self.main_window, 'config', {})
        try:
            tool = next((tool for tool in configured_tools(config) if tool.get('id') == 'vscode-line'), None)
        except ValueError as error:
            show_toast(self, tr("外部工具"), str(error), level="error")
            return
        if tool is None:
            show_toast(self, tr("外部工具"), tr("请先配置 VS Code 匹配行工具"), level="warning")
            return
        launch_context_tool(self, tool, ToolContext(
            cwd=os.path.dirname(result['path']), selected=(result['path'],),
            line=result['line'], column=result.get('column', 1)))

    def _release_search_resources(self):
        self.is_searching = False
        task = getattr(self, '_search_task', None)
        if task:
            task.cancel_event.set()
        self.queue_overflow_count = 0
        self._queue_idle_ticks = 0

        if self.ui_update_timer:
            self.ui_update_timer.stop()

        self._drain_result_queue()

        if hasattr(self, 'result_model') and self.result_model:
            self.result_model.clear()
        self.current_result_count = 0
        self.search_thread = None

    def closeEvent(self, event):
        self._release_search_resources()
        super().closeEvent(event)

    def _ensure_ui_update_timer(self, interval=None):
        if not self.ui_update_timer:
            return
        if interval is not None and self.ui_update_timer.interval() != interval:
            self.ui_update_timer.setInterval(interval)
        if not self.ui_update_timer.isActive():
            self.ui_update_timer.start()
    
    def update_ui_from_queue(self):
        """从队列中取出结果并更新UI（在主线程中调用，批量优化）"""
        try:
            pending = 0
            try:
                pending = self.result_queue.qsize()
            except Exception:
                pending = 0

            # 自适应轮询：有积压时提速，空闲时降频，减少空转开销
            target_interval = 80
            if self.is_searching:
                if pending > 600:
                    target_interval = 20
                elif pending > 120:
                    target_interval = 35
                else:
                    target_interval = 60
            else:
                target_interval = 220 if pending == 0 else 80
            if self.ui_update_timer.interval() != target_interval:
                self.ui_update_timer.setInterval(target_interval)

            # 批量处理结果（一次处理最多200个，加快队列消费）
            batch_results = []
            batch_count = 0
            # 根据队列积压动态调整消费批次，减少结果爆发时的UI延迟
            max_batch = 200
            try:
                if pending > 800:
                    max_batch = 500
                elif pending > 300:
                    max_batch = 350
            except Exception:
                pass
            
            while batch_count < max_batch:
                try:
                    item = self.result_queue.get_nowait()
                    
                    if item['type'] == 'result':
                        batch_results.append(item)
                        batch_count += 1
                    elif item['type'] == 'result_batch':
                        items = item.get('items') or []
                        if items:
                            batch_results.extend(items)
                            batch_count += len(items)
                    elif item['type'] == 'status':
                        self.status_label.setText(item['text'])
                    elif item['type'] == 'error':
                        self.status_label.setText(item['text'])
                    elif item['type'] == 'button':
                        if item['button'] == 'search':
                            self.search_btn.setEnabled(item['enabled'])
                        elif item['button'] == 'stop':
                            self.stop_btn.setEnabled(item['enabled'])
                    elif item['type'] == 'enable_sorting':
                        # 搜索完成后启用排序
                        self.result_list.setSortingEnabled(True)
                    elif item['type'] == 'finished':
                        self.is_searching = False
                        self.search_btn.setEnabled(True)
                        self.stop_btn.setEnabled(False)
                        self.result_list.setSortingEnabled(True)
                except Exception:
                    break  # 队列为空
            
            # 批量添加结果到表格（性能优化）
            if batch_results:
                self._append_results_to_table(batch_results)
                self._queue_idle_ticks = 0
            elif not self.is_searching:
                self._queue_idle_ticks += 1
                if self._queue_idle_ticks > 3 and pending == 0 and self.ui_update_timer.isActive():
                    self.ui_update_timer.stop()
                elif self._queue_idle_ticks > 1 and self.ui_update_timer.interval() != 250:
                    self.ui_update_timer.setInterval(250)
                
        except Exception as e:
            debug_print(f"[Search] UI update error: {e}")

    def _append_results_to_table(self, results):
        """批量渲染搜索结果到表格。"""
        from PyQt5.QtCore import Qt

        if not results:
            return

        remaining_capacity = self.max_results - self.current_result_count
        if remaining_capacity <= 0:
            return
        rows_to_add = results[:remaining_capacity]

        self.result_list.setUpdatesEnabled(False)
        added_count = self.result_model.append_results(rows_to_add)
        self.current_result_count += added_count
        self.result_list.setUpdatesEnabled(True)
        if not self.result_list.currentIndex().isValid() and self.current_result_count:
            self.result_list.setCurrentIndex(self.result_model.index(0, 0))
    
    def add_search_result(self, item):
        """添加单条搜索结果（通过队列，线程安全）。"""
        self.result_queue.put({'type': 'result', **item})

    def add_search_results_batch(self, items, timeout=0.5):
        """批量添加搜索结果到队列，减少高并发下的队列争用。"""
        if not items:
            return
        self.result_queue.put({'type': 'result_batch', 'items': items}, timeout=timeout)
    
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
        debug_print(tr("[Search] 搜索缓存已清除"))

    def _collect_search_filters(self):
        """读取大小/修改时间筛选；没有任何筛选时返回 None，大小无法识别时抛出 ValueError。"""
        min_size = parse_size_text(self.min_size_input.text())
        max_size = parse_size_text(self.max_size_input.text())
        if min_size is not None and max_size is not None and min_size > max_size:
            raise ValueError(f"{self.min_size_input.text().strip()} > {self.max_size_input.text().strip()}")
        filters = {'min_size': min_size, 'max_size': max_size,
                   'modified_after': modified_after_timestamp(self.modified_combo.currentData())}
        return filters if any(value is not None for value in filters.values()) else None

    # ── 常用搜索（保存在 config.json 的 search_presets，不含搜索路径） ──────────────
    def _search_presets(self):
        config = getattr(self.main_window, 'config', None)
        if not isinstance(config, dict):
            return None
        presets = config.get('search_presets')
        if not isinstance(presets, list):
            presets = config['search_presets'] = []
        return presets

    def _current_search_conditions(self):
        return {
            'keyword': self.search_input.currentText().strip(),
            'file_types': self.file_type_input.text().strip(),
            'search_filename': self.search_filename_cb.isChecked(),
            'search_content': self.search_content_cb.isChecked(),
            'match_case': self.match_case_cb.isChecked(),
            'match_whole_word': self.match_whole_word_cb.isChecked(),
            'use_regex': self.use_regex_cb.isChecked(),
            'use_everything': self.use_everything_cb.isChecked(),
            'min_size': self.min_size_input.text().strip(),
            'max_size': self.max_size_input.text().strip(),
            'modified_days': self.modified_combo.currentData(),
        }

    def apply_search_preset(self, preset):
        self.search_input.setCurrentText(str(preset.get('keyword', '')))
        self.file_type_input.setText(str(preset.get('file_types', '')))
        self.search_filename_cb.setChecked(bool(preset.get('search_filename', True)))
        self.match_case_cb.setChecked(bool(preset.get('match_case', False)))
        self.match_whole_word_cb.setChecked(bool(preset.get('match_whole_word', False)))
        self.use_regex_cb.setChecked(bool(preset.get('use_regex', False)))
        if self.use_everything_cb.isEnabled():
            self.use_everything_cb.setChecked(bool(preset.get('use_everything', False)))
        if self.search_content_cb.isEnabled():
            self.search_content_cb.setChecked(bool(preset.get('search_content', False)))
        self.min_size_input.setText(str(preset.get('min_size', '')))
        self.max_size_input.setText(str(preset.get('max_size', '')))
        index = self.modified_combo.findData(preset.get('modified_days', 0))
        self.modified_combo.setCurrentIndex(max(index, 0))

    def save_search_preset(self):
        from PyQt5.QtWidgets import QInputDialog
        presets = self._search_presets()
        if presets is None:
            return
        default = self.search_input.currentText().strip() or self.file_type_input.text().strip()
        name, accepted = QInputDialog.getText(self, tr("保存常用搜索"), tr("名称"), text=default)
        name = name.strip()
        if not accepted or not name:
            return
        preset = dict(self._current_search_conditions(), name=name)
        presets[:] = [item for item in presets if isinstance(item, dict) and item.get('name') != name] + [preset]
        self.main_window.save_config()
        show_toast(self, tr("常用搜索"), tr("已保存“{}”").format(name), level="success")

    def delete_search_preset(self, name):
        presets = self._search_presets()
        if presets is None:
            return
        presets[:] = [item for item in presets if isinstance(item, dict) and item.get('name') != name]
        self.main_window.save_config()

    def _rebuild_presets_menu(self):
        menu = self.presets_menu
        menu.clear()
        presets = self._search_presets()
        save_action = menu.addAction(tr("保存当前条件..."), self.save_search_preset)
        save_action.setEnabled(presets is not None)
        presets = [item for item in presets or [] if isinstance(item, dict) and item.get('name')]
        if not presets:
            return
        menu.addSeparator()
        for preset in presets:
            menu.addAction(preset['name'], lambda checked=False, item=preset: self.apply_search_preset(item))
        delete_menu = menu.addMenu(tr("删除常用搜索"))
        for preset in presets:
            delete_menu.addAction(preset['name'], lambda checked=False, name=preset['name']: self.delete_search_preset(name))
    
    def start_search(self):
        keyword = self.search_input.currentText().strip()  # 改用currentText获取输入或选中的文本
        keyword_is_empty = not keyword
        
        # 仅在有关键词时才添加到历史记录
        if not keyword_is_empty:
            if self.main_window and hasattr(self.main_window, 'add_search_history'):
                self.main_window.add_search_history(keyword)
                # 更新下拉列表
                self.search_input.clear()
                if hasattr(self.main_window, 'search_history'):
                    self.search_input.addItems(self.main_window.search_history)
                # 设置当前文本为刚刚搜索的关键词
                self.search_input.setCurrentText(keyword)
        
        # 空关键词时：禁用内容搜索（搜索空内容无意义），确保文件名搜索开启
        do_search_filename = self.search_filename_cb.isChecked()
        do_search_content = self.search_content_cb.isChecked() and not keyword_is_empty
        if keyword_is_empty and not do_search_filename:
            do_search_filename = True  # 空关键词时自动启用文件名搜索
        
        if not do_search_filename and not do_search_content:
            show_toast(self, tr("提示"), tr("请至少选择一种搜索类型"), level="warning")
            return
        
        # 获取并验证搜索路径
        search_path = self.path_input.text().strip()
        if not search_path:
            show_toast(self, tr("提示"), tr("请输入搜索路径"), level="warning")
            return
        
        # 检查是否是特殊路径（不支持搜索）
        if search_path.startswith('shell:'):
            show_toast(self, tr("不支持"), tr("不支持搜索特殊路径（shell:）"), level="warning")
            return
        
        # 更新搜索路径
        self.search_path = search_path
        use_regex = self.use_regex_cb.isChecked() and not keyword_is_empty
        try:
            compile_search_pattern(keyword, self.match_case_cb.isChecked(), self.match_whole_word_cb.isChecked(), use_regex)
            filters = self._collect_search_filters()
        except re.error as error:
            show_toast(self, tr("正则表达式无效"), str(error), level="warning")
            return
        except ValueError as error:
            show_toast(self, tr("大小格式无效"), tr("请输入如 10KB、5MB、1.5GB 的大小：{}").format(error), level="warning")
            return
        extra_key = f"{use_regex}|{self.min_size_input.text().strip()}|{self.max_size_input.text().strip()}|{self.modified_combo.currentData()}"
        previous_task = getattr(self, '_search_task', None)
        if previous_task:
            previous_task.cancel_event.set()
        self.is_searching = False
        self._search_task = _SearchTask(search_path, self.everything_path, self.max_results, self.ripgrep_path)
        self.result_queue = self._search_task.result_queue
        
        # 获取文件类型过滤
        file_types = self.file_type_input.text().strip()
        
        # 检查缓存
        global _search_cache
        use_everything = self.use_everything_cb.isChecked() if self.everything_path else False
        cache_key = _search_cache.get_key(
            search_path, keyword,
            do_search_filename,
            do_search_content,
            file_types,
            self.match_case_cb.isChecked(),
            self.match_whole_word_cb.isChecked(),
            use_everything,
            extra=extra_key,
        )
        cached_results = _search_cache.get(cache_key)
        
        if cached_results is not None:
            # 使用缓存结果
            debug_print(f"[Search] 使用缓存结果，共 {len(cached_results)} 个")
            self.result_model.clear()
            self.current_result_count = 0
            self.status_label.setText(tr("正在加载缓存结果..."))
            
            # 批量添加缓存结果
            sorting_enabled = self.result_list.isSortingEnabled()
            self.result_list.setSortingEnabled(False)

            self._append_results_to_table(cached_results)
            
            self.result_list.setSortingEnabled(sorting_enabled)
            self.result_list.setSortingEnabled(True)
            self.search_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            shown_count = min(len(cached_results), self.max_results)
            self.status_label.setText(tr("搜索完成（缓存），共显示 {} 个结果").format(shown_count))
            if self.ui_update_timer and self.ui_update_timer.isActive():
                self.ui_update_timer.stop()
            return
        
        # 清空之前的结果
        self.result_model.clear()
        self.current_result_count = 0  # 重置计数器
        
        # 搜索期间完全禁用排序（性能优化）
        self.result_list.setSortingEnabled(False)
        
        self.is_searching = True
        self.search_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        if keyword_is_empty:
            file_types_hint = f"（{file_types}）" if file_types else ""
            self.status_label.setText(tr("列举文件中{}... (最多显示{}个结果)").format(file_types_hint, self.max_results))
        else:
            self.status_label.setText(tr("搜索中... (最多显示{}个结果)").format(self.max_results))
        if self.ui_update_timer:
            self._queue_idle_ticks = 0
            self._ensure_ui_update_timer(20)
        
        # 在后台线程执行搜索
        import threading
        use_everything = self.use_everything_cb.isChecked() if self.everything_path else False
        match_case = self.match_case_cb.isChecked()
        match_whole_word = self.match_whole_word_cb.isChecked()
        self.search_thread = threading.Thread(
            target=self._search_task.run,
            args=(keyword, do_search_filename, do_search_content, file_types, cache_key, use_everything, match_case,
                  match_whole_word, use_regex, filters)
        )
        self.search_thread.daemon = True
        self.search_thread.start()
    
    def stop_search(self):
        self.is_searching = False
        task = getattr(self, '_search_task', None)
        if task:
            task.cancel_event.set()
        self._drain_result_queue()
        self.result_list.setSortingEnabled(True)
        self.search_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText(tr("已停止"))
        if self.ui_update_timer:
            self._queue_idle_ticks = 0
            self._ensure_ui_update_timer(220)
    
    def search_with_everything(self, keyword, search_path, file_types="", match_case=False, match_whole_word=False,
                               use_regex=False, filters=None):
        """使用Everything进行搜索"""
        import subprocess

        pattern = compile_search_pattern(keyword, match_case, match_whole_word, use_regex)

        def _matches_filename(name):
            if pattern is not None:
                return bool(pattern.search(name))
            if match_case:
                return keyword in name
            return keyword.lower() in name.lower()
        
        try:
            # 构建Everything命令
            cmd = [self.everything_path, '-max-results', str(self.max_results)]
            
            # 用 es.exe 的 -path 选项限定目录；内嵌引号的 path:"..." 经 Python 转义后 es.exe 无法解析，会返回空结果
            if search_path:
                cmd += ['-path', search_path]
            
            if keyword:
                cmd += ['-r', keyword] if use_regex else [keyword]
            # Everything 端先粗筛大小和日期，结果再由 passes_search_filters 精确过滤
            if filters:
                if filters.get('min_size') is not None:
                    cmd.append(f"size:>={filters['min_size']}")
                if filters.get('max_size') is not None:
                    cmd.append(f"size:<={filters['max_size']}")
                if filters.get('modified_after') is not None:
                    cmd.append('dm:>=' + time.strftime('%Y-%m-%d', time.localtime(filters['modified_after'])))
            
            # 添加文件类型过滤（为空则不限制）
            if file_types:
                extensions = []
                for ft in file_types.split(','):
                    ft = ft.strip()
                    if ft.startswith('*.'):
                        ft = ft[2:]  # 移除 *.
                    ft = ft.lstrip('.')
                    if ft:
                        extensions.append(ft)
                if extensions:
                    cmd.append('ext:' + ';'.join(extensions))
            
            # 执行Everything搜索
            if _debuglog._DEBUG_MODE:
                debug_print(f"[Everything] Executing: {' '.join(cmd)}")
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='ignore',
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
            deadline = time.monotonic() + 30
            try:
                while True:
                    if not self.is_searching:
                        return []
                    if time.monotonic() >= deadline:
                        raise subprocess.TimeoutExpired(cmd, 30)
                    try:
                        output, error_output = process.communicate(timeout=0.1)
                        break
                    except subprocess.TimeoutExpired:
                        continue
            finally:
                if process.poll() is None:
                    process.kill()
                process.communicate()
            
            if process.returncode == 0:
                # 解析结果（每行一个文件路径）
                lines = output.strip().split('\n')
                results = []
                for line in lines:
                    line = line.strip()
                    # es.exe 输出本身来自索引，避免逐条 exists 造成大量额外 I/O
                    if line and _matches_filename(os.path.basename(line)):
                        results.append(line)
                        if len(results) >= self.max_results:
                            break
                
                debug_print(f"[Everything] Found {len(results)} results")
                return results
            else:
                debug_print(f"[Everything] Error: {error_output}")
                raise RuntimeError(error_output or tr("Everything 搜索失败"))
        
        except subprocess.TimeoutExpired:
            debug_print("[Everything] Search timeout")
            raise RuntimeError(tr("Everything 搜索超时"))
        except Exception as e:
            debug_print(f"[Everything] Error: {e}")
            raise
    
    def do_search(self, keyword, search_filename, search_content, file_types="", cache_key=None, use_everything=False, match_case=False, match_whole_word=False,
                  use_regex=False, filters=None):
        metadata_degrade_count = 0
        whole_word_pattern = compile_search_pattern(keyword, match_case, match_whole_word, use_regex)
        whole_word_pattern_bytes = None
        if match_whole_word and not use_regex and keyword.isascii():
            byte_flags = 0 if match_case else re.IGNORECASE
            whole_word_pattern_bytes = re.compile(rb'\b' + re.escape(keyword.encode('ascii', errors='ignore')) + rb'\b', byte_flags)

        def _matches_text(value):
            if not isinstance(value, str) or not value:
                return False
            if whole_word_pattern is not None:
                return bool(whole_word_pattern.search(value))
            if match_case:
                return keyword in value
            return keyword.lower() in value.lower()

        def _matches_bytes(value):
            if not value:
                return False
            if whole_word_pattern_bytes is not None:
                return bool(whole_word_pattern_bytes.search(value))
            if match_case:
                return keyword_bytes in value
            return keyword_bytes_lower in value.lower()

        def _should_degrade_metadata():
            if not SEARCH_METADATA_DEGRADE_ENABLED or filters:
                return False
            try:
                pending = self.result_queue.qsize()
                queue_max = max(1, self.result_queue.maxsize)
                return (pending / queue_max) >= SEARCH_METADATA_DEGRADE_QUEUE_RATIO
            except Exception:
                return False

        # 如果使用Everything搜索
        if use_everything and self.everything_path:
            self.result_queue.put({'type': 'status', 'text': tr('Using Everything搜索引擎...')})
            
            try:
                results = self.search_with_everything(keyword, self.search_path, file_types, match_case=match_case,
                                                      match_whole_word=match_whole_word, use_regex=use_regex,
                                                      filters=filters)
                
                if not self.is_searching:
                    return
                
                # 将Everything结果添加到显示队列
                batch_size = 200
                batch_items = []
                shown_count = 0
                for file_path in results:
                    if not self.is_searching:
                        break
                    
                    try:
                        basename = os.path.basename(file_path)
                        name_without_ext, file_ext = os.path.splitext(basename)
                        path_without_ext = os.path.join(os.path.dirname(file_path), name_without_ext)
                        file_type = file_ext[1:].upper() if file_ext else tr("无")
                        sort_date_ts = None
                        sort_size_bytes = None
                        stat_info = None
                        if _should_degrade_metadata():
                            metadata_degrade_count += 1
                            mtime = "-"
                            size_str = "-"
                            is_dir = os.path.isdir(file_path)
                        else:
                            try:
                                stat_info = os.stat(file_path)
                                is_dir = stat.S_ISDIR(stat_info.st_mode)
                                mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                                sort_date_ts = stat_info.st_mtime
                                if is_dir:
                                    size_str = "-"
                                else:
                                    sort_size_bytes = stat_info.st_size
                                    size_str = format_file_size(stat_info.st_size)
                            except Exception:
                                mtime = "-"
                                size_str = "-"
                                is_dir = os.path.isdir(file_path)

                        if filters and (stat_info is None or not passes_search_filters(
                                filters, stat_info.st_size, stat_info.st_mtime, is_dir)):
                            continue

                        if is_dir:
                            display_name = f"📁 {file_path}"
                            file_type = tr('文件夹')
                        else:
                            display_name = f"📄 {path_without_ext}"

                        batch_items.append({
                            'path': file_path,
                            'name': display_name,
                            'full_path': f"{'📁' if is_dir else '📄'} {file_path}",
                            'file_type': file_type,
                            'date': mtime,
                            'size': size_str,
                            'sort_date_ts': sort_date_ts,
                            'sort_size_bytes': sort_size_bytes,
                        })
                        shown_count += 1
                        if len(batch_items) >= batch_size:
                            self.add_search_results_batch(batch_items, timeout=0.5)
                            batch_items = []
                    except Exception:
                        pass

                if batch_items:
                    try:
                        self.add_search_results_batch(batch_items, timeout=0.5)
                    except Exception:
                        pass
                
                # 搜索完成
                final_count = shown_count
                degrade_note = tr("，元数据降级 {} 条").format(metadata_degrade_count) if metadata_degrade_count > 0 else ""
                self.result_queue.put({'type': 'status', 'text': tr("Everything搜索完成，共找到 {} 个结果{}").format(final_count, degrade_note)})
                
            except Exception as e:
                self.result_queue.put({'type': 'error', 'text': f'Everything搜索错误: {str(e)}'})
            
            return
        
        # 原有的搜索逻辑
        found_count = 0
        keyword_lower = keyword.lower()
        # 正则模式不走按字节子串匹配的快速路径，统一解码后用正则匹配
        keyword_is_ascii = keyword.isascii() and not use_regex
        keyword_bytes = keyword.encode('ascii', errors='ignore') if keyword_is_ascii else b''
        keyword_bytes_lower = keyword_lower.encode('ascii', errors='ignore') if keyword_is_ascii else b''
        results_buffer = []  # 结果缓冲区
        base_buffer_size = SEARCH_RESULT_BATCH_BASE
        max_cached_results = MAX_CACHED_RESULTS_PER_QUERY
        all_results = [] if cache_key else None  # 仅在需要缓存时保存结果
        ext_encoding_cache = {}  # 扩展名 -> 最近成功编码（减少重复试错）

        def _adaptive_buffer_size():
            """根据结果队列积压动态调整发送批次，降低高压场景争用。"""
            try:
                pending = self.result_queue.qsize()
                queue_max = max(1, self.result_queue.maxsize)
                ratio = pending / queue_max
                if ratio >= 0.75:
                    return min(SEARCH_RESULT_BATCH_MAX, base_buffer_size * 3)
                if ratio >= 0.4:
                    return min(SEARCH_RESULT_BATCH_MAX, base_buffer_size * 2)
                if ratio <= 0.1:
                    return max(SEARCH_RESULT_BATCH_MIN, base_buffer_size // 2)
            except Exception:
                pass
            return base_buffer_size

        def _flush_results_buffer(force=False):
            nonlocal results_buffer
            if not results_buffer:
                return
            if not force:
                threshold = _adaptive_buffer_size()
                if len(results_buffer) < threshold:
                    return
            try:
                self.add_search_results_batch(results_buffer, timeout=0.5)
            except _SearchCancelled:
                raise
            except Exception:
                self.queue_overflow_count += len(results_buffer)
            results_buffer = []

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

        # 文本扩展名白名单：命中时跳过二进制探测，减少一次额外文件读取
        text_file_extensions = {
            'txt', 'md', 'rst', 'log', 'ini', 'cfg', 'conf', 'toml', 'yaml', 'yml', 'json', 'xml',
            'csv', 'tsv', 'sql', 'bat', 'ps1', 'sh', 'c', 'h', 'cpp', 'hpp', 'cc', 'cs', 'java',
            'py', 'js', 'ts', 'jsx', 'tsx', 'html', 'htm', 'css', 'scss', 'less', 'go', 'rs', 'php',
            'rb', 'swift', 'kt', 'm', 'mm', 'vue', 'svelte', 'dockerfile', 'gitignore', 'arxml', 'xdm'
        }
        
        # 对于文件内容搜索，使用编译的正则表达式可能更快（可选优化）
        # 但Python的内置字符串搜索已经很快，这里保持简单
        
        # 解析文件类型过滤（支持*.ext格式，逗号分隔）
        file_extensions = set()
        if file_types:
            for ft in file_types.split(','):
                ft = ft.strip()
                if ft.startswith('*.'):
                    file_extensions.add(ft[2:].lower())  # 去掉*.，只保留扩展名
                elif ft.startswith('.'):
                    file_extensions.add(ft[1:].lower())  # 去掉.，只保留扩展名
                elif ft:
                    file_extensions.add(ft.lower())  # 直接使用输入的扩展名
        
        # 调试信息：输出搜索路径
        debug_print(tr("[Search] 开始搜索路径: {}").format(self.search_path))
        debug_print(tr("[Search] 搜索关键词: {}").format(keyword))
        debug_print(tr("[Search] 搜索文件名: {}, 搜索内容: {}").format(search_filename, search_content))
        debug_print(tr("[Search] 文件类型过滤: {}").format(file_extensions if file_extensions else '所有类型'))
        
        scanned_files = 0
        skipped_binary_files = 0
        search_warnings = []
        search_failed = False
        streamed_paths = set()
        try:
            rg_matches = None
            if search_content and keyword and getattr(self, 'ripgrep_path', None):
                self.result_queue.put({'type': 'status', 'text': tr('使用 ripgrep 搜索文件内容...')})
                from contextlib import closing
                try:
                    with closing(iter_ripgrep_matches(
                            self, keyword, self.search_path, file_extensions, match_case,
                            match_whole_word, use_regex,
                            flush_callback=lambda: _flush_results_buffer(force=True))) as matches:
                        for match in matches:
                            path = match['path']
                            normalized = os.path.normcase(path)
                            if normalized in streamed_paths:
                                continue
                            extension = os.path.splitext(path)[1][1:].lower()
                            if extension in binary_file_extensions:
                                skipped_binary_files += 1
                                continue
                            scanned_files += 1
                            try:
                                info = os.stat(path)
                            except OSError as error:
                                debug_print(f"[Search] Metadata unavailable: {path}: {error}")
                                if len(search_warnings) < 3:
                                    search_warnings.append(tr("部分文件无法读取: {}").format(error))
                                continue
                            if not passes_search_filters(filters, info.st_size, info.st_mtime, False):
                                continue
                            streamed_paths.add(normalized)
                            item = _content_result(match, info)
                            results_buffer.append(item)
                            found_count += 1
                            if all_results is not None and len(all_results) < max_cached_results:
                                all_results.append(item)
                            # First results must be visible while the engine is still scanning.
                            _flush_results_buffer(force=found_count == 1)
                            if found_count >= max_results:
                                results_limited = True
                                break
                    rg_matches = streamed_paths
                    _flush_results_buffer(force=True)
                except _SearchCancelled:
                    raise
                except (OSError, ValueError, KeyError) as error:
                    warning = tr("ripgrep 不可用，已回退到普通搜索: {}").format(error)
                    debug_print(f"[Search] {warning}")
                    search_warnings.append(warning)
                    self.result_queue.put({'type': 'status', 'text': warning})
            folder_count = 0
            last_status_update_ms = int(time.time() * 1000)
            entries = () if results_limited or (rg_matches is not None and not search_filename) else os.walk(self.search_path)
            for root, dirs, files in entries:
                if not self.is_searching:
                    debug_print(tr("[Search] 搜索被中断"))
                    break
                
                folder_count += 1
                
                # 搜索文件夹名（空关键词时跳过目录，因为目录无扩展名无法匹配文件类型过滤）
                if search_filename and keyword:
                    for dirname in dirs:
                        if not self.is_searching:
                            break
                        
                        # 检查是否达到结果限制
                        if found_count >= max_results:
                            results_limited = True
                            break
                        
                        # 使用Python内置的字符串搜索（已优化）
                        if _matches_text(dirname):
                            dir_path = os.path.join(root, dirname)
                            if filters:
                                try:
                                    dir_stat = os.stat(dir_path)
                                except OSError:
                                    continue
                                if not passes_search_filters(filters, 0, dir_stat.st_mtime, True):
                                    continue
                            found_count += 1
                            
                            # 获取文件夹信息
                            if _should_degrade_metadata():
                                metadata_degrade_count += 1
                                mtime = "-"
                                size_str = "-"
                                sort_date_ts = None
                            else:
                                try:
                                    stat_info = os.stat(dir_path)
                                    mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                                    size_str = "-"  # 文件夹不显示大小
                                    sort_date_ts = stat_info.st_mtime
                                except Exception:
                                    mtime = "-"
                                    size_str = "-"
                                    sort_date_ts = None
                            
                            result_item = {
                                'path': dir_path,
                                'name': f"📁 {dir_path}",
                                'full_path': f"📁 {dir_path}",
                                'file_type': tr('文件夹'),
                                'date': mtime,
                                'size': size_str,
                                'sort_date_ts': sort_date_ts,
                                'sort_size_bytes': None,
                            }
                            results_buffer.append(result_item)
                            if all_results is not None and len(all_results) < max_cached_results:
                                all_results.append(result_item)  # 保存到缓存列表
                            
                            # 批量更新UI（队列满时等待）
                            _flush_results_buffer()
                
                # 检查是否达到结果限制
                if found_count >= max_results:
                    results_limited = True
                    break
                
                # 搜索文件名和文件内容
                for filename in files:
                    if not self.is_searching:
                        debug_print(tr("[Search] 搜索被中断（文件循环）"))
                        break

                    filename_lower = filename.lower()
                    name_without_ext, ext = os.path.splitext(filename)
                    file_ext = ext[1:].lower() if ext else ''
                    
                    # 检查是否达到结果限制
                    if found_count >= max_results:
                        results_limited = True
                        break
                    
                    # 优化：如果只搜索文件名，快速过滤不匹配的文件
                    if search_filename and not search_content:
                        if not _matches_text(filename):
                            continue  # 文件名不匹配，跳过
                    
                    # 检查文件类型过滤
                    if file_extensions and file_ext not in file_extensions:
                        # 调试：显示被过滤的文件（仅对特定文件名）
                        if 'TstMgr' in filename or scanned_files < 5:
                            debug_print(tr("[Search] 文件被类型过滤跳过: {}").format(filename))
                        continue  # 跳过不匹配的文件类型
                    
                    scanned_files += 1
                    
                    # 状态更新节流：每300个文件必更新；否则每32个文件检查一次时间阈值
                    should_update_status = False
                    if scanned_files % 300 == 0:
                        should_update_status = True
                    elif scanned_files % 32 == 0:
                        now_ms = int(time.time() * 1000)
                        if (now_ms - last_status_update_ms) >= 500:
                            should_update_status = True

                    if should_update_status:
                        status_text = tr("搜索中... 已扫描 {} 个文件，找到 {} 个结果").format(scanned_files, found_count)
                        try:
                            self.result_queue.put({'type': 'status', 'text': status_text}, timeout=0.1)
                            last_status_update_ms = int(time.time() * 1000)
                        except Exception:
                            pass  # 超时后继续搜索
                    
                    file_path = os.path.join(root, filename)
                    if os.path.normcase(os.path.normpath(file_path)) in streamed_paths:
                        continue
                    file_stat = None
                    if filters:
                        # 先按大小/日期过滤，避免对不符合条件的文件做内容扫描
                        try:
                            file_stat = os.stat(file_path)
                        except OSError:
                            continue
                        if not passes_search_filters(filters, file_stat.st_size, file_stat.st_mtime, False):
                            continue
                    matched = False
                    match_type = ""
                    
                    # 搜索文件名（Python内置优化）
                    if search_filename and _matches_text(filename):
                        matched = True
                        match_type = "📄"
                    
                    # 搜索文件内容：ripgrep 已预先给出命中文件集合，这里只按集合判断
                    if search_content and not matched and rg_matches is not None:
                        if file_ext in binary_file_extensions:
                            skipped_binary_files += 1
                            continue
                        if os.path.normcase(os.path.normpath(file_path)) in rg_matches:
                            matched = True
                            match_type = "📄"

                    # 搜索文件内容（智能检测文本文件）
                    if search_content and not matched and rg_matches is None:
                        # 1. 首先检查黑名单（明确的二进制文件）
                        if file_ext in binary_file_extensions:
                            skipped_binary_files += 1
                            continue

                        # 2. 预取文件大小：空文件不可能命中关键词，直接跳过；同时复用size避免重复stat
                        try:
                            file_size = os.path.getsize(file_path)
                        except OSError:
                            continue
                        if file_size == 0:
                            continue

                        # 3. 文本白名单直接通过；其余文件走探测
                        if file_ext not in text_file_extensions and not is_text_file(file_path):
                            skipped_binary_files += 1
                            continue
                        
                        try:
                            # 文件大小已在上方预取（file_size）

                            # 分块读取与单文件扫描上限
                            chunk_size = CONTENT_SEARCH_CHUNK_SIZE
                            max_scan_bytes = CONTENT_SEARCH_MAX_BYTES_PER_FILE
                            in_memory_threshold = CONTENT_SEARCH_IN_MEMORY_THRESHOLD

                            # ASCII关键词快速路径：直接按字节匹配，跳过多编码解码
                            if keyword_is_ascii and keyword_bytes:
                                read_limit = min(file_size, max_scan_bytes)
                                if read_limit <= in_memory_threshold:
                                    with open(file_path, 'rb') as bf:
                                        raw_content = bf.read(read_limit)
                                    # 兼容 UTF-16(无BOM) 等含 NULL 字节文本：移除 NULL 后再匹配一次
                                    raw_no_null = raw_content.replace(b'\x00', b'')
                                    if (
                                        _matches_bytes(raw_content)
                                        or _matches_bytes(raw_no_null)
                                    ):
                                        matched = True
                                        match_type = "📄"
                                else:
                                    overlap_bytes = max(1, len(keyword_bytes) * 2)
                                    scanned_bytes = 0
                                    with open(file_path, 'rb') as bf:
                                        while True:
                                            if scanned_bytes >= max_scan_bytes:
                                                break
                                            chunk = bf.read(chunk_size)
                                            if not chunk:
                                                break
                                            scanned_bytes += len(chunk)
                                            chunk_no_null = chunk.replace(b'\x00', b'')
                                            if (
                                                _matches_bytes(chunk)
                                                or _matches_bytes(chunk_no_null)
                                            ):
                                                matched = True
                                                match_type = "📄"
                                                break
                                            if len(chunk) == chunk_size:
                                                bf.seek(bf.tell() - overlap_bytes)

                            # 编码顺序：优先使用该扩展名最近成功编码
                            if not matched:  # 跳过编码尝试如果已匹配
                                preferred_encoding = ext_encoding_cache.get(file_ext)
                                if preferred_encoding and preferred_encoding in CONTENT_SEARCH_ENCODINGS:
                                    encodings = [preferred_encoding] + [enc for enc in CONTENT_SEARCH_ENCODINGS if enc != preferred_encoding]
                                else:
                                    encodings = list(CONTENT_SEARCH_ENCODINGS)
                                content_matched = False
                                
                                for encoding in encodings:
                                    try:
                                        read_limit = min(file_size, max_scan_bytes)
                                        if read_limit <= in_memory_threshold:
                                            # 小文件优化：只读一次二进制，再在内存中尝试不同编码，避免重复磁盘I/O
                                            if encoding == encodings[0]:
                                                with open(file_path, 'rb') as bf:
                                                    raw_content = bf.read(read_limit)
                                            content = raw_content.decode(encoding, errors='ignore')
                                            if _matches_text(content):
                                                matched = True
                                                match_type = "📄"
                                                content_matched = True
                                                if file_ext:
                                                    ext_encoding_cache[file_ext] = encoding
                                                break
                                        else:
                                            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                                                # 大文件分块读取（有总扫描上限）
                                                overlap = len(keyword) * 2  # 重叠区域，防止关键词被分割
                                                scanned_bytes = 0
                                                while True:
                                                    if scanned_bytes >= max_scan_bytes:
                                                        break
                                                    chunk = f.read(chunk_size)
                                                    if not chunk:
                                                        break
                                                    scanned_bytes += len(chunk)
                                                    if _matches_text(chunk):
                                                        matched = True
                                                        match_type = "📄"
                                                        content_matched = True
                                                        if file_ext:
                                                            ext_encoding_cache[file_ext] = encoding
                                                        break
                                                    # 回退overlap字节，避免关键词跨块
                                                    if len(chunk) == chunk_size:
                                                        f.seek(f.tell() - overlap)
                                                if content_matched:
                                                    break
                                    except UnicodeError:
                                        continue
                                    except Exception as e:
                                        # 其他错误，记录日志并尝试下一个编码
                                        debug_print(tr("[Search] 读取文件失败 {} (编码 {}): {}").format(file_path, encoding, e))
                                        continue
                        except Exception as e:
                            # 如果无法以文本方式读取，记录日志并跳过该文件
                            debug_print(tr("[Search] 无法读取文件 {}: {}").format(file_path, e))
                            pass
                    
                    if matched:
                        found_count += 1
                        
                        # 获取文件信息
                        sort_date_ts = None
                        sort_size_bytes = None
                        if _should_degrade_metadata():
                            metadata_degrade_count += 1
                            mtime = "-"
                            size_str = "-"
                        else:
                            try:
                                stat_info = file_stat or os.stat(file_path)
                                mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat_info.st_mtime))
                                size_bytes = stat_info.st_size
                                sort_date_ts = stat_info.st_mtime
                                sort_size_bytes = size_bytes
                                # 格式化大小
                                size_str = format_file_size(size_bytes)
                            except Exception:
                                mtime = "-"
                                size_str = "-"
                        
                        # 获取不带扩展名的完整路径
                        path_without_ext = os.path.join(root, name_without_ext)
                        file_type = file_ext.upper() if file_ext else tr("无")
                        
                        result_item = {
                            'path': file_path,
                            'name': f"{match_type} {path_without_ext}",
                            'full_path': f"{match_type} {file_path}",
                            'file_type': file_type,
                            'date': mtime,
                            'size': size_str,
                            'sort_date_ts': sort_date_ts,
                            'sort_size_bytes': sort_size_bytes,
                        }
                        results_buffer.append(result_item)
                        if all_results is not None and len(all_results) < max_cached_results:
                            all_results.append(result_item)  # 保存到缓存列表
                        
                        # 批量更新UI（每20个结果更新一次）
                        _flush_results_buffer()
        except _SearchCancelled:
            raise
        except Exception as e:
            debug_print(f"[Search] error: {e}")
            search_failed = True
            search_warnings.append(tr("搜索未完成: {}").format(e))
        
        # 添加剩余的结果（队列满时等待）
        if results_buffer:
            _flush_results_buffer(force=True)
        
        # 调试信息
        debug_print(tr("[Search] 搜索完成，共扫描 {} 个文件，找到 {} 个结果").format(scanned_files, found_count))
        if search_content and skipped_binary_files > 0:
            debug_print(tr("[Search] 跳过 {} 个二进制文件（不搜索内容）").format(skipped_binary_files))
        if self.queue_overflow_count > 0:
            debug_print(tr("[Search] ⚠️ 队列溢出 {} 次（部分结果未显示）").format(self.queue_overflow_count))
        
        # 将结果存入缓存（限制缓存大小，防止内存溢出）
        if (cache_key and all_results and self.is_searching and not results_limited
                and not search_warnings and found_count <= max_cached_results):
            global _search_cache
            cached_results = all_results
            _search_cache.put(cache_key, cached_results)
            debug_print(f"[Search] 已将 {len(cached_results)} 个结果存入缓存")
        
        # 重置搜索状态（先重置，避免后续更新被跳过）
        self.is_searching = False
        self.queue_overflow_count = 0  # 重置溢出计数
        
        # 搜索完成，更新UI状态（使用带超时的put，避免卡死）
        if results_limited:
            final_status = tr("搜索完成（已限制），显示前 {} 个结果（扫描了 {} 个文件）⚠️").format(found_count, scanned_files)
        else:
            final_status = tr("搜索完成，共找到 {} 个结果（扫描了 {} 个文件）").format(found_count, scanned_files)
        if metadata_degrade_count > 0:
            final_status += tr("，元数据降级 {} 条").format(metadata_degrade_count)
        if rg_matches is not None:
            final_status = tr("搜索完成，共找到 {} 个结果").format(found_count)
            if results_limited:
                final_status += tr("（已达结果上限）")
        if search_warnings:
            final_status += '\n' + '\n'.join(search_warnings)
        if search_failed:
            final_status = '\n'.join(search_warnings)
        
        # 使用超时put，防止队列满时卡死
        try:
            self.result_queue.put({'type': 'status', 'text': final_status}, timeout=1)
            self.result_queue.put({'type': 'button', 'button': 'search', 'enabled': True}, timeout=1)
            self.result_queue.put({'type': 'button', 'button': 'stop', 'enabled': False}, timeout=1)
            # 搜索完成后启用排序
            self.result_queue.put({'type': 'enable_sorting'}, timeout=1)
        except Exception:
            debug_print(tr('[Search] ⚠️ 队列满，最终状态更新失败'))
        
        debug_print(tr('[Search] UI更新已调度（使用队列）'))
    
    def on_result_double_clicked(self, index):
        """双击搜索结果，打开文件所在文件夹或文件夹本身，并选中文件"""
        if not index.isValid():
            return
        file_path = self.result_model.path_for_row(index.row())
        if file_path and os.path.exists(file_path):
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

    def _launch_file_with_program(self, file_path, program_path, display_name):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("该搜索结果不是可打开的文件"), level="warning")
            return False
        if not program_path:
            show_toast(self, tr("提示"), tr("未找到 {}").format(display_name), level="warning")
            return False
        try:
            launch_detached_async([program_path, file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法使用 {} 打开文件: {}").format(display_name, e), level="error")
            return False

    def open_result_with_notepad(self, file_path):
        self._launch_file_with_program(file_path, 'notepad.exe', tr('记事本'))

    def open_result_with_notepad_plus_plus(self, file_path):
        self._launch_file_with_program(file_path, self.notepad_plus_plus_path, 'Notepad++')

    def open_result_with_default_app(self, file_path):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("该搜索结果不是可打开的文件"), level="warning")
            return False
        try:
            if os.name == 'nt':
                os.startfile(file_path)
            else:
                launch_detached([file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法使用系统默认程序打开文件: {}").format(e), level="error")
            return False

    def open_result_with_system_dialog(self, file_path):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("该搜索结果不是可打开的文件"), level="warning")
            return False
        if os.name != 'nt':
            show_toast(self, tr("提示"), tr("当前系统不支持打开“选择其他应用”对话框"), level="warning")
            return False
        try:
            launch_detached_async(['rundll32.exe', 'shell32.dll,OpenAs_RunDLL', file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开“选择其他应用”对话框: {}").format(e), level="error")
            return False

    def show_result_context_menu(self, pos):
        index = self.result_list.indexAt(pos)
        if not index.isValid():
            return

        row = index.row()
        file_path = self.result_model.path_for_row(row)
        if not file_path or not os.path.exists(file_path):
            return

        self.result_list.selectRow(row)

        menu = QMenu(self)

        open_action = menu.addAction(tr("打开"))
        open_action.triggered.connect(lambda: self.on_result_double_clicked(index))

        open_folder_action = menu.addAction(tr("打开所在目录"))
        open_folder_action.triggered.connect(lambda: self._open_result_parent_folder(file_path))

        if os.path.isfile(file_path):
            from .context_tools import ToolContext, configured_tools, populate_tools_menu
            tools_menu = menu.addMenu(tr("外部工具"))
            result = self.result_model.result_for_row(row)
            try:
                populate_tools_menu(
                    tools_menu, configured_tools(getattr(self.main_window, 'config', {})),
                    ToolContext(cwd=os.path.dirname(file_path), selected=(file_path,),
                                line=result.get('line', 0), column=result.get('column', 1)), self)
            except ValueError as error:
                tools_menu.addAction(str(error)).setEnabled(False)
            default_action = menu.addAction(tr("用系统默认程序打开"))
            default_action.triggered.connect(lambda: self.open_result_with_default_app(file_path))

            menu.addSeparator()

            analyze_action = menu.addAction(tr("🤖 AI分析此文件"))
            analyze_action.triggered.connect(lambda: self.request_ai_file_analysis(file_path))

            menu.addSeparator()

            notepad_action = menu.addAction(tr("用记事本打开"))
            notepad_action.triggered.connect(lambda: self.open_result_with_notepad(file_path))

            notepadpp_action = menu.addAction(tr("用 Notepad++ 打开"))
            notepadpp_action.setEnabled(bool(getattr(self, 'notepad_plus_plus_path', None)))
            if not getattr(self, 'notepad_plus_plus_path', None):
                notepadpp_action.setToolTip(tr("未检测到 Notepad++"))
            notepadpp_action.triggered.connect(lambda: self.open_result_with_notepad_plus_plus(file_path))

            menu.addSeparator()

            system_dialog_action = menu.addAction(tr("选择其他应用..."))
            system_dialog_action.triggered.connect(lambda: self.open_result_with_system_dialog(file_path))

        menu.exec_(self.result_list.viewport().mapToGlobal(pos))

    def request_ai_file_analysis(self, file_path):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("该搜索结果不是可打开的文件"), level="warning")
            return
        if not self.main_window or not hasattr(self.main_window, 'send_prompt_to_ai'):
            show_toast(self, tr("提示"), tr("AI 助手未启用，请先在设置中开启"), level="warning")
            return

        max_chars = 12000
        try:
            file_size = os.path.getsize(file_path)
            ext = os.path.splitext(file_path)[1].lower() or "(none)"
            enc = 'utf-8'
            try:
                with open(file_path, 'r', encoding='utf-8') as _f:
                    _f.read(1)
            except UnicodeDecodeError:
                enc = 'gbk'

            with open(file_path, 'r', encoding=enc, errors='replace') as f:
                preview = f.read(max_chars)

            omitted_hint = ""
            if len(preview) >= max_chars:
                omitted_hint = f"\n\n[内容已截断，最多展示 {max_chars} 字符]"

            prompt = (
                "请对下面文件做结构化分析，输出以下小节：\n"
                "1) 文件用途\n2) 关键逻辑\n3) 风险点\n4) 建议下一步查看的相关文件\n"
                "如果信息不足，可以在结尾给出需要进一步读取的文件路径建议。\n\n"
                f"文件路径: {file_path}\n"
                f"文件扩展名: {ext}\n"
                f"文件大小: {file_size} bytes\n"
                f"预览编码: {enc}\n"
                "文件内容预览:\n"
                f"{preview}{omitted_hint}"
            )
        except Exception as e:
            show_toast(self, tr("错误"), tr("读取文件预览失败: {}").format(e), level="error")
            return

        if self.main_window.send_prompt_to_ai(prompt):
            show_toast(self, tr("提示"), tr("✅ 已发送到 AI 分析: {}").format(os.path.basename(file_path)), level="info")

    def request_ai_search_summary(self):
        rows = self.result_model.snapshot_rows(limit=120)
        if not rows:
            show_toast(self, tr("提示"), tr("请先完成搜索后再进行 AI 总结"), level="warning")
            return
        if not self.main_window or not hasattr(self.main_window, 'send_prompt_to_ai'):
            show_toast(self, tr("提示"), tr("AI 助手未启用，请先在设置中开启"), level="warning")
            return

        keyword = self.search_input.currentText().strip()
        file_types = self.file_type_input.text().strip()
        lines = []
        for row in rows:
            path = row.get('path', '')
            ftype = row.get('file_type', '')
            lines.append(f"- [{ftype}] {path}")

        prompt = (
            "请对下面的搜索结果做总结，输出以下小节：\n"
            "1) 结果分布（哪些目录最集中）\n"
            "2) 可能的核心实现文件（Top 10）\n"
            "3) 噪音/次要结果类型\n"
            "4) 建议的阅读顺序\n"
            "5) 如需后续动作，请给出可执行建议\n\n"
            f"搜索根路径: {self.search_path}\n"
            f"关键词: {keyword or '(空，当前为列目录模式)'}\n"
            f"文件类型过滤: {file_types or '(无)'}\n"
            f"结果总数(当前已加载): {self.result_model.rowCount()}\n"
            f"用于总结的样本数: {len(rows)}\n\n"
            "结果样本:\n" + "\n".join(lines)
        )

        if self.main_window.send_prompt_to_ai(prompt):
            show_toast(self, tr("提示"), tr("✅ 已发送搜索结果到 AI，总计 {} 条").format(len(rows)), level="info")

    def _open_result_parent_folder(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return
        if os.path.isdir(file_path):
            folder_path = file_path
            select_file = None
        else:
            folder_path = os.path.dirname(file_path)
            select_file = os.path.basename(file_path)
        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
            self.main_window.add_new_tab(folder_path, select_file=select_file)

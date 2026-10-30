"""常用目录：按访问次数和最近访问时间（frecency）排序，保存在配置目录的 frequent_dirs.json。"""

import json
import os
import time

from PyQt5.QtCore import QEvent, Qt
from PyQt5.QtWidgets import (
    QDialog, QDialogButtonBox, QHeaderView, QLineEdit, QTreeWidget, QTreeWidgetItem, QVBoxLayout,
)

from .i18n import tr
from .debuglog import debug_print
from .paths import get_app_config_path

MAX_FREQUENT_DIRS = 500
# 访问次数总和超过该值时整体衰减，旧的高频目录逐渐让位给近期常用目录
FREQUENT_AGING_THRESHOLD = 5000


def _frecency(entry, now):
    age = now - float(entry.get('last', 0) or 0)
    if age < 3600:
        weight = 4.0
    elif age < 86400:
        weight = 2.0
    elif age < 7 * 86400:
        weight = 0.5
    else:
        weight = 0.25
    return float(entry.get('count', 0) or 0) * weight


def is_recordable_dir(path):
    """只记录本地盘符路径和 UNC 共享路径，不记录 shell: 等虚拟位置。"""
    if not isinstance(path, str) or not path or path.startswith(('shell:', '::')):
        return False
    return bool(os.path.splitdrive(path)[0]) and os.path.isabs(path)


class FrequentDirs:
    def __init__(self, path=None):
        self.path = os.fspath(path) if path else get_app_config_path('frequent_dirs.json')
        self.entries = {}
        self.dirty = False
        self._load()

    @staticmethod
    def _key(path):
        return os.path.normcase(os.path.normpath(path))

    def _load(self):
        try:
            with open(self.path, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
        except FileNotFoundError:
            return
        except (OSError, ValueError) as error:
            debug_print(f"[Frequent] load failed: {error}")
            return
        for item in data if isinstance(data, list) else []:
            if isinstance(item, dict) and is_recordable_dir(item.get('path')):
                try:
                    count, last = float(item.get('count', 1)), float(item.get('last', 0))
                except (TypeError, ValueError):
                    continue
                self.entries[self._key(item['path'])] = {'path': item['path'], 'count': count, 'last': last}

    def record(self, path, now=None):
        if not is_recordable_dir(path):
            return False
        path = os.path.normpath(path)
        now = time.time() if now is None else now
        entry = self.entries.setdefault(self._key(path), {'path': path, 'count': 0.0, 'last': now})
        entry['path'] = path
        entry['count'] += 1
        entry['last'] = now
        if sum(item['count'] for item in self.entries.values()) > FREQUENT_AGING_THRESHOLD:
            for key in list(self.entries):
                self.entries[key]['count'] *= 0.9
                if self.entries[key]['count'] < 1:
                    del self.entries[key]
        if len(self.entries) > MAX_FREQUENT_DIRS:
            ranked = sorted(self.entries, key=lambda key: _frecency(self.entries[key], now))
            for key in ranked[:len(self.entries) - MAX_FREQUENT_DIRS]:
                del self.entries[key]
        self.dirty = True
        return True

    def remove(self, path):
        if self.entries.pop(self._key(path), None) is None:
            return False
        self.dirty = True
        return True

    def ranked(self, query='', limit=200, now=None):
        """按 frecency 降序返回条目；query 中每个词都需出现在路径里（不区分大小写）。"""
        now = time.time() if now is None else now
        words = query.casefold().split()
        items = [entry for entry in self.entries.values()
                 if all(word in entry['path'].casefold() for word in words)]
        items.sort(key=lambda entry: (-_frecency(entry, now), entry['path'].casefold()))
        return items[:limit]

    def save(self):
        if not self.dirty:
            return True
        tmp_path = self.path + '.tmp'
        try:
            with open(tmp_path, 'w', encoding='utf-8') as handle:
                json.dump(list(self.entries.values()), handle, ensure_ascii=False, indent=1)
            os.replace(tmp_path, self.path)
            self.dirty = False
            return True
        except OSError as error:
            debug_print(f"[Frequent] save failed: {error}")
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            return False


class FrequentDirsDialog(QDialog):
    """常用目录跳转：输入筛选，Enter 在当前标签打开，Ctrl+Enter 在新标签打开，Delete 移除记录。"""

    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.store = store
        self.selected_path = ''
        self.open_in_new_tab = False
        self.setWindowTitle(tr("常用目录"))
        self.resize(760, 440)
        layout = QVBoxLayout(self)
        self.filter_input = QLineEdit(self)
        self.filter_input.setPlaceholderText(tr("输入目录名或路径的一部分筛选；Enter 打开，Ctrl+Enter 新标签打开，Delete 移除"))
        self.filter_input.setClearButtonEnabled(True)
        layout.addWidget(self.filter_input)
        self.table = QTreeWidget(self)
        self.table.setHeaderLabels([tr("目录"), tr("访问次数"), tr("最近访问")])
        self.table.setRootIsDecorated(False)
        self.table.setUniformRowHeights(True)
        self.table.setAlternatingRowColors(True)
        self.table.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.setColumnWidth(1, 80)
        self.table.setColumnWidth(2, 140)
        layout.addWidget(self.table)
        buttons = QDialogButtonBox(QDialogButtonBox.Close, parent=self)
        buttons.button(QDialogButtonBox.Close).setText(tr("关闭"))
        open_button = buttons.addButton(tr("在当前标签打开"), QDialogButtonBox.AcceptRole)
        new_tab_button = buttons.addButton(tr("在新标签打开"), QDialogButtonBox.ActionRole)
        remove_button = buttons.addButton(tr("移除记录"), QDialogButtonBox.ActionRole)
        open_button.clicked.connect(lambda: self.activate_selected(False))
        new_tab_button.clicked.connect(lambda: self.activate_selected(True))
        remove_button.clicked.connect(self.remove_selected)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.filter_input.textChanged.connect(self.populate)
        self.filter_input.installEventFilter(self)
        self.table.installEventFilter(self)
        self.table.itemActivated.connect(lambda *_args: self.activate_selected(False))
        self.populate('')
        self.filter_input.setFocus()

    def populate(self, text=None):
        query = self.filter_input.text() if text is None else text
        self.table.clear()
        for entry in self.store.ranked(query):
            last = time.strftime('%Y-%m-%d %H:%M', time.localtime(entry['last'])) if entry['last'] else ''
            item = QTreeWidgetItem([entry['path'], str(int(round(entry['count']))), last])
            item.setToolTip(0, entry['path'])
            self.table.addTopLevelItem(item)
        if self.table.topLevelItemCount():
            self.table.setCurrentItem(self.table.topLevelItem(0))

    def eventFilter(self, watched, event):
        if event.type() == QEvent.KeyPress:
            key, modifiers = event.key(), event.modifiers()
            if watched is self.filter_input and key in (Qt.Key_Down, Qt.Key_Up):
                count = self.table.topLevelItemCount()
                if count:
                    row = self.table.indexOfTopLevelItem(self.table.currentItem()) if self.table.currentItem() else 0
                    row = max(0, min(count - 1, row + (1 if key == Qt.Key_Down else -1)))
                    self.table.setCurrentItem(self.table.topLevelItem(row))
                return True
            if key in (Qt.Key_Return, Qt.Key_Enter):
                self.activate_selected(bool(modifiers & Qt.ControlModifier))
                return True
            if key == Qt.Key_Delete and (watched is self.table or not self.filter_input.text()):
                self.remove_selected()
                return True
        return super().eventFilter(watched, event)

    def activate_selected(self, new_tab):
        item = self.table.currentItem()
        if item is None:
            return
        self.selected_path = item.text(0)
        self.open_in_new_tab = bool(new_tab)
        self.accept()

    def remove_selected(self):
        item = self.table.currentItem()
        if item is None:
            return
        row = self.table.indexOfTopLevelItem(item)
        self.store.remove(item.text(0))
        self.populate()
        count = self.table.topLevelItemCount()
        if count:
            self.table.setCurrentItem(self.table.topLevelItem(min(row, count - 1)))

"""A session-local collection of file references; removing entries never touches the filesystem."""

import os

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QDialog, QGridLayout, QLabel, QMessageBox, QPushButton, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout,
)

from .i18n import tr
from .widgets import show_toast


class FileBasketDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(tr("文件收集篮"))
        self.resize(800, 450)
        layout = QVBoxLayout(self)
        hint = QLabel(tr("只收集路径，不移动原文件；移除或清空不会删除文件。未选择时复制全部，关闭程序后清空。"), self)
        hint.setWordWrap(True)
        layout.addWidget(hint)
        self.table = QTreeWidget(self)
        self.table.setHeaderLabels([tr("名称"), tr("路径")])
        self.table.setRootIsDecorated(False)
        self.table.setUniformRowHeights(True)
        self.table.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.table.setColumnWidth(0, 200)
        layout.addWidget(self.table)
        self.count_label = QLabel(self)
        layout.addWidget(self.count_label)
        controls = QGridLayout()
        for index, (attr, label, callback) in enumerate((
                ('copy_button', '复制到当前目录...', self.copy_to_current),
                ('paths_button', '复制路径', self.copy_paths),
                ('remove_button', '移出收集篮', self.remove_selected),
                ('clear_button', '清空收集篮', self.clear))):
            button = QPushButton(tr(label), self)
            button.clicked.connect(callback)
            controls.addWidget(button, index // 2, index % 2)
            setattr(self, attr, button)
        layout.addLayout(controls)
        self.table.itemSelectionChanged.connect(self.update_controls)
        self.update_controls()

    def all_paths(self):
        return [self.table.topLevelItem(index).text(1) for index in range(self.table.topLevelItemCount())]

    def selected_or_all_paths(self):
        selected = self.table.selectedItems()
        return [item.text(1) for item in selected] if selected else self.all_paths()

    def add_paths(self, paths):
        known = {os.path.normcase(path) for path in self.all_paths()}
        added, invalid = 0, []
        for path in paths:
            if not isinstance(path, str) or not os.path.isabs(path) or path.startswith(('shell:', '::')):
                invalid.append(str(path))
                continue
            path = os.path.normpath(path)
            key = os.path.normcase(path)
            if key in known:
                continue
            item = QTreeWidgetItem([os.path.basename(path) or path, path])
            item.setToolTip(0, path)
            item.setToolTip(1, path)
            self.table.addTopLevelItem(item)
            known.add(key)
            added += 1
        if invalid:
            show_toast(self, tr("文件收集篮"), tr("不能收集非文件系统路径: {}").format('\n'.join(invalid[:3])),
                       level='warning')
        self.update_controls()
        return added

    def update_controls(self):
        count = self.table.topLevelItemCount()
        selected = len(self.table.selectedItems())
        self.count_label.setText(tr("已收集 {} 项，选中 {} 项").format(count, selected))
        for button in (self.copy_button, self.paths_button, self.clear_button):
            button.setEnabled(count > 0)
        self.remove_button.setEnabled(selected > 0)

    def remove_selected(self):
        for item in self.table.selectedItems():
            self.table.takeTopLevelItem(self.table.indexOfTopLevelItem(item))
        self.update_controls()

    def clear(self):
        self.table.clear()
        self.update_controls()

    def copy_paths(self):
        paths = self.selected_or_all_paths()
        if paths:
            QApplication.clipboard().setText('\n'.join(paths))
            show_toast(self, tr("文件收集篮"), tr("已复制 {} 项路径").format(len(paths)), level='success')

    def copy_to_current(self):
        owner = self.parent()
        pane = owner.get_active_pane()
        destination = getattr(pane, 'current_path', '')
        if not destination or destination.startswith(('shell:', '::')) or not os.path.isabs(destination):
            show_toast(self, tr("文件收集篮"), tr("请先在主窗口打开目标文件夹"), level='warning')
            return
        paths = self.selected_or_all_paths()
        if not paths:
            return
        text = tr("将 {} 项复制到 {}。同名项保留两者；来源失效时在文件任务中报告。").format(len(paths), destination)
        if QMessageBox.question(self, tr("文件收集篮"), text,
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        owner.get_file_task_panel().start_task(
            'copy', paths, destination, max_workers=owner.config.get('file_op_max_workers', 0))
        owner.show_file_tasks()

"""Search existing UI actions without duplicating their execution logic."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QLabel, QLineEdit, QTreeWidget, QTreeWidgetItem, QVBoxLayout

from .i18n import tr


def menu_commands(menu, prefix=''):
    for action in menu.actions():
        if action.isSeparator() or not action.isVisible():
            continue
        label = f'{prefix} / {action.text()}' if prefix else action.text()
        if action.menu() is not None:
            yield from menu_commands(action.menu(), label)
        else:
            yield label, action


class CommandPalette(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(tr('查找功能'))
        self.resize(660, 440)
        self.selected_action = None
        self.commands = []
        layout = QVBoxLayout(self)
        self.query = QLineEdit(self)
        self.query.setPlaceholderText(tr('输入功能名称或快捷键，例如：同步、比较、工作区'))
        self.query.setClearButtonEnabled(True)
        layout.addWidget(self.query)
        self.results = QTreeWidget(self)
        self.results.setHeaderLabels([tr('功能'), tr('状态')])
        self.results.setRootIsDecorated(False)
        self.results.setUniformRowHeights(True)
        self.results.setColumnWidth(0, 390)
        layout.addWidget(self.results)
        self.explanation = QLabel(self)
        self.explanation.setWordWrap(True)
        self.explanation.setTextFormat(Qt.PlainText)
        layout.addWidget(self.explanation)
        self.query.textChanged.connect(self.filter_commands)
        self.query.returnPressed.connect(self.choose_current)
        self.results.itemActivated.connect(self.choose_current)
        self.results.currentItemChanged.connect(self.describe_current)
        self.query.installEventFilter(self)

    def set_commands(self, commands):
        self.commands = list(commands)
        self.filter_commands()
        self.query.setFocus()

    def filter_commands(self):
        words = self.query.text().casefold().split()
        self.results.clear()
        for label, action in self.commands:
            searchable = f'{label} {action.toolTip()}'.casefold()
            if not all(word in searchable for word in words):
                continue
            reason = str(action.property('unavailable_reason') or '')
            state = reason if not action.isEnabled() else (
                tr('已开启') if action.isCheckable() and action.isChecked() else '')
            row = QTreeWidgetItem([label, state])
            row.setData(0, Qt.UserRole, action)
            row.setToolTip(0, action.toolTip())
            row.setToolTip(1, reason)
            row.setIcon(0, action.icon())
            if not action.isEnabled():
                from PyQt5.QtGui import QPalette
                brush = self.palette().brush(QPalette.Disabled, QPalette.Text)
                row.setForeground(0, brush)
                row.setForeground(1, brush)
            self.results.addTopLevelItem(row)
        if self.results.topLevelItemCount():
            self.results.setCurrentItem(self.results.topLevelItem(0))
        else:
            self.explanation.setText(tr('没有匹配的功能'))

    def describe_current(self, *args):
        row = self.results.currentItem()
        if row is not None:
            action = row.data(0, Qt.UserRole)
            self.explanation.setText(str(action.property('unavailable_reason') or action.toolTip()))

    def choose_current(self, *args):
        row = self.results.currentItem()
        if row is None:
            return
        action = row.data(0, Qt.UserRole)
        if not action.isEnabled():
            self.describe_current()
            return
        self.selected_action = action
        self.accept()

    def eventFilter(self, watched, event):
        from PyQt5.QtCore import QEvent
        if watched is self.query and event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Down, Qt.Key_Up):
            count = self.results.topLevelItemCount()
            index = self.results.indexOfTopLevelItem(self.results.currentItem())
            if count:
                index = max(0, min(count - 1, index + (1 if event.key() == Qt.Key_Down else -1)))
                self.results.setCurrentItem(self.results.topLevelItem(index))
            return True
        return super().eventFilter(watched, event)

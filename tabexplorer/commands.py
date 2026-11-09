"""VS Code 式快速输入：查找并执行现有功能、快速打开标签/书签/常用目录、选择颜色主题。

只复用已有的界面动作，不复制执行逻辑：命令来自菜单与工具栏按钮，快速打开的条目和执行函数由主窗口提供。"""

import re

from PyQt5.QtCore import QEvent, QEventLoop, QPointF, QRectF, QSize, Qt, QTimer
from PyQt5.QtGui import QColor, QFont, QFontMetrics, QIcon, QPainter, QPainterPath
from PyQt5.QtWidgets import (
    QDialog, QFrame, QLabel, QLineEdit, QStyle, QStyledItemDelegate, QTreeWidget, QTreeWidgetItem, QVBoxLayout,
)

from . import theme as _theme
from .i18n import tr


MAX_RECENT_COMMANDS = 12
MAX_VISIBLE_ROWS = 12
MAX_QUICK_OPEN_ROWS = 300
_HOTKEY_SUFFIX = re.compile(r'^(.*?)\s*\(((?:Ctrl|Alt|Shift|F\d{1,2})(?:\+[^()]*)?)\)$')
_PATH_QUERY = re.compile(r'^(?:[A-Za-z]:[\\/]|\\\\|%[^%]+%|shell:)', re.IGNORECASE)
_WORD_BREAKS = ' /\\-_.:()[]·'
_ROLE_LABEL, _ROLE_MATCHES, _ROLE_DESCRIPTION, _ROLE_KEYS, _ROLE_GROUP, _ROLE_SEPARATOR, _ROLE_ENABLED = (
    Qt.UserRole + offset for offset in range(1, 8))

_QUICK_INPUT_CSS = """
QFrame#quickInputFrame { background: @widget-bg; border: 1px solid @widget-border; border-radius: 6px; }
QLineEdit { background: @input-bg; color: @input-fg; border: 1px solid @focus-border; border-radius: 2px;
            padding: 4px 6px; selection-background-color: @selection-bg; selection-color: @selection-fg; }
QTreeWidget { background: transparent; border: none; outline: 0; }
QTreeWidget::item, QTreeWidget::item:selected, QTreeWidget::item:hover { background: transparent; border: none; }
QLabel { color: @fg-muted; background: transparent; padding: 0px 4px 2px 4px; }
"""


def menu_commands(menu, prefix=''):
    for action in menu.actions():
        if action.isSeparator() or not action.isVisible():
            continue
        label = f'{prefix} / {action.text()}' if prefix else action.text()
        if action.menu() is not None:
            yield from menu_commands(action.menu(), label)
        else:
            yield label, action


def split_hotkey(label):
    """'查找功能 (Ctrl+Shift+P)' -> ('查找功能', 'Ctrl+Shift+P')；没有快捷键提示时第二项为空。"""
    match = _HOTKEY_SUFFIX.match(label or '')
    return (match.group(1), match.group(2)) if match else (label or '', '')


def command_key(label):
    """“最近使用”的记录键：去掉末尾的快捷键提示，改键后仍对应同一功能。"""
    return split_hotkey(label)[0].strip()


def _match_word(word, text):
    index = text.find(word)
    if index >= 0:
        boundary = index == 0 or text[index - 1] in _WORD_BREAKS
        return 100 + len(word) * 4 + (40 if boundary else 0) - min(index, 40), range(index, index + len(word))
    positions, start, gaps = [], 0, 0
    for char in word:
        found = text.find(char, start)
        if found < 0:
            return None
        if positions and found != positions[-1] + 1:
            gaps += 1
        positions.append(found)
        start = found + 1
    return 40 + len(word) * 2 - gaps * 6 - min(positions[0], 30), positions


def fuzzy_match(query, label, extra=''):
    """查询中的每个词都要命中：先在 label 中找连续子串（词首优先），再找按顺序出现的分散字符；
    label 不中时允许在 extra（路径、快捷键等）中连续命中，但不高亮。

    返回 (得分, label 中命中的字符下标)；不匹配返回 None。"""
    words = query.lower().split()
    if not words:
        return 0, []
    label_text, extra_text = label.lower(), extra.lower()
    score, positions = 0, set()
    for word in words:
        found = _match_word(word, label_text)
        if found is not None:
            score += found[0]
            positions.update(found[1])
        elif word in extra_text:
            score += 20
        else:
            return None
    return score, sorted(positions)


class _ResultDelegate(QStyledItemDelegate):
    """一行结果：图标、带命中高亮的名称、淡色说明，右侧为快捷键和分组名。"""

    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner

    def sizeHint(self, option, index):
        return QSize(option.rect.width(), self.owner.row_height)

    def paint(self, painter, option, index):
        owner = self.owner
        rect = option.rect
        selected = bool(option.state & QStyle.State_Selected)
        enabled = index.data(_ROLE_ENABLED) is not False
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        if index.data(_ROLE_SEPARATOR):
            painter.fillRect(rect.left(), rect.top(), rect.width(), 1, _theme.qcolor('picker-group-border'))
        body = QPainterPath()
        body.addRoundedRect(QRectF(rect).adjusted(0.5, 1.5, -0.5, -0.5), 3, 3)
        if selected:
            fill, border = _theme.qcolor('list-active-bg'), _theme.qcolor('list-active-border')
            painter.fillPath(body, fill)
            if border != fill:
                painter.setPen(border)
                painter.drawPath(body)
        elif option.state & QStyle.State_MouseOver:
            painter.fillPath(body, _theme.qcolor('list-hover-bg'))
        text_color = _theme.qcolor('list-active-fg' if selected else 'widget-fg')
        if not enabled:
            text_color = _theme.qcolor('fg-disabled')
        muted = QColor(text_color)
        muted.setAlphaF(muted.alphaF() * 0.72)
        pad = owner.scaled(8)
        left, right = rect.left() + pad, rect.right() - pad
        icon = index.data(Qt.DecorationRole)
        if isinstance(icon, QIcon) and not icon.isNull():
            size = owner.scaled(16)
            icon.paint(painter, left, rect.top() + (rect.height() - size) // 2, size, size, Qt.AlignCenter,
                       QIcon.Normal if enabled else QIcon.Disabled)
            left += size + owner.scaled(6)
        font = QFont(option.font)
        small = QFont(font)
        small.setPointSizeF(max(7.0, font.pointSizeF() * 0.88))
        group = index.data(_ROLE_GROUP) or ''
        if group:
            painter.setFont(small)
            painter.setPen(muted if selected else _theme.qcolor('picker-group-fg'))
            width = QFontMetrics(small).horizontalAdvance(group)
            painter.drawText(QRectF(right - width, rect.top(), width, rect.height()), Qt.AlignVCenter | Qt.AlignRight, group)
            right -= width + owner.scaled(12)
        keys = index.data(_ROLE_KEYS) or ''
        if keys:
            right = self._paint_keys(painter, keys, right, rect, small, text_color if selected else None)
        highlight = _theme.qcolor('list-active-highlight-fg' if selected else 'list-highlight-fg')
        left = self._paint_label(painter, index.data(_ROLE_LABEL) or '', set(index.data(_ROLE_MATCHES) or ()),
                                 left, right - owner.scaled(8), rect, font, text_color, highlight)
        description = index.data(_ROLE_DESCRIPTION) or ''
        room = right - left - owner.scaled(10)
        if description and room > owner.scaled(30):
            metrics = QFontMetrics(font)
            painter.setFont(font)
            painter.setPen(muted)
            painter.drawText(QRectF(left + owner.scaled(10), rect.top(), room, rect.height()),
                             Qt.AlignVCenter | Qt.AlignLeft, metrics.elidedText(description, Qt.ElideMiddle, int(room)))
        painter.restore()

    @staticmethod
    def _paint_label(painter, label, matches, left, limit, rect, font, color, highlight):
        bold = QFont(font)
        bold.setBold(True)
        runs = []
        for position, char in enumerate(label):
            hit = position in matches
            if runs and runs[-1][0] == hit:
                runs[-1][1] += char
            else:
                runs.append([hit, char])
        normal_metrics = QFontMetrics(font)
        baseline = rect.top() + (rect.height() + normal_metrics.ascent() - normal_metrics.descent()) // 2
        for hit, text in runs:
            run_font = bold if hit else font
            metrics = QFontMetrics(run_font)
            width = metrics.horizontalAdvance(text)
            clipped = left + width > limit
            if clipped:
                text = metrics.elidedText(text, Qt.ElideRight, max(0, int(limit - left)))
                width = metrics.horizontalAdvance(text)
            painter.setFont(run_font)
            painter.setPen(highlight if hit else color)
            painter.drawText(int(left), int(baseline), text)
            left += width
            if clipped:
                break
        return left

    def _paint_keys(self, painter, keys, right, rect, font, color):
        """VS Code 式按键标签：每个键一个小方框，键之间用 + 连接。"""
        owner = self.owner
        metrics = QFontMetrics(font)
        height = min(rect.height() - owner.scaled(6), metrics.height() + owner.scaled(4))
        top = rect.top() + (rect.height() - height) / 2.0
        fill, border = _theme.qcolor('keybinding-bg'), _theme.qcolor('keybinding-border')
        bottom, text_color = _theme.qcolor('keybinding-bottom-border'), color or _theme.qcolor('keybinding-fg')
        painter.setFont(font)
        parts = [part for part in keys.split('+') if part] or [keys]
        for position, part in enumerate(reversed(parts)):
            if position:
                plus = metrics.horizontalAdvance('+')
                painter.setPen(text_color)
                painter.drawText(QRectF(right - plus, rect.top(), plus, rect.height()), Qt.AlignCenter, '+')
                right -= plus + owner.scaled(2)
            width = metrics.horizontalAdvance(part) + owner.scaled(8)
            box = QRectF(right - width, top, width, height)
            path = QPainterPath()
            path.addRoundedRect(box, 3, 3)
            painter.fillPath(path, fill)
            painter.setPen(border)
            painter.drawPath(path)
            painter.setPen(bottom)
            painter.drawLine(QPointF(box.left() + 2, box.bottom()), QPointF(box.right() - 2, box.bottom()))
            painter.setPen(text_color)
            painter.drawText(box, Qt.AlignCenter, part)
            right -= width + owner.scaled(2)
        return right - owner.scaled(8)


class CommandPalette(QDialog):
    """VS Code 式快速输入浮层：贴在主窗口顶部居中，输入即筛选，Enter 选定，Esc 或点到其他地方关闭。

    - 命令（> 前缀；没有快速打开数据时不需要前缀）：菜单与工具栏中的功能，最近使用的排在最前，右侧显示快捷键，
      不可用的功能显示原因。
    - 快速打开（不带前缀）：在已开标签、书签、常用目录和最近关闭的标签中查找；输入路径可直接打开，? 查看前缀用法。
    - 颜色主题（set_theme_choices 后）：上下键选择时实时预览，Enter 确定；取消后由调用方还原。
    """

    def __init__(self, parent=None, quick_open=None, open_path=None):
        super().__init__(parent, Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setWindowTitle(tr('查找功能'))
        self._scale = max(1.0, self.logicalDpiX() / 96.0)
        self._shadow = self.scaled(12)
        self.row_height = max(self.scaled(24), self.fontMetrics().height() + self.scaled(8))
        self.selected_action = None
        self.selected_entry = None
        self.selected_label = ''
        self.commands = []
        self.recent = []
        self._quick_open = quick_open
        self._open_path = open_path
        self._entries = None
        self._themes = None
        self._mode = None
        self._rows = []
        self._content_width = self.scaled(600)

        self.frame = QFrame(self)
        self.frame.setObjectName('quickInputFrame')
        outer = QVBoxLayout(self)
        outer.setContentsMargins(self._shadow, self.scaled(4), self._shadow, self._shadow)
        outer.addWidget(self.frame)
        layout = QVBoxLayout(self.frame)
        margin = self.scaled(6)
        layout.setContentsMargins(margin, margin, margin, margin)
        layout.setSpacing(self.scaled(5))
        self.query = QLineEdit(self.frame)
        self.query.setAccessibleName(tr('快速输入'))
        layout.addWidget(self.query)
        self.results = QTreeWidget(self.frame)
        self.results.setHeaderHidden(True)
        self.results.setColumnCount(1)
        self.results.setRootIsDecorated(False)
        self.results.setIndentation(0)
        self.results.setUniformRowHeights(True)
        self.results.setFocusPolicy(Qt.NoFocus)
        self.results.setFrameShape(QFrame.NoFrame)
        self.results.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.results.setMouseTracking(True)
        self.results.setItemDelegate(_ResultDelegate(self))
        layout.addWidget(self.results)
        self.explanation = QLabel(self.frame)
        self.explanation.setWordWrap(True)
        self.explanation.setTextFormat(Qt.PlainText)
        layout.addWidget(self.explanation)
        _theme.bind_style(self.frame, _QUICK_INPUT_CSS)

        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(90)
        self._preview_timer.timeout.connect(self._preview_current)
        self.query.textChanged.connect(self.filter_commands)
        self.query.returnPressed.connect(self.choose_current)
        self.results.itemClicked.connect(lambda item, _column=0: self._choose_item(item))
        self.results.currentItemChanged.connect(self.describe_current)
        self.query.installEventFilter(self)
        self.filter_commands()

    def scaled(self, value):
        return max(1, int(round(value * self._scale)))

    @property
    def mode(self):
        return self._mode

    def set_commands(self, commands, recent=()):
        self.commands = list(commands)
        self.recent = [key for key in recent if isinstance(key, str)]
        self.filter_commands()
        self.query.setFocus()

    def set_theme_choices(self, entries, current, preview):
        """进入主题选择：entries 为 [{'id', 'label', 'group', 'description'}]，preview(id) 用于实时预览。"""
        self._themes = {'entries': list(entries), 'preview': preview}
        self.setWindowTitle(tr('选择颜色主题'))
        self.query.clear()
        self.filter_commands()
        for position in range(self.results.topLevelItemCount()):
            item = self.results.topLevelItem(position)
            if self._row(item)['data'].get('id') == current:
                self.results.setCurrentItem(item)
                self.results.scrollToItem(item)
                break
        self._preview_timer.stop()

    def _parse_query(self):
        text = self.query.text()
        if self._themes is not None:
            return 'themes', text
        if self._quick_open is None:
            return 'commands', text[1:] if text.startswith('>') else text
        if text.startswith('>'):
            return 'commands', text[1:]
        if text.startswith('?'):
            return 'help', text[1:]
        return 'files', text

    def filter_commands(self, *args):
        mode, query = self._parse_query()
        if mode != self._mode:
            self._mode = mode
            self.query.setPlaceholderText({
                'commands': tr('输入功能名称或快捷键，例如：同步、比较、工作区'),
                'files': tr('搜索标签、书签和常用目录（输入 > 查找功能，? 查看帮助）'),
                'help': '',
                'themes': tr('选择颜色主题（上下键预览）'),
            }[mode])
        builders = {'commands': self._command_rows, 'files': self._file_rows,
                    'help': self._help_rows, 'themes': self._theme_rows}
        self._populate(builders[mode](query))

    def _command_rows(self, query):
        rank_of = {key: rank for rank, key in enumerate(self.recent)}
        found = []
        for order, (label, action) in enumerate(self.commands):
            name, keys = split_hotkey(label)
            match = fuzzy_match(query, name, ' '.join((label, action.toolTip() or '', keys)))
            if match is not None:
                found.append((rank_of.get(command_key(label)), match[0], order, label, name, keys, match[1], action))
        recent = sorted((entry for entry in found if entry[0] is not None), key=lambda entry: entry[0])
        others = [entry for entry in found if entry[0] is None]
        if query.strip():
            others.sort(key=lambda entry: (-entry[1], entry[2]))
        rows = []
        for group, entries in ((tr('最近使用'), recent), (tr('其他命令') if recent else '', others)):
            for _rank, _score, _order, label, name, keys, matches, action in entries:
                enabled = action.isEnabled()
                reason = '' if enabled else str(action.property('unavailable_reason') or '')
                state = reason or (tr('已开启') if action.isCheckable() and action.isChecked() else '')
                rows.append({'data': action, 'key': label, 'label': name, 'keys': keys, 'matches': matches,
                             'group': group, 'description': state, 'enabled': enabled, 'explanation': reason,
                             'tooltip': action.toolTip()})
        return rows

    def _file_rows(self, query):
        rows = []
        path = query.strip().strip('"')
        if self._open_path is not None and _PATH_QUERY.match(path):
            rows.append({'data': {'run': lambda target=path: self._open_path(target)}, 'label': tr('打开路径'),
                         'description': path, 'group': tr('路径')})
        if self._entries is None:
            self._entries = list(self._quick_open() or [])
        groups = {}
        for position, entry in enumerate(self._entries):
            match = fuzzy_match(query, entry.get('label', ''),
                                ' '.join((entry.get('description', ''), entry.get('aliases', ''))))
            if match is not None:
                groups.setdefault(entry.get('group', ''), []).append((match[0], position, entry, match[1]))
        for group, items in groups.items():
            if query.strip():
                items.sort(key=lambda item: (-item[0], item[1]))
            for _score, _position, entry, matches in items:
                rows.append({'data': entry, 'label': entry.get('label', ''), 'matches': matches, 'group': group,
                             'description': entry.get('description', ''), 'icon': entry.get('icon'),
                             'tooltip': entry.get('description', '')})
                if len(rows) >= MAX_QUICK_OPEN_ROWS:
                    return rows
        return rows

    def _help_rows(self, query):
        rows = []
        for prefix, text in (('', tr('转到标签、书签和常用目录')), ('>', tr('显示并运行命令'))):
            if fuzzy_match(query, text) is not None:
                rows.append({'data': prefix, 'label': prefix or '…', 'description': text})
        return rows

    def _theme_rows(self, query):
        rows = []
        for entry in self._themes['entries']:
            match = fuzzy_match(query, entry['label'])
            if match is not None:
                rows.append({'data': entry, 'label': entry['label'], 'matches': match[1],
                             'group': entry.get('group', ''), 'description': entry.get('description', '')})
        return rows

    def _populate(self, rows):
        self._rows = rows
        self.results.clear()
        previous_group = None
        for position, row in enumerate(rows):
            item = QTreeWidgetItem([row['label']])
            item.setData(0, Qt.UserRole, position)
            item.setData(0, _ROLE_LABEL, row['label'])
            item.setData(0, _ROLE_MATCHES, list(row.get('matches') or ()))
            item.setData(0, _ROLE_DESCRIPTION, row.get('description', ''))
            item.setData(0, _ROLE_KEYS, row.get('keys', ''))
            item.setData(0, _ROLE_ENABLED, row.get('enabled', True))
            group = row.get('group', '')
            if group != previous_group:
                item.setData(0, _ROLE_GROUP, group)
                item.setData(0, _ROLE_SEPARATOR, position > 0)
                previous_group = group
            if isinstance(row.get('icon'), QIcon):
                item.setIcon(0, row['icon'])
            if row.get('tooltip'):
                item.setToolTip(0, row['tooltip'])
            self.results.addTopLevelItem(item)
        if rows:
            self.results.setCurrentItem(self.results.topLevelItem(0))
        else:
            self.explanation.setText({'commands': tr('没有匹配的功能'), 'files': tr('没有匹配的标签、书签或目录'),
                                      'help': tr('没有匹配的功能'), 'themes': tr('没有匹配的主题')}[self._mode])
        self._update_size()

    def _row(self, item):
        return self._rows[item.data(0, Qt.UserRole)]

    def describe_current(self, *args):
        item = self.results.currentItem()
        if item is None:
            return
        self.explanation.setText(self._row(item).get('explanation', ''))
        self._update_size()
        if self._themes is not None:
            self._preview_timer.start()

    def choose_current(self, *args):
        item = self.results.currentItem()
        if item is not None:
            self._choose_item(item)

    def _choose_item(self, item):
        row = self._row(item)
        if not row.get('enabled', True):
            self.describe_current()
            return
        if self._mode == 'help':
            self.query.setText(row['data'])
            self.query.setFocus()
            return
        if self._mode == 'commands':
            self.selected_action = row['data']
            self.selected_label = row['key']
        else:
            self.selected_entry = row['data']
        self._preview_timer.stop()
        self.accept()

    def _preview_current(self):
        item = self.results.currentItem()
        if item is not None and self._themes is not None and self.isVisible():
            self._themes['preview'](self._row(item)['data']['id'])
            self.update()

    def _move(self, step):
        count = self.results.topLevelItemCount()
        if not count:
            return
        current = self.results.currentItem()
        position = self.results.indexOfTopLevelItem(current) if current is not None else -1
        if abs(step) == 1:
            position = (position + step) % count if position >= 0 else (0 if step > 0 else count - 1)
        else:
            position = max(0, min(count - 1, position + step))
        item = self.results.topLevelItem(position)
        self.results.setCurrentItem(item)
        self.results.scrollToItem(item)

    def eventFilter(self, watched, event):
        if watched is self.query and event.type() == QEvent.KeyPress:
            key, modifiers = event.key(), event.modifiers()
            page = MAX_VISIBLE_ROWS - 1
            steps = {Qt.Key_Down: 1, Qt.Key_Up: -1, Qt.Key_PageDown: page, Qt.Key_PageUp: -page}
            if key in steps:
                self._move(steps[key])
                return True
            if (key == Qt.Key_P and modifiers & Qt.ControlModifier and self._quick_open is not None
                    and self._themes is None):
                # 与 VS Code 一致：Ctrl+Shift+P 切到命令，Ctrl+P 回到快速打开，已在快速打开时选下一项
                text = self.query.text()
                if modifiers & Qt.ShiftModifier:
                    if not text.startswith('>'):
                        self.query.setText('>' + text)
                elif text.startswith('>'):
                    self.query.setText(text[1:].lstrip())
                else:
                    self._move(1)
                return True
            if key == Qt.Key_Escape:
                self.reject()
                return True
        return super().eventFilter(watched, event)

    def place(self, anchor):
        """anchor 为主窗口内容区的全局矩形：浮层宽度约为窗口的 62%（420–640），贴在顶部居中。"""
        available = max(self.scaled(320), anchor.width() - self.scaled(24))
        self._content_width = int(min(available, max(self.scaled(420), min(self.scaled(640), anchor.width() * 0.62))))
        self._update_size()
        self.move(anchor.center().x() - self.width() // 2, anchor.top() - self.scaled(2))

    def _update_size(self):
        rows = min(self.results.topLevelItemCount(), MAX_VISIBLE_ROWS)
        self.results.setVisible(rows > 0)
        self.results.setFixedHeight(rows * self.row_height + 2)
        self.explanation.setVisible(bool(self.explanation.text()))
        self.setFixedWidth(self._content_width + 2 * self._shadow)
        self.adjustSize()

    def run(self, anchor=None):
        """显示并等待关闭；不阻塞其他窗口，点到别处即关闭。返回是否选定了一项。"""
        if anchor is not None:
            self.place(anchor)
        loop = QEventLoop()
        self.finished.connect(loop.quit)
        self.show()
        self.raise_()
        self.activateWindow()
        self.query.setFocus(Qt.ActiveWindowFocusReason)
        if self.isVisible():
            loop.exec_()
        return self.result() == QDialog.Accepted

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QEvent.ActivationChange and self.isVisible() and not self.isActiveWindow():
            QTimer.singleShot(0, self._dismiss_if_inactive)

    def _dismiss_if_inactive(self):
        if self.isVisible() and not self.isActiveWindow():
            self.reject()

    def paintEvent(self, event):
        shadow = _theme.qcolor('widget-shadow')
        if not shadow.isValid() or shadow.alpha() == 0:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        frame = QRectF(self.frame.geometry())
        steps = self._shadow
        for step in range(1, steps + 1):
            layer = QColor(shadow)
            layer.setAlphaF(min(1.0, shadow.alphaF() * 0.14 * (1.0 - step / (steps + 1.0))))
            painter.setBrush(layer)
            painter.drawRoundedRect(frame.adjusted(-step, -step * 0.4, step, step * 1.2), 6 + step, 6 + step)
        painter.end()

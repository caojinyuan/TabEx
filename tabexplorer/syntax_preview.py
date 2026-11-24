"""Syntax highlighting for bounded, read-only previews."""

import os

from PyQt5.QtCore import QRect, QSize, Qt
from PyQt5.QtGui import QColor, QFont, QPainter, QSyntaxHighlighter, QTextCharFormat
from PyQt5.QtWidgets import QPlainTextEdit, QWidget
from pygments import lex
from pygments.lexers import get_lexer_by_name, get_lexer_for_filename
from pygments.token import Comment, Keyword, Literal, Name, Operator
from pygments.util import ClassNotFound

from . import theme

HIGHLIGHT_LIMIT = 512 * 1024


def highlight_spans(path, text):
    if len(text) > HIGHLIGHT_LIMIT:
        return None
    try:
        if os.path.splitext(path)[1].lower() in ('.c', '.h'):
            lexer = get_lexer_by_name('c', stripnl=False, ensurenl=False)
        else:
            lexer = get_lexer_for_filename(path, stripnl=False, ensurenl=False)
    except ClassNotFound:
        return None
    lines = {}
    line = offset = spans = 0
    for token, value in lex(text, lexer):
        if token in Comment.Preproc:
            category = 'preprocessor'
        elif token in Comment:
            category = 'comment'
        elif token in Keyword:
            category = 'keyword'
        elif token in Literal.String:
            category = 'string'
        elif token in Literal.Number:
            category = 'number'
        elif token in Name.Function or token in Name.Class:
            category = 'name'
        elif token in Operator:
            category = 'operator'
        else:
            category = None
        parts = value.split('\n')
        for index, part in enumerate(parts):
            length = len(part.encode('utf-16-le')) // 2
            if category and length:
                lines.setdefault(line, []).append((offset, length, category))
                spans += 1
                if spans > 100000:
                    return None
            offset += length
            if index < len(parts) - 1:
                line += 1
                offset = 0
    return lines


class CodeHighlighter(QSyntaxHighlighter):
    def __init__(self, document, spans):
        super().__init__(document)
        self._spans = spans
        colors = {
            'preprocessor': '#AF00DB', 'comment': '#008000', 'keyword': '#0000FF',
            'string': '#A31515', 'number': '#098658', 'name': '#795E26', 'operator': '#000000',
        }
        self._formats = {}
        for name, color in colors.items():
            fmt = QTextCharFormat()
            fmt.setForeground(QColor(theme.fg(color)))
            if name in ('keyword', 'preprocessor'):
                fmt.setFontWeight(QFont.Bold)
            self._formats[name] = fmt

    def highlightBlock(self, text):
        for start, length, category in self._spans.get(self.currentBlock().blockNumber(), ()):
            self.setFormat(start, length, self._formats[category])


class _LineNumbers(QWidget):
    def __init__(self, editor):
        super().__init__(editor)
        self.editor = editor

    def sizeHint(self):
        return QSize(self.editor.line_number_width(), 0)

    def paintEvent(self, event):
        self.editor.paint_line_numbers(event)


class PreviewTextEdit(QPlainTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._show_numbers = False
        self._numbers = _LineNumbers(self)
        self._numbers.hide()
        self.blockCountChanged.connect(self._update_margin)
        self.updateRequest.connect(self._update_numbers)

    def set_line_numbers_visible(self, visible):
        self._show_numbers = visible
        self._numbers.setVisible(visible)
        self._update_margin()

    def line_number_width(self):
        if not self._show_numbers:
            return 0
        digits = len(str(max(1, self.blockCount())))
        return 16 + self.fontMetrics().horizontalAdvance('9') * digits

    def _update_margin(self, *_):
        width = self.line_number_width()
        self.setViewportMargins(width, 0, 0, 0)
        rect = self.contentsRect()
        self._numbers.setGeometry(QRect(rect.left(), rect.top(), width, rect.height()))

    def _update_numbers(self, rect, dy):
        if dy:
            self._numbers.scroll(0, dy)
        else:
            self._numbers.update(0, rect.y(), self._numbers.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self._update_margin()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        rect = self.contentsRect()
        self._numbers.setGeometry(QRect(rect.left(), rect.top(), self.line_number_width(), rect.height()))

    def paint_line_numbers(self, event):
        painter = QPainter(self._numbers)
        painter.fillRect(event.rect(), theme.qcolor('chrome-bg'))
        painter.setPen(theme.qcolor('fg-muted'))
        block = self.firstVisibleBlock()
        top = round(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        while block.isValid() and top <= event.rect().bottom():
            height = round(self.blockBoundingRect(block).height())
            if block.isVisible() and top + height >= event.rect().top():
                painter.drawText(0, top, self._numbers.width() - 8, self.fontMetrics().height(),
                                 Qt.AlignRight, str(block.blockNumber() + 1))
            block = block.next()
            top += height

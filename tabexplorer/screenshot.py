"""屏幕截图：区域选择（窗口/控件识别）、标注、复制/保存/贴图与取色器。"""

import ctypes
import ctypes.wintypes as wt
import math
import os
from datetime import datetime

from PyQt5.QtCore import (
    QAbstractNativeEventFilter, QObject, QPoint, QPointF, QRect, QRectF, QStandardPaths, Qt, QTimer, pyqtSignal,
)
from PyQt5.QtGui import (
    QColor, QCursor, QFont, QFontMetrics, QFontMetricsF, QGuiApplication, QImage, QPainter, QPainterPath,
    QPainterPathStroker, QPen, QPixmap, QPolygonF,
)
from PyQt5.QtWidgets import QApplication, QFileDialog, QFrame, QHBoxLayout, QLineEdit, QMenu, QToolButton, QVBoxLayout, QWidget

from .debuglog import debug_print
from .i18n import tr
from .menus import exec_menu


MIN_SELECTION = 2
HANDLE_NAMES = ('tl', 't', 'tr', 'r', 'br', 'b', 'bl', 'l')
HANDLE_CURSORS = {
    'tl': Qt.SizeFDiagCursor, 'br': Qt.SizeFDiagCursor, 'tr': Qt.SizeBDiagCursor, 'bl': Qt.SizeBDiagCursor,
    't': Qt.SizeVerCursor, 'b': Qt.SizeVerCursor, 'l': Qt.SizeHorCursor, 'r': Qt.SizeHorCursor,
}
ACCENT = QColor('#2f8cff')
PALETTE = ('#FF3B30', '#FF9500', '#FFCC00', '#34C759', '#0A84FF', '#FFFFFF', '#000000')
WIDTHS = (2, 4, 6)
TOOLS = (
    ('rect', '▭', '矩形'), ('ellipse', '◯', '椭圆'), ('arrow', '↗', '箭头'), ('pen', '✎', '画笔'),
    ('mosaic', '▦', '马赛克'), ('text', 'A', '文字'), ('number', '①', '序号'),
)
ACTIONS = (
    ('undo', '↶', '撤销 (Ctrl+Z)'), ('pin', '📌', '贴图 (F3)'), ('save', '💾', '保存 (Ctrl+S)'),
    ('cancel', '✕', '取消 (Esc)'), ('copy', '✓', '复制到剪贴板 (Enter)'),
)
MOSAIC_BLOCK = 10
MAGNIFIER_GRID = 15
MAGNIFIER_ZOOM = 8

_pins = []
_last_save_dir = ''


# ---------------------------------------------------------------- Win32 窗口枚举

def _top_level_windows():
    """自上而下（Z 序）返回 [(hwnd, QRect 物理像素)]，跳过隐藏、最小化、被 DWM 隐藏和点击穿透的窗口。"""
    if os.name != 'nt':
        return []
    user32, dwmapi = ctypes.windll.user32, ctypes.windll.dwmapi
    user32.EnumWindows.argtypes = (ctypes.c_void_p, wt.LPARAM)
    user32.GetClassNameW.argtypes = (wt.HWND, wt.LPWSTR, ctypes.c_int)
    skip_classes = {'Progman', 'WorkerW'}
    result = []

    @ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
    def callback(hwnd, _lparam):
        try:
            if not user32.IsWindowVisible(hwnd) or user32.IsIconic(hwnd):
                return True
            cloaked = ctypes.c_int(0)
            if dwmapi.DwmGetWindowAttribute(hwnd, 14, ctypes.byref(cloaked), 4) == 0 and cloaked.value:
                return True
            ex_style = user32.GetWindowLongW(hwnd, -20)
            if ex_style & 0x20 and ex_style & 0x80000:  # WS_EX_TRANSPARENT + WS_EX_LAYERED：点击穿透的覆盖层
                return True
            buffer = ctypes.create_unicode_buffer(64)
            user32.GetClassNameW(hwnd, buffer, 64)
            if buffer.value in skip_classes:
                return True
            rect = wt.RECT()
            if dwmapi.DwmGetWindowAttribute(hwnd, 9, ctypes.byref(rect), ctypes.sizeof(rect)) != 0:
                if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                    return True
            if rect.right - rect.left > 1 and rect.bottom - rect.top > 1:
                result.append((int(hwnd), QRect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)))
        except Exception:
            pass
        return True

    user32.EnumWindows(callback, 0)
    return result


def _child_rects(hwnd, x, y):
    """从顶层窗口向下逐级查找包含物理坐标 (x, y) 的子控件，返回各级 QRect（不含顶层本身）。"""
    if os.name != 'nt':
        return []
    user32 = ctypes.windll.user32
    user32.ChildWindowFromPointEx.argtypes = (wt.HWND, wt.POINT, wt.UINT)
    user32.ChildWindowFromPointEx.restype = wt.HWND
    rects, current = [], hwnd
    for _ in range(8):
        point = wt.POINT(x, y)
        user32.ScreenToClient(wt.HWND(current), ctypes.byref(point))
        child = user32.ChildWindowFromPointEx(wt.HWND(current), point, 0x1 | 0x4)  # SKIPINVISIBLE | SKIPTRANSPARENT
        if not child or int(child) == current:
            break
        rect = wt.RECT()
        if not user32.GetWindowRect(wt.HWND(child), ctypes.byref(rect)):
            break
        rects.append(QRect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top))
        current = int(child)
    return rects


class _ScreenMap:
    """屏幕物理像素与叠加层逻辑坐标的换算；Qt5 中 QScreen 几何的左上角是原生物理坐标，尺寸是逻辑像素。"""

    def __init__(self, geometry, dpr):
        self.geometry = geometry
        self.dpr = dpr or 1.0

    def rect_to_widget(self, rect):
        gx, gy, dpr = self.geometry.x(), self.geometry.y(), self.dpr
        left, top = round((rect.left() - gx) / dpr), round((rect.top() - gy) / dpr)
        right, bottom = round((rect.right() + 1 - gx) / dpr), round((rect.bottom() + 1 - gy) / dpr)
        return QRect(left, top, max(1, right - left), max(1, bottom - top))

    def point_to_physical(self, point):
        return (round(self.geometry.x() + point.x() * self.dpr), round(self.geometry.y() + point.y() * self.dpr))


# ---------------------------------------------------------------- 标注

class _Item:
    def __init__(self, kind, color, width, point):
        self.kind = kind
        self.color = QColor(color)
        self.width = width
        self.p1 = QPointF(point)
        self.p2 = QPointF(point)
        self.points = [QPointF(point)]
        self.text = ''
        self.number = 0


def _rect_between(a, b):
    """两点之间的矩形，宽高等于坐标差（QRect(QPoint, QPoint) 会多算 1 像素）。"""
    return QRect(min(a.x(), b.x()), min(a.y(), b.y()), abs(a.x() - b.x()), abs(a.y() - b.y()))


def _contrast(color):
    return QColor('#000000') if color.lightness() > 170 else QColor('#FFFFFF')


def _text_pixels(width):
    return 10 + width * 3


def _text_font(width):
    font = QFont('Microsoft YaHei UI')
    font.setPixelSize(_text_pixels(width))
    return font


def _polyline(points):
    path = QPainterPath(points[0])
    for point in points[1:]:
        path.lineTo(point)
    if len(points) == 1:
        path.lineTo(QPointF(points[0].x() + 0.01, points[0].y()))
    return path


def _paint_arrow(painter, item):
    a, b = item.p1, item.p2
    dx, dy = b.x() - a.x(), b.y() - a.y()
    length = math.hypot(dx, dy)
    if length < 1:
        return
    ux, uy = dx / length, dy / length
    head = min(length, max(12.0, item.width * 4.5))
    half = head * 0.45
    base = QPointF(b.x() - ux * head, b.y() - uy * head)
    painter.setPen(QPen(item.color, item.width, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
    painter.drawLine(a, base)
    painter.setBrush(item.color)
    painter.drawPolygon(QPolygonF([b, QPointF(base.x() - uy * half, base.y() + ux * half),
                                   QPointF(base.x() + uy * half, base.y() - ux * half)]))


def _paint_item(painter, item, mosaic):
    """painter 使用叠加层逻辑坐标；最终输出时由调用方设置平移/缩放。"""
    painter.save()
    painter.setRenderHint(QPainter.Antialiasing, True)
    kind = item.kind
    if kind in ('rect', 'ellipse'):
        painter.setPen(QPen(item.color, item.width, Qt.SolidLine, Qt.SquareCap, Qt.MiterJoin))
        painter.setBrush(Qt.NoBrush)
        rect = QRectF(item.p1, item.p2).normalized()
        painter.drawRect(rect) if kind == 'rect' else painter.drawEllipse(rect)
    elif kind == 'arrow':
        _paint_arrow(painter, item)
    elif kind == 'pen':
        painter.setPen(QPen(item.color, item.width, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        painter.setBrush(Qt.NoBrush)
        painter.drawPath(_polyline(item.points))
    elif kind == 'mosaic':
        if mosaic is not None:
            stroker = QPainterPathStroker()
            stroker.setWidth(item.width * 5 + 6)
            stroker.setCapStyle(Qt.RoundCap)
            stroker.setJoinStyle(Qt.RoundJoin)
            painter.setClipPath(stroker.createStroke(_polyline(item.points)), Qt.IntersectClip)
            painter.drawPixmap(0, 0, mosaic)
    elif kind == 'text':
        font = _text_font(item.width)
        painter.setFont(font)
        painter.setPen(item.color)
        painter.drawText(QPointF(item.p1.x(), item.p1.y() + QFontMetricsF(font).ascent()), item.text)
    elif kind == 'number':
        radius = 9 + item.width * 2
        painter.setPen(Qt.NoPen)
        painter.setBrush(item.color)
        painter.drawEllipse(item.p1, radius, radius)
        font = QFont('Segoe UI')
        font.setPixelSize(int(radius * 1.15))
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(_contrast(item.color))
        painter.drawText(QRectF(item.p1.x() - radius, item.p1.y() - radius, radius * 2, radius * 2),
                         Qt.AlignCenter, str(item.number))
    painter.restore()


# ---------------------------------------------------------------- 工具栏

class _Toolbar(QFrame):
    toolChanged = pyqtSignal(str)
    colorChanged = pyqtSignal(QColor)
    widthChanged = pyqtSignal(int)
    actionTriggered = pyqtSignal(str)

    def __init__(self, parent, color, width):
        super().__init__(parent)
        self.setObjectName('screenshotToolbar')
        self.setAttribute(Qt.WA_NoMousePropagation, True)
        self.setCursor(Qt.ArrowCursor)
        self.setStyleSheet(
            "#screenshotToolbar { background: #ffffff; border: 1px solid #c8c8c8; border-radius: 4px; }"
            "QToolButton { border: 1px solid transparent; border-radius: 3px; color: #202020; font-size: 15px; "
            "font-family: 'Segoe UI Symbol', 'Segoe UI Emoji', 'Microsoft YaHei UI'; }"
            "QToolButton:hover { background: #e8f0fe; }"
            "QToolButton:checked { background: #cfe3ff; border-color: #2f8cff; }"
            "QToolButton:disabled { color: #b0b0b0; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)
        self._options = QWidget(self)
        options_layout = QHBoxLayout(self._options)
        options_layout.setContentsMargins(0, 0, 0, 0)
        options_layout.setSpacing(3)
        self._color_buttons, self._width_buttons, self._tool_buttons = [], [], {}
        for value in PALETTE:
            button = self._button('', '', checkable=True)
            button.setFixedSize(22, 22)
            button.setStyleSheet(f"QToolButton {{ background: {value}; border: 1px solid #909090; border-radius: 3px; }}"
                                 "QToolButton:checked { border: 2px solid #2f8cff; }")
            button.setChecked(value.lower() == color.name().lower())
            button.clicked.connect(lambda _checked, v=value: self._pick_color(v))
            self._color_buttons.append(button)
            options_layout.addWidget(button)
        options_layout.addSpacing(8)
        for size, value in zip((8, 12, 16), WIDTHS):
            button = self._button('●', '', checkable=True)
            button.setStyleSheet(f"QToolButton {{ font-size: {size}px; }}")
            button.setChecked(value == width)
            button.clicked.connect(lambda _checked, v=value: self._pick_width(v))
            self._width_buttons.append(button)
            options_layout.addWidget(button)
        options_layout.addStretch(1)
        layout.addWidget(self._options)
        row = QHBoxLayout()
        row.setSpacing(2)
        for name, glyph, label in TOOLS:
            button = self._button(glyph, label, checkable=True)
            button.clicked.connect(lambda checked, n=name: self._pick_tool(n, checked))
            self._tool_buttons[name] = button
            row.addWidget(button)
        row.addSpacing(8)
        self._action_buttons = {}
        for name, glyph, label in ACTIONS:
            button = self._button(glyph, label)
            button.clicked.connect(lambda _checked, n=name: self.actionTriggered.emit(n))
            self._action_buttons[name] = button
            row.addWidget(button)
        self._action_buttons['undo'].setEnabled(False)
        layout.addLayout(row)
        self._options.setVisible(False)
        self.adjustSize()

    def _button(self, text, tooltip, checkable=False):
        button = QToolButton(self)
        button.setText(text)
        button.setToolTip(tr(tooltip) if tooltip else '')
        button.setCheckable(checkable)
        button.setFixedSize(30, 30)
        button.setFocusPolicy(Qt.NoFocus)
        return button

    def _pick_tool(self, name, checked):
        for other, button in self._tool_buttons.items():
            button.setChecked(checked and other == name)
        self._options.setVisible(checked)
        self.adjustSize()
        self.toolChanged.emit(name if checked else '')

    def _pick_color(self, value):
        for button, candidate in zip(self._color_buttons, PALETTE):
            button.setChecked(candidate == value)
        self.colorChanged.emit(QColor(value))

    def _pick_width(self, value):
        for button, candidate in zip(self._width_buttons, WIDTHS):
            button.setChecked(candidate == value)
        self.widthChanged.emit(value)

    def set_can_undo(self, enabled):
        self._action_buttons['undo'].setEnabled(enabled)


class _TextInput(QLineEdit):
    committed = pyqtSignal()
    cancelled = pyqtSignal()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.committed.emit()
        elif event.key() == Qt.Key_Escape:
            self.cancelled.emit()
        else:
            super().keyPressEvent(event)

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self.committed.emit()


# ---------------------------------------------------------------- 叠加层

class ScreenshotOverlay(QWidget):
    closed = pyqtSignal()

    def __init__(self, controller, screen, shot, windows):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self._controller = controller
        self._map = _ScreenMap(screen.geometry(), screen.devicePixelRatio())
        self._shot = shot
        self._image = shot.toImage()
        self._windows = windows
        self._mosaic = None
        self._sel = None
        self._drag_rect = None
        self._chain, self._chain_index = [], 0
        self._mode = None
        self._handle = None
        self._press_pos = QPoint()
        self._orig = QRect()
        self._cursor_pos = QPoint(-1, -1)
        self._items = []
        self._draft = None
        self._tool = ''
        self._color = QColor(PALETTE[0])
        self._width = WIDTHS[1]
        self._toolbar = None
        self._text_edit = None
        self._hex_format = True
        self._flash = ''
        self._flash_timer = QTimer(self)
        self._flash_timer.setSingleShot(True)
        self._flash_timer.timeout.connect(self._clear_flash)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setAttribute(Qt.WA_QuitOnClose, False)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setCursor(Qt.CrossCursor)
        self.setGeometry(self._map.geometry)

    def begin(self):
        self.show()
        self.setGeometry(self._map.geometry)
        self.raise_()
        self.activateWindow()
        self.setFocus()
        try:
            ctypes.windll.user32.SetForegroundWindow(wt.HWND(int(self.winId())))
        except Exception:
            pass
        self._cursor_pos = self.mapFromGlobal(QCursor.pos())
        self._update_hover()
        self.update()

    # -- 几何辅助
    def _display_rect(self):
        if self._sel is not None:
            return self._sel
        if self._drag_rect is not None:
            return self._drag_rect
        if self._chain:
            return self._chain[self._chain_index]
        return None

    def _update_hover(self):
        x, y = self._map.point_to_physical(self._cursor_pos)
        chain = []
        for hwnd, rect in self._windows:
            if rect.contains(x, y):
                chain = [rect] + _child_rects(hwnd, x, y)
                break
        bounds = self.rect()
        widget_rects = []
        for rect in chain:
            widget_rect = self._map.rect_to_widget(rect).intersected(bounds)
            if not widget_rect.isEmpty() and (not widget_rects or widget_rects[-1] != widget_rect):
                widget_rects.append(widget_rect)
        self._chain = widget_rects or [bounds]
        self._chain_index = len(self._chain) - 1

    def _handle_points(self):
        r = self._sel
        cx, cy = (r.left() + r.right() + 1) / 2, (r.top() + r.bottom() + 1) / 2
        left, top, right, bottom = r.left(), r.top(), r.right() + 1, r.bottom() + 1
        return {'tl': QPointF(left, top), 't': QPointF(cx, top), 'tr': QPointF(right, top), 'r': QPointF(right, cy),
                'br': QPointF(right, bottom), 'b': QPointF(cx, bottom), 'bl': QPointF(left, bottom),
                'l': QPointF(left, cy)}

    def _hit_test(self, pos):
        if self._sel is None:
            return None
        for name, point in self._handle_points().items():
            if abs(point.x() - pos.x()) <= 6 and abs(point.y() - pos.y()) <= 6:
                return name
        return 'move' if self._sel.contains(pos) else None

    def _update_cursor(self, pos):
        hit = self._hit_test(pos)
        if hit in HANDLE_CURSORS:
            self.setCursor(HANDLE_CURSORS[hit])
        elif hit == 'move':
            self.setCursor(Qt.CrossCursor if self._tool else Qt.SizeAllCursor)
        else:
            self.setCursor(Qt.ArrowCursor if self._sel is not None else Qt.CrossCursor)

    # -- 绘制
    def _mosaic_pixmap(self):
        if self._mosaic is None:
            width, height = self._image.width(), self._image.height()
            small = self._image.scaled(max(1, width // MOSAIC_BLOCK), max(1, height // MOSAIC_BLOCK),
                                       Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
            pixmap = QPixmap.fromImage(small.scaled(width, height, Qt.IgnoreAspectRatio, Qt.FastTransformation))
            pixmap.setDevicePixelRatio(self._map.dpr)
            self._mosaic = pixmap
        return self._mosaic

    def _paint_items(self, painter):
        mosaic = self._mosaic_pixmap() if any(i.kind == 'mosaic' for i in self._items + ([self._draft] if self._draft else [])) else None
        for item in self._items:
            _paint_item(painter, item, mosaic)
        if self._draft is not None:
            _paint_item(painter, self._draft, mosaic)

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self._shot)
        rect = self._display_rect()
        dim = QPainterPath()
        dim.addRect(QRectF(self.rect()))
        if rect is not None:
            hole = QPainterPath()
            hole.addRect(QRectF(rect))
            dim = dim.subtracted(hole)
        painter.fillPath(dim, QColor(0, 0, 0, 110))
        if rect is not None:
            painter.save()
            painter.setClipRect(rect)
            self._paint_items(painter)
            painter.restore()
            painter.setPen(QPen(ACCENT, 2))
            painter.setBrush(Qt.NoBrush)
            painter.drawRect(QRectF(rect).adjusted(-1, -1, 1, 1))
            self._paint_size_label(painter, rect)
            if self._sel is not None:
                painter.setPen(QPen(ACCENT, 1))
                painter.setBrush(QColor('#ffffff'))
                for point in self._handle_points().values():
                    painter.drawRect(QRectF(point.x() - 3.5, point.y() - 3.5, 7, 7))
        if self._sel is None:
            self._paint_hint(painter)
            if self._mode in (None, 'select') and self._text_edit is None:
                self._paint_magnifier(painter)
        painter.end()

    def _paint_label(self, painter, text, origin):
        font = painter.font()
        font.setPixelSize(13)
        painter.setFont(font)
        metrics = QFontMetrics(font)
        box = QRect(origin.x(), origin.y(), metrics.horizontalAdvance(text) + 12, metrics.height() + 6)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(0, 0, 0, 190))
        painter.drawRoundedRect(box, 3, 3)
        painter.setPen(QColor('#ffffff'))
        painter.drawText(box, Qt.AlignCenter, text)

    def _paint_size_label(self, painter, rect):
        dpr = self._map.dpr
        text = f"{round(rect.width() * dpr)} × {round(rect.height() * dpr)}"
        y = rect.top() - 28
        self._paint_label(painter, text, QPoint(max(0, rect.left()), y if y >= 0 else rect.top() + 4))

    def _paint_hint(self, painter):
        text = tr("拖拽选择区域 · 单击选中窗口/控件 · 滚轮切换层级 · C 复制颜色 · Esc 退出")
        font = painter.font()
        font.setPixelSize(13)
        width = QFontMetrics(font).horizontalAdvance(text) + 12
        self._paint_label(painter, text, QPoint(max(0, (self.width() - width) // 2), 12))

    def _color_at(self, physical):
        if 0 <= physical[0] < self._image.width() and 0 <= physical[1] < self._image.height():
            return self._image.pixelColor(physical[0], physical[1])
        return None

    def _color_text(self, color):
        if color is None:
            return ''
        if self._hex_format:
            return color.name().upper()
        return f"RGB({color.red()}, {color.green()}, {color.blue()})"

    def _paint_magnifier(self, painter):
        pos = self._cursor_pos
        if pos.x() < 0 or not self.rect().contains(pos):
            return
        dpr = self._map.dpr
        px, py = round(pos.x() * dpr), round(pos.y() * dpr)
        size = MAGNIFIER_GRID * MAGNIFIER_ZOOM
        box_h = size + 62
        x = pos.x() + 22 if pos.x() + 22 + size <= self.width() else pos.x() - 22 - size
        y = pos.y() + 22 if pos.y() + 22 + box_h <= self.height() else pos.y() - 22 - box_h
        painter.save()
        painter.setRenderHint(QPainter.SmoothPixmapTransform, False)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(30, 30, 30, 230))
        painter.drawRect(QRect(x - 2, y - 2, size + 4, box_h + 4))
        half = MAGNIFIER_GRID // 2
        source = self._image.copy(px - half, py - half, MAGNIFIER_GRID, MAGNIFIER_GRID)
        painter.drawImage(QRect(x, y, size, size), source)
        painter.setPen(QPen(QColor(47, 140, 255, 200), 1))
        painter.setBrush(Qt.NoBrush)
        centre = half * MAGNIFIER_ZOOM
        painter.drawRect(QRect(x + centre, y + centre, MAGNIFIER_ZOOM, MAGNIFIER_ZOOM))
        color = self._color_at((px, py))
        painter.setPen(QColor('#ffffff'))
        font = painter.font()
        font.setPixelSize(12)
        painter.setFont(font)
        lines = [f"({px}, {py})", self._color_text(color),
                 self._flash or tr("C 复制颜色 · Shift 切换格式")]
        for index, line in enumerate(lines):
            painter.drawText(QRect(x + 4, y + size + 3 + index * 18, size - 8, 18), Qt.AlignVCenter, line)
        if color is not None:
            painter.setBrush(color)
            painter.setPen(QPen(QColor('#ffffff'), 1))
            painter.drawRect(QRect(x + size - 20, y + size + 24, 14, 14))
        painter.restore()

    # -- 鼠标
    def mousePressEvent(self, event):
        pos = event.pos()
        self.setFocus()
        if event.button() == Qt.RightButton:
            self._reset_or_close()
            return
        if event.button() != Qt.LeftButton:
            return
        if self._sel is None:
            self._press_pos, self._mode, self._drag_rect = pos, 'select', None
            return
        hit = self._hit_test(pos)
        if hit in HANDLE_CURSORS:
            self._mode, self._handle, self._orig = 'resize', hit, QRect(self._sel)
        elif hit == 'move':
            if self._tool:
                self._start_draw(pos)
            else:
                self._mode, self._press_pos, self._orig = 'move', pos, QRect(self._sel)

    def mouseDoubleClickEvent(self, event):
        if (event.button() == Qt.LeftButton and self._sel is not None and not self._tool
                and self._sel.contains(event.pos())):
            self._perform('copy')
        else:
            self.mousePressEvent(event)

    def mouseMoveEvent(self, event):
        pos = event.pos()
        self._cursor_pos = pos
        bounds = self.rect()
        if self._mode == 'select':
            if self._drag_rect is not None or (pos - self._press_pos).manhattanLength() > 4:
                self._drag_rect = _rect_between(self._press_pos, pos).intersected(bounds)
        elif self._mode == 'move':
            moved = self._orig.translated(pos - self._press_pos)
            moved.moveLeft(max(0, min(moved.left(), bounds.width() - moved.width())))
            moved.moveTop(max(0, min(moved.top(), bounds.height() - moved.height())))
            self._sel = moved
            self._place_toolbar()
        elif self._mode == 'resize':
            self._resize(pos)
            self._place_toolbar()
        elif self._mode == 'draw':
            self._extend_draft(pos)
        elif self._sel is None:
            self._update_hover()
        else:
            self._update_cursor(pos)
        self.update()

    def mouseReleaseEvent(self, event):
        if event.button() != Qt.LeftButton:
            return
        mode, self._mode = self._mode, None
        if mode == 'select':
            candidate = self._drag_rect
            if candidate is None and self._chain:
                candidate = self._chain[self._chain_index]
            self._drag_rect = None
            if candidate is not None and candidate.width() >= MIN_SELECTION and candidate.height() >= MIN_SELECTION:
                self._sel = QRect(candidate)
                self._show_toolbar()
        elif mode == 'draw':
            self._commit_draft()
        elif mode in ('move', 'resize'):
            self._place_toolbar()
        self.update()

    def wheelEvent(self, event):
        if self._sel is None and self._chain and self._mode is None:
            step = -1 if event.angleDelta().y() > 0 else 1
            self._chain_index = max(0, min(len(self._chain) - 1, self._chain_index + step))
            self.update()

    def _resize(self, pos):
        bounds = self.rect()
        x = max(0, min(pos.x(), bounds.width()))
        y = max(0, min(pos.y(), bounds.height()))
        rect = QRect(self._orig)
        if 'l' in self._handle:
            rect.setLeft(x)
        if 'r' in self._handle:
            rect.setRight(x - 1)
        if 't' in self._handle:
            rect.setTop(y)
        if 'b' in self._handle:
            rect.setBottom(y - 1)
        self._sel = rect.normalized()

    # -- 键盘
    def keyPressEvent(self, event):
        key, mods = event.key(), event.modifiers()
        ctrl = bool(mods & Qt.ControlModifier)
        if key == Qt.Key_Escape:
            self.close()
        elif key in (Qt.Key_Return, Qt.Key_Enter) and self._sel is not None:
            self._perform('copy')
        elif ctrl and key == Qt.Key_Z:
            self._undo()
        elif ctrl and key == Qt.Key_C and self._sel is not None:
            self._perform('copy')
        elif ctrl and key == Qt.Key_S and self._sel is not None:
            self._perform('save')
        elif key == Qt.Key_F3 and self._sel is not None:
            self._perform('pin')
        elif self._sel is None and key == Qt.Key_Shift:
            self._hex_format = not self._hex_format
            self.update()
        elif self._sel is None and key == Qt.Key_C and not ctrl:
            self._copy_color()
        else:
            super().keyPressEvent(event)

    def _copy_color(self):
        text = self._color_text(self._color_at(self._map.point_to_physical(self._cursor_pos)))
        if text:
            QApplication.clipboard().setText(text)
            self._flash = tr("已复制 {}").format(text)
            self._flash_timer.start(1500)
            self.update()

    def _clear_flash(self):
        self._flash = ''
        self.update()

    # -- 标注
    def _start_draw(self, pos):
        if self._tool == 'text':
            self._begin_text(pos)
            return
        item = _Item(self._tool, self._color, self._width, pos)
        if self._tool == 'number':
            item.number = 1 + sum(1 for other in self._items if other.kind == 'number')
            self._items.append(item)
            self._sync_undo()
            self.update()
            return
        self._draft, self._mode = item, 'draw'

    def _extend_draft(self, pos):
        draft = self._draft
        if draft is None:
            return
        draft.p2 = QPointF(pos)
        if draft.kind in ('pen', 'mosaic'):
            draft.points.append(QPointF(pos))

    def _commit_draft(self):
        draft, self._draft = self._draft, None
        if draft is None:
            return
        if draft.kind in ('rect', 'ellipse', 'arrow') and math.hypot(draft.p2.x() - draft.p1.x(),
                                                                       draft.p2.y() - draft.p1.y()) < 3:
            return
        self._items.append(draft)
        self._sync_undo()

    def _undo(self):
        if self._text_edit is not None:
            return
        if self._items:
            self._items.pop()
            self._sync_undo()
            self.update()

    def _sync_undo(self):
        if self._toolbar is not None:
            self._toolbar.set_can_undo(bool(self._items))

    def _begin_text(self, pos):
        self._finish_text(True)
        font = _text_font(self._width)
        edit = _TextInput(self)
        edit.setFont(font)
        edit.setFixedHeight(QFontMetrics(font).height() + 2)
        edit.setStyleSheet(f"QLineEdit {{ background: rgba(255,255,255,70); color: {self._color.name()}; "
                           f"border: 1px dashed {self._color.name()}; padding: 0px; margin: 0px; }}")
        edit.setFixedWidth(80)
        edit.move(pos)
        edit.textChanged.connect(lambda text, e=edit: e.setFixedWidth(
            min(max(80, QFontMetrics(e.font()).horizontalAdvance(text) + 14), max(80, self.width() - e.x()))))
        edit.committed.connect(lambda: self._finish_text(True))
        edit.cancelled.connect(lambda: self._finish_text(False))
        self._text_edit = edit
        edit.show()
        edit.setFocus()

    def _finish_text(self, commit):
        edit, self._text_edit = self._text_edit, None
        if edit is None:
            return
        text = edit.text().strip()
        if commit and text:
            item = _Item('text', self._color, self._width, edit.pos() + QPoint(2, 1))
            item.text = text
            item.color = QColor(self._color)
            self._items.append(item)
            self._sync_undo()
        edit.hide()
        edit.deleteLater()
        self.setFocus()
        self.update()

    # -- 工具栏
    def _show_toolbar(self):
        if self._toolbar is None:
            self._toolbar = _Toolbar(self, self._color, self._width)
            self._toolbar.toolChanged.connect(self._on_tool)
            self._toolbar.colorChanged.connect(self._on_color)
            self._toolbar.widthChanged.connect(self._on_width)
            self._toolbar.actionTriggered.connect(self._perform)
        self._toolbar.show()
        self._place_toolbar()
        self.setCursor(Qt.ArrowCursor)

    def _place_toolbar(self):
        bar, sel = self._toolbar, self._sel
        if bar is None or sel is None:
            return
        bar.adjustSize()
        x = max(0, min(sel.right() + 1 - bar.width(), self.width() - bar.width()))
        y = sel.bottom() + 8
        if y + bar.height() > self.height():
            y = sel.top() - bar.height() - 8
            if y < 0:
                y = max(0, sel.bottom() - bar.height() - 8)
        bar.move(x, y)
        bar.raise_()

    def _on_tool(self, name):
        self._finish_text(True)
        self._tool = name
        self._place_toolbar()
        self._update_cursor(self._cursor_pos)

    def _on_color(self, color):
        self._color = color

    def _on_width(self, width):
        self._width = width

    def _reset_or_close(self):
        if self._sel is None and self._drag_rect is None:
            self.close()
            return
        self._finish_text(False)
        self._sel, self._drag_rect, self._mode, self._draft = None, None, None, None
        self._items.clear()
        self._tool = ''
        if self._toolbar is not None:
            self._toolbar.hide()
            self._toolbar.deleteLater()
            self._toolbar = None
        self.setCursor(Qt.CrossCursor)
        self._update_hover()
        self.update()

    # -- 输出
    def render_result(self):
        """裁剪选区并烧录标注，返回物理像素的 QPixmap。"""
        self._finish_text(True)
        dpr, sel = self._map.dpr, self._sel
        source = QRect(round(sel.x() * dpr), round(sel.y() * dpr), round(sel.width() * dpr), round(sel.height() * dpr))
        result = self._shot.copy(source)
        result.setDevicePixelRatio(1.0)
        if self._items:
            painter = QPainter(result)
            painter.translate(-source.x(), -source.y())
            painter.scale(dpr, dpr)
            painter.setClipRect(QRectF(source.x() / dpr, source.y() / dpr, source.width() / dpr, source.height() / dpr))
            mosaic = self._mosaic_pixmap() if any(i.kind == 'mosaic' for i in self._items) else None
            for item in self._items:
                _paint_item(painter, item, mosaic)
            painter.end()
        return result

    def _perform(self, action):
        if action == 'undo':
            self._undo()
            return
        if action == 'cancel':
            self.close()
            return
        if self._sel is None:
            return
        pixmap = self.render_result()
        if action == 'copy':
            QApplication.clipboard().setPixmap(pixmap)
            self._controller.notify(tr("截图已复制到剪贴板"))
            self.close()
        elif action == 'save':
            self.hide()
            path = save_pixmap_dialog(pixmap)
            if path is None:
                self.show()
                self.activateWindow()
                self.setFocus()
                return
            if path:
                self._controller.notify(tr("截图已保存：{}").format(path), 'info' if os.path.exists(path) else 'warning')
            self.close()
        elif action == 'pin':
            pixmap.setDevicePixelRatio(self._map.dpr)
            origin = self._map.geometry.topLeft()
            self._controller.pin(pixmap, origin + self._sel.topLeft())
            self.close()

    def closeEvent(self, event):
        self._finish_text(False)
        super().closeEvent(event)
        self.closed.emit()


def save_pixmap_dialog(pixmap, parent=None):
    """弹出保存对话框；用户取消返回 None，写入失败返回空字符串，成功返回路径。"""
    global _last_save_dir
    directory = _last_save_dir or QStandardPaths.writableLocation(QStandardPaths.PicturesLocation) or os.path.expanduser('~')
    default = os.path.join(directory, datetime.now().strftime('TabEx_%Y%m%d_%H%M%S.png'))
    path, _selected = QFileDialog.getSaveFileName(
        parent, tr("保存截图"), default, "PNG (*.png);;JPEG (*.jpg *.jpeg);;BMP (*.bmp)")
    if not path:
        return None
    if not os.path.splitext(path)[1]:
        path += '.png'
    _last_save_dir = os.path.dirname(path)
    return path if pixmap.save(path) else ''


# ---------------------------------------------------------------- 贴图

class PinWindow(QWidget):
    """悬浮在桌面上的截图：左键拖动，滚轮缩放，Ctrl+滚轮调透明度，双击/Esc 关闭，右键菜单。"""

    def __init__(self, pixmap, position):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self._base = pixmap
        self._drag_offset = None
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setAttribute(Qt.WA_QuitOnClose, False)
        self.setFocusPolicy(Qt.StrongFocus)
        dpr = pixmap.devicePixelRatio() or 1.0
        self._logical = (pixmap.width() / dpr, pixmap.height() / dpr)
        self.resize(max(1, round(self._logical[0])), max(1, round(self._logical[1])))
        self.move(position)

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
        painter.drawPixmap(self.rect(), self._base)
        painter.setPen(QPen(ACCENT, 1))
        painter.drawRect(self.rect().adjusted(0, 0, -1, -1))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_offset = event.globalPos() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if self._drag_offset is not None and event.buttons() & Qt.LeftButton:
            self.move(event.globalPos() - self._drag_offset)

    def mouseReleaseEvent(self, _event):
        self._drag_offset = None

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.close()

    def wheelEvent(self, event):
        steps = event.angleDelta().y() / 120.0
        if event.modifiers() & Qt.ControlModifier:
            self.setWindowOpacity(max(0.1, min(1.0, self.windowOpacity() + 0.1 * steps)))
            return
        zoom = max(0.1, min(8.0, (self.width() / self._logical[0]) * (1.1 ** steps)))
        centre = self.geometry().center()
        self.resize(max(8, round(self._logical[0] * zoom)), max(8, round(self._logical[1] * zoom)))
        self.move(centre - QPoint(self.width() // 2, self.height() // 2))

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        elif event.key() == Qt.Key_C and event.modifiers() & Qt.ControlModifier:
            self._copy()
        else:
            super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        menu.addAction(tr("复制"), self._copy)
        menu.addAction(tr("保存..."), lambda: save_pixmap_dialog(self._base))
        menu.addSeparator()
        menu.addAction(tr("关闭"), self.close)
        exec_menu(menu, event.globalPos())

    def _copy(self):
        clean = QPixmap(self._base)
        clean.setDevicePixelRatio(1.0)
        QApplication.clipboard().setPixmap(clean)

    def closeEvent(self, event):
        if self in _pins:
            _pins.remove(self)
        super().closeEvent(event)


def close_all_pins():
    for pin in list(_pins):
        pin.close()


# ---------------------------------------------------------------- 全局热键与入口

class _GlobalHotkey(QAbstractNativeEventFilter):
    """RegisterHotKey 全局热键；WM_HOTKEY 投递到注册线程（UI 线程）的消息队列。"""

    _next_id = 0x7E57

    def __init__(self, callback):
        super().__init__()
        self._callback = callback
        self._registered = False
        self._installed = False
        self._id = _GlobalHotkey._next_id
        _GlobalHotkey._next_id += 1

    def register(self, ctrl, shift, alt, vk):
        self.unregister()
        if os.name != 'nt' or not (ctrl or shift or alt):
            return False
        modifiers = 0x4000 | (0x2 if ctrl else 0) | (0x4 if shift else 0) | (0x1 if alt else 0)  # MOD_NOREPEAT
        if not ctypes.windll.user32.RegisterHotKey(None, self._id, modifiers, vk):
            debug_print(f"[Screenshot] RegisterHotKey failed err={ctypes.get_last_error() or ctypes.GetLastError()}")
            return False
        if not self._installed:
            QApplication.instance().installNativeEventFilter(self)
            self._installed = True
        self._registered = True
        return True

    def unregister(self):
        if self._registered:
            ctypes.windll.user32.UnregisterHotKey(None, self._id)
            self._registered = False

    def nativeEventFilter(self, event_type, message):
        if self._registered and bytes(event_type) == b'windows_generic_MSG':
            msg = wt.MSG.from_address(int(message))
            if msg.message == 0x0312 and msg.wParam == self._id:
                QTimer.singleShot(0, self._callback)
                return True, 0
        return False, 0


class ScreenshotController(QObject):
    """截图入口：抓取光标所在屏幕、管理叠加层与全局热键。同一时刻只允许一个截图会话。"""

    def __init__(self, notify=None, parent=None):
        super().__init__(parent)
        self._notify = notify
        self._overlay = None
        self._pending = False
        self._hotkey = _GlobalHotkey(self.start)

    def notify(self, message, level='info'):
        if self._notify is not None:
            self._notify(message, level)

    def apply_hotkey(self, parsed):
        """parsed 为 (ctrl, shift, alt, vk) 或 None；返回全局热键是否注册成功。"""
        if parsed is None:
            self._hotkey.unregister()
            return False
        ctrl, shift, alt, vk = parsed
        return self._hotkey.register(ctrl, shift, alt, vk)

    def is_active(self):
        return self._overlay is not None or self._pending

    def start(self, delay_ms=0):
        if self.is_active():
            return
        self._pending = True
        QTimer.singleShot(max(0, int(delay_ms)), self._start_now)

    def _start_now(self):
        try:
            screen = QGuiApplication.screenAt(QCursor.pos()) or QGuiApplication.primaryScreen()
            windows = _top_level_windows()
            shot = screen.grabWindow(0)
            if shot.isNull():
                self.notify(tr("截图失败：无法抓取屏幕"), 'warning')
                return
            overlay = ScreenshotOverlay(self, screen, shot, windows)
            overlay.closed.connect(self._on_closed)
            self._overlay = overlay
            overlay.begin()
        except Exception as exc:
            debug_print(f"[Screenshot] start failed: {exc}")
            self._overlay = None
            self.notify(tr("截图失败：{}").format(exc), 'warning')
        finally:
            self._pending = False

    def _on_closed(self):
        self._overlay = None

    def pin(self, pixmap, position):
        window = PinWindow(pixmap, position)
        _pins.append(window)
        window.show()
        window.raise_()

    def shutdown(self):
        self._hotkey.unregister()
        if self._overlay is not None:
            self._overlay.close()
        close_all_pins()

"""通用界面组件：提示通知、工具图标、手势层与分割条。"""

import os
import sys

from PyQt5.QtCore import pyqtSignal, QMimeData, QSize, Qt, QTimer
from PyQt5.QtGui import QDrag, QMouseEvent
from PyQt5.QtWidgets import (
    QApplication, QHBoxLayout, QLabel, QSizePolicy, QSplitter, QSplitterHandle as _QSplitterHandle,
    QToolButton, QVBoxLayout, QWidget,
)

from .paths import get_app_base_dir
from .i18n import tr
from .constants import MAX_ACTIVE_TOASTS
from .debuglog import debug_print
from .system import find_git_install_root
from .title_shortcuts import TitleShortcutBar


# 全局轻量提示气泡，用于替换阻塞式消息框
_active_toasts = []


_TOOL_ICON_FILES = {
    'go-previous': 'arrow-left', 'go-next': 'arrow-right', 'tab-new': 'square-plus',
    'edit-undo': 'undo-2', 'edit-find': 'search', 'view-group': 'group',
    'view-split-left-right': 'columns-2', 'user-bookmarks': 'bookmark',
    'preferences-system': 'settings', 'help-contents': 'bot',
    'view-list-details': 'list-filter', 'workspace-tools': 'panels-top-left',
    'file-tasks': 'clipboard-list', 'task-warning': 'triangle-alert',
    'window-close': 'x', 'document-open': 'folder-open', 'view-refresh': 'rotate-cw',
    'process-stop': 'square', 'accessories-calculator': 'calculator',
    'app-cmd': 'terminal', 'app-powershell': 'square-terminal', 'app-git-bash': 'git-fork',
    'app-tortoisegit-log': 'git-branch', 'app-tortoisegit-commit': 'git-commit-horizontal',
}
_TOOL_ASSET_CACHE = {}
_TOOL_NATIVE_CACHE = {}
_RED_PIN_ICON = None


def _pinned_tab_icon():
    global _RED_PIN_ICON
    if _RED_PIN_ICON is None:
        from PyQt5.QtGui import QFont, QIcon, QPainter, QPixmap
        icon = QIcon()
        for size in (16, 24, 32, 48, 64):
            pixmap = QPixmap(size, size)
            pixmap.fill(Qt.transparent)
            painter = QPainter(pixmap)
            font = QFont('Segoe UI Emoji')
            font.setPixelSize(max(12, round(size * 0.85)))
            painter.setFont(font)
            painter.drawText(pixmap.rect(), Qt.AlignCenter, '📌')
            painter.end()
            icon.addPixmap(pixmap)
        _RED_PIN_ICON = icon
    return _RED_PIN_ICON


def _tool_asset_icon(name):
    from PyQt5.QtGui import QIcon
    from PyQt5.QtSvg import QSvgRenderer
    root = getattr(sys, '_MEIPASS', get_app_base_dir())
    path = os.path.join(root, 'icons', name + '.svg')
    if path not in _TOOL_ASSET_CACHE:
        _TOOL_ASSET_CACHE[path] = QIcon(path) if QSvgRenderer(path).isValid() else QIcon()
    return _TOOL_ASSET_CACHE[path]


def _native_tool_executable(tool_name):
    if os.name != 'nt':
        return None
    system = os.environ.get('SystemRoot', r'C:\Windows')
    if tool_name == 'cmd':
        candidates = [os.environ.get('COMSPEC', ''), os.path.join(system, 'System32', 'cmd.exe')]
    elif tool_name == 'powershell':
        candidates = [os.path.join(system, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')]
    elif tool_name == 'git-bash':
        root = find_git_install_root()
        candidates = [os.path.join(root, 'git-bash.exe')] if root else []
    elif tool_name == 'tortoisegit':
        candidates = [os.path.join(base, 'TortoiseGit', 'bin', 'TortoiseGitProc.exe') for base in (
            os.environ.get('ProgramW6432', r'C:\Program Files'),
            os.environ.get('ProgramFiles', r'C:\Program Files'),
            os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'))]
    else:
        return None
    return next((path for path in candidates if path and not path.startswith('\\\\')
                 and os.path.isfile(path)), None)


def _native_tool_icon(tool_name):
    from PyQt5.QtGui import QIcon
    if tool_name not in _TOOL_NATIVE_CACHE:
        path = _native_tool_executable(tool_name)
        icon = TitleShortcutBar._extract_icon_fast(path) if path else None
        _TOOL_NATIVE_CACHE[tool_name] = icon if icon is not None and not icon.isNull() else QIcon()
    return _TOOL_NATIVE_CACHE[tool_name]


def _commit_badged_icon(native):
    from PyQt5.QtGui import QColor, QIcon, QPainter, QPixmap
    icon = QIcon()
    badge = _tool_asset_icon('check')
    for size in (16, 24, 32, 48, 64):
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        native.paint(painter, 0, 0, size, size)
        badge_size = max(8, size // 2)
        origin = size - badge_size
        painter.fillRect(origin, origin, badge_size, badge_size, QColor('#e0f4e5'))
        badge.paint(painter, origin, origin, badge_size, badge_size)
        painter.end()
        icon.addPixmap(pixmap)
    return icon


def _set_tool_icon(button, theme_name, fallback, size=18):
    from PyQt5.QtGui import QIcon
    native_tools = {'app-cmd': 'cmd', 'app-powershell': 'powershell', 'app-git-bash': 'git-bash',
                    'app-tortoisegit-log': 'tortoisegit', 'app-tortoisegit-commit': 'tortoisegit'}
    icon = _native_tool_icon(native_tools[theme_name]) if theme_name in native_tools else QIcon()
    if theme_name == 'app-tortoisegit-commit' and not icon.isNull():
        icon = _commit_badged_icon(icon)
    if icon.isNull() and theme_name in _TOOL_ICON_FILES:
        icon = _tool_asset_icon(_TOOL_ICON_FILES[theme_name])
    if icon.isNull():
        icon = QIcon.fromTheme(theme_name, button.style().standardIcon(fallback))
    button.setIcon(icon)
    button.setIconSize(QSize(size, size))
    button.setAccessibleName(button.toolTip() or button.text())


def _position_toasts():
    offsets = {}
    for toast in list(_active_toasts):
        anchor = toast.parentWidget()
        bounds = anchor.window().geometry() if anchor else QApplication.primaryScreen().availableGeometry()
        offset = offsets.get(anchor, 0)
        toast.move(bounds.right() - toast.width() - 20,
                   bounds.bottom() - toast.height() - 20 - offset)
        offsets[anchor] = offset + toast.height() + 10


class ToastMessage(QWidget):
    """右下角弹出的轻量提示，5s 自动消失"""

    def __init__(self, parent, title, message, level="info", duration=5000, action_text='', action=None):
        super().__init__(parent)
        self.duration = duration
        self.level = level
        self.remaining_seconds = duration // 1000  # 剩余秒数
        self.setWindowFlags(
            Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.NoDropShadowWindowHint
        )
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WA_DeleteOnClose, True)

        bg_map = {
            "info": "#1767b5",
            "warning": "#946000",
            "error": "#b42318",
            "critical": "#b42318",
            "success": "#18733b",
        }
        bg_color = bg_map.get(level, bg_map['info'])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        # 标题行（标题 + 倒计时）
        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(8)
        
        title_label = QLabel(title)
        title_label.setTextFormat(Qt.PlainText)
        title_label.setWordWrap(True)
        title_label.setStyleSheet("font-weight: bold; color: #ffffff; background: transparent;")
        title_layout.addWidget(title_label)
        
        title_layout.addStretch()
        
        self.countdown_label = QLabel(f"{self.remaining_seconds}s")
        self.countdown_label.setStyleSheet("color: #ffffff; background: transparent; font-size: 11px;")
        title_layout.addWidget(self.countdown_label)
        from PyQt5.QtWidgets import QStyle
        self.close_button = QToolButton(self)
        self.close_button.setToolTip(tr('关闭'))
        self.close_button.setFixedSize(24, 24)
        _set_tool_icon(self.close_button, 'window-close', QStyle.SP_TitleBarCloseButton, 14)
        self.close_button.clicked.connect(self.close)
        title_layout.addWidget(self.close_button)
        
        layout.addLayout(title_layout)
        
        msg_label = QLabel(message)
        msg_label.setTextFormat(Qt.PlainText)
        msg_label.setWordWrap(True)
        msg_label.setStyleSheet("color: #ffffff; background: transparent;")
        msg_label.setMinimumWidth(0)
        msg_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)

        layout.addWidget(msg_label)
        self.action_button = None
        if action is not None and action_text:
            self.action_button = QToolButton(self)
            self.action_button.setText(action_text)
            self.action_button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
            self.action_button.setToolTip(action_text)
            _set_tool_icon(self.action_button, 'document-open', QStyle.SP_DirOpenIcon, 16)

            def activate():
                self.close()
                action()

            self.action_button.clicked.connect(activate)
            layout.addWidget(self.action_button, 0, Qt.AlignRight)

        self.setObjectName('tabexToast')
        self.setStyleSheet(
            f"QWidget#tabexToast {{ background: {bg_color}; border: 1px solid rgba(0, 0, 0, 40); border-radius: 6px; }}"
            "QToolButton { color: #202020; background: #ffffff; padding: 3px; border: 1px solid #d3deea; border-radius: 3px; }"
            "QToolButton:hover { background: #edf3f9; }"
            "QToolButton:pressed { background: #d9e4f0; }"
            "QToolButton:focus { border-color: #202020; }"
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
        anchor = self.parent() if isinstance(self.parent(), QWidget) else None
        
        # 使用软件窗口的几何信息，而不是屏幕的几何信息
        if anchor and anchor.window():
            window_geo = anchor.window().geometry()
        else:
            # 如果没有父窗口，使用屏幕几何作为后备
            window_geo = QApplication.primaryScreen().availableGeometry()
        
        margin = 20
        self.setFixedWidth(min(420, max(220, window_geo.width() - margin * 2)))
        self.adjustSize()
        _position_toasts()

    def enterEvent(self, event):
        if self._timer.isActive():
            self._remaining_ms = max(1, self._timer.remainingTime())
        self._timer.stop()
        self._countdown_timer.stop()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self.isVisible():
            self._timer.start(getattr(self, '_remaining_ms', self.duration))
            self._countdown_timer.start(1000)
        super().leaveEvent(event)

    def closeEvent(self, event):
        if self in _active_toasts:
            _active_toasts.remove(self)
        # 停止倒计时定时器
        if hasattr(self, '_countdown_timer'):
            self._countdown_timer.stop()
        self._timer.stop()
        super().closeEvent(event)
        _position_toasts()


def show_toast(parent, title, message, level="info", duration=5000, action_text='', action=None):
    """在右下角显示非阻塞提示"""
    anchor = parent.window() if isinstance(parent, QWidget) else None
    while len(_active_toasts) >= MAX_ACTIVE_TOASTS:
        old_toast = _active_toasts.pop(0)
        try:
            old_toast.close()
        except Exception:
            pass
    toast = ToastMessage(anchor, title, message, level=level, duration=duration,
                         action_text=action_text, action=action)
    _active_toasts.append(toast)
    toast.show()
    return toast


class GestureOverlay(QWidget):
    """鼠标手势可视化覆盖层（类似 Mouse Gestures）。

    - 全屏透明、鼠标穿透、始终置顶。
    - 拖拽时绘制跟随光标的轨迹线。
    - 实时显示当前识别到的方向箭头 + 动作名称浮窗。
    坐标统一使用全局屏幕坐标（与 WH_MOUSE_LL 钩子一致）。
    """

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(
            Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
            | Qt.WindowTransparentForInput | Qt.NoDropShadowWindowHint
        )
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self._points = []      # 全局坐标点列表 [QPoint, ...]
        self._arrow = ""       # 当前方向箭头
        self._label = ""       # 当前方向动作名

    def start(self):
        """开始一次手势：铺满虚拟桌面并清空轨迹。"""
        try:
            vg = QApplication.primaryScreen().virtualGeometry()
            self.setGeometry(vg)
        except Exception:
            self.setGeometry(QApplication.primaryScreen().geometry())
        self._points = []
        self._arrow = ""
        self._label = ""
        self.show()
        self.raise_()

    def add_point(self, gx, gy):
        from PyQt5.QtCore import QPoint
        self._points.append(QPoint(int(gx), int(gy)))
        self.update()

    def set_hint(self, arrow, label):
        if arrow != self._arrow or label != self._label:
            self._arrow = arrow
            self._label = label
            self.update()

    def finish(self):
        self.hide()
        self._points = []
        self._arrow = ""
        self._label = ""

    def paintEvent(self, event):
        if not self._points:
            return
        from PyQt5.QtGui import QPainter, QPen, QColor, QFont, QFontMetrics
        from PyQt5.QtCore import QRect
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        try:
            scale = max(1.0, self.logicalDpiX() / 96.0)
        except Exception:
            scale = 1.0
        origin = self.geometry().topLeft()

        # ── 轨迹线 ───────────────────────────────────────────────
        pen = QPen(QColor(0, 200, 120, 230))  # 半透明绿色
        pen.setWidth(max(3, int(4 * scale)))
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        prev = None
        for p in self._points:
            lp = p - origin
            if prev is not None:
                painter.drawLine(prev, lp)
            prev = lp

        # ── 实时方向提示浮窗 ─────────────────────────────────────
        if self._arrow:
            text = (f"{self._arrow}  {self._label}").strip()
            font = QFont()
            font.setPointSizeF(max(12.0, 15.0 * scale))
            font.setBold(True)
            painter.setFont(font)
            fm = QFontMetrics(font)
            tw = fm.boundingRect(text).width()
            th = fm.height()
            pad = int(12 * scale)
            last = self._points[-1] - origin
            bx = last.x() + int(24 * scale)
            by = last.y() + int(24 * scale)
            rect = QRect(bx, by, tw + pad * 2, th + pad * 2)
            if rect.right() > self.width():
                rect.moveRight(self.width() - int(8 * scale))
            if rect.bottom() > self.height():
                rect.moveBottom(self.height() - int(8 * scale))
            if rect.left() < 0:
                rect.moveLeft(int(8 * scale))
            if rect.top() < 0:
                rect.moveTop(int(8 * scale))
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(0, 0, 0, 180))
            painter.drawRoundedRect(rect, int(8 * scale), int(8 * scale))
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(rect, Qt.AlignCenter, text)


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


class _ResizableSplitterHandle(_QSplitterHandle):
    """分割条手柄：作为原生窗口并强制显示左右调整光标，并在中央绘制可见抓取条。

    内嵌的 IExplorerBrowser 是原生窗口（HWND），夹在两个原生面板之间的普通（alien）
    手柄常常收不到鼠标悬停、光标不会变成调整光标。将手柄本身设为原生窗口可与相邻的
    Shell 原生窗口正确竞争命中测试，确保能拖动并显示左右调整光标。
    """

    def __init__(self, orientation, parent):
        super().__init__(orientation, parent)
        # 原生窗口：解决夹在原生 Shell 窗口之间时光标/命中测试失效的问题
        self.setAttribute(Qt.WA_NativeWindow, True)
        self.setCursor(Qt.SplitHCursor if orientation == Qt.Horizontal else Qt.SplitVCursor)
        self._hovered = False

    def enterEvent(self, event):
        self._hovered = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hovered = False
        self.update()
        super().leaveEvent(event)

    def paintEvent(self, event):
        from PyQt5.QtGui import QPainter, QColor
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#aab2c0") if self._hovered else QColor("#d2d7e0"))
        # 中央竖向抓取条（三段短竖线），提升可识别度
        cx = self.width() // 2
        cy = self.height() // 2
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#6b7686"))
        if self.orientation() == Qt.Horizontal:
            for dy in (-12, 0, 12):
                p.drawRect(cx - 1, cy + dy - 6, 2, 12)
        else:
            for dx in (-12, 0, 12):
                p.drawRect(cx + dx - 6, cy - 1, 12, 2)
        p.end()


class ResizableSplitter(QSplitter):
    """使用更宽、带原生窗口、明确光标与可见抓取条的手柄，便于在内嵌原生 Shell 窗口旁拖动调整宽度。"""

    def createHandle(self):
        return _ResizableSplitterHandle(self.orientation(), self)


def _build_te_icon():
    """生成多分辨率 TE 图标。仅依赖 QApplication 已创建，可在主窗口构建前调用。"""
    from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QIcon
    te_icon = QIcon()
    # 按 256px 基准等比例生成各尺寸图标
    for size in [256, 128, 96, 64, 48, 32, 24, 18, 16]:
        pix = QPixmap(size, size)
        pix.fill(Qt.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.Antialiasing)
        blue = QColor("#2196F3")
        white = QColor("white")
        # 外层蓝色圆角背景
        outer_radius = max(2, size * 40 // 256)
        p.setBrush(blue)
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(0, 0, size, size, outer_radius, outer_radius)
        # 内层白色圆角容器（形成蓝色边框效果）
        margin = max(2, size * 28 // 256)
        inner_radius = max(2, size * 24 // 256)
        p.setBrush(white)
        p.drawRoundedRect(margin, margin, size - 2*margin, size - 2*margin, inner_radius, inner_radius)
        # 中央蓝色 TE 文字
        p.setPen(blue)
        f = QFont()
        f.setBold(True)
        f.setPointSize(max(5, size * 130 // 256))
        f.setStretch(70)  # 压窄字体，使其看起来更高
        p.setFont(f)
        p.drawText(pix.rect(), Qt.AlignCenter, "TE")
        p.end()
        te_icon.addPixmap(pix, QIcon.Normal)
    return te_icon

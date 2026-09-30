"""标题栏快捷方式区域。"""

import ctypes.wintypes
import os

from PyQt5.QtCore import pyqtSignal, QEvent, QFileInfo, QMimeData, QSize, Qt, QTimer
from PyQt5.QtGui import QDrag, QIcon
from PyQt5.QtWidgets import QApplication, QFileIconProvider, QHBoxLayout, QLabel, QMenu, QToolButton, QWidget

from .i18n import tr
from .system import is_supported_title_shortcut_path


class TitleShortcutBar(QWidget):
    """标题栏快捷方式区域：支持拖拽常用启动文件并点击运行。"""

    shortcutDropped = pyqtSignal(str)
    shortcutClicked = pyqtSignal(str)
    shortcutsChanged = pyqtSignal(list)

    def __init__(self, parent=None, icon_size=18, button_size=28):
        super().__init__(parent)
        self._icon_size = icon_size
        self._button_size = button_size
        self._paths = []
        self._icon_cache = {}  # path → QIcon (avoids re-calling SHGetFileInfo)
        self._icon_provider = QFileIconProvider()
        self._drag_start_pos = None
        self._drag_source_index = -1
        self.setAcceptDrops(True)

        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(8)

        self._hint_label = QLabel(tr("拖入应用或快捷方式"))
        self._hint_label.setStyleSheet("QLabel { color: #666; font-size: 9pt; padding: 0 4px; }")
        self._hint_label.setAcceptDrops(True)
        self._hint_label.installEventFilter(self)
        self._layout.addWidget(self._hint_label)

    def set_shortcuts(self, paths):
        cleaned = []
        for path in paths or []:
            if isinstance(path, str) and path and path not in cleaned:
                cleaned.append(path)
        self._paths = cleaned[:20]
        self._rebuild_buttons()

    def get_shortcuts(self):
        return list(self._paths)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-tabex-shortcut-index"):
            event.acceptProposedAction()
            return
        if self._collect_shortcuts_from_mime(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat("application/x-tabex-shortcut-index"):
            event.acceptProposedAction()
            return
        if self._collect_shortcuts_from_mime(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasFormat("application/x-tabex-shortcut-index"):
            try:
                source_index = int(bytes(event.mimeData().data("application/x-tabex-shortcut-index")).decode("utf-8"))
            except Exception:
                event.ignore()
                return
            target_index = self._get_insert_index(event.pos().x())
            self._reorder_shortcut(source_index, target_index)
            event.acceptProposedAction()
            return

        paths = self._collect_shortcuts_from_mime(event.mimeData())
        if not paths:
            event.ignore()
            return
        for path in paths:
            self.shortcutDropped.emit(path)
        event.acceptProposedAction()

    def eventFilter(self, obj, event):
        if isinstance(obj, QToolButton) and obj.parent() is self:
            if event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
                path = obj.property("shortcut_path")
                if path in self._paths:
                    self._drag_start_pos = event.globalPos()
                    self._drag_source_index = self._paths.index(path)
                return False
            if event.type() == QEvent.MouseMove and (event.buttons() & Qt.LeftButton):
                if self._drag_start_pos is not None and self._drag_source_index >= 0:
                    if (event.globalPos() - self._drag_start_pos).manhattanLength() >= QApplication.startDragDistance():
                        source_index = self._drag_source_index
                        self._drag_start_pos = None
                        self._drag_source_index = -1
                        self._start_drag_from_index(source_index)
                        return True
                return False
            if event.type() == QEvent.MouseButtonRelease:
                self._drag_start_pos = None
                self._drag_source_index = -1
                return False

        if event.type() in (QEvent.DragEnter, QEvent.DragMove):
            mime_data = event.mimeData() if hasattr(event, 'mimeData') else None
            if mime_data and (mime_data.hasFormat("application/x-tabex-shortcut-index") or self._collect_shortcuts_from_mime(mime_data)):
                event.acceptProposedAction()
                return True
        elif event.type() == QEvent.Drop:
            mime_data = event.mimeData() if hasattr(event, 'mimeData') else None
            if not mime_data:
                return super().eventFilter(obj, event)

            if mime_data.hasFormat("application/x-tabex-shortcut-index"):
                try:
                    source_index = int(bytes(mime_data.data("application/x-tabex-shortcut-index")).decode("utf-8"))
                except Exception:
                    event.ignore()
                    return True
                local_pos = obj.mapTo(self, event.pos()) if isinstance(obj, QWidget) else event.pos()
                target_index = self._get_insert_index(local_pos.x())
                self._reorder_shortcut(source_index, target_index)
                event.acceptProposedAction()
                return True

            paths = self._collect_shortcuts_from_mime(mime_data)
            if paths:
                for path in paths:
                    self.shortcutDropped.emit(path)
                event.acceptProposedAction()
                return True

        return super().eventFilter(obj, event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_start_pos = event.pos()
            self._drag_source_index = self._button_index_at(event.pos())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.LeftButton):
            super().mouseMoveEvent(event)
            return
        if self._drag_start_pos is None or self._drag_source_index < 0:
            super().mouseMoveEvent(event)
            return
        if (event.pos() - self._drag_start_pos).manhattanLength() < QApplication.startDragDistance():
            super().mouseMoveEvent(event)
            return

        source_index = self._drag_source_index
        self._drag_start_pos = None
        self._drag_source_index = -1
        self._start_drag_from_index(source_index)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_start_pos = None
        self._drag_source_index = -1
        super().mouseReleaseEvent(event)

    def _collect_shortcuts_from_mime(self, mime_data):
        if not mime_data or not mime_data.hasUrls():
            return []
        paths = []
        for url in mime_data.urls():
            if not url.isLocalFile():
                continue
            path = url.toLocalFile()
            if not path:
                continue
            if is_supported_title_shortcut_path(path):
                paths.append(path)
        return paths

    def _resolve_shortcut_icon(self, path):
        if not isinstance(path, str) or not path:
            return QIcon()

        lower_path = path.lower()

        # For .exe files, use ExtractIconEx (fast, no shell timeout)
        if lower_path.endswith('.exe') and os.name == 'nt':
            icon = self._extract_icon_fast(path)
            if icon and not icon.isNull():
                return icon
            return self._icon_provider.icon(QFileInfo(path))

        if lower_path.endswith('.lnk') and os.name == 'nt':
            try:
                from win32com.client import Dispatch

                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(path)

                icon_location = getattr(shortcut, 'IconLocation', '') or ''
                if icon_location:
                    icon_path = icon_location.split(',', 1)[0].strip().strip('"')
                    icon_path = os.path.expandvars(icon_path)
                    if os.path.exists(icon_path):
                        icon = self._extract_icon_fast(icon_path) if icon_path.lower().endswith('.exe') else None
                        if not icon or icon.isNull():
                            icon = self._icon_provider.icon(QFileInfo(icon_path))
                        if not icon.isNull():
                            return icon

                target_path = getattr(shortcut, 'Targetpath', '') or getattr(shortcut, 'TargetPath', '') or ''
                target_path = os.path.expandvars(str(target_path).strip().strip('"'))
                if target_path and os.path.exists(target_path):
                    icon = self._extract_icon_fast(target_path) if target_path.lower().endswith('.exe') else None
                    if not icon or icon.isNull():
                        icon = self._icon_provider.icon(QFileInfo(target_path))
                    if not icon.isNull():
                        return icon
            except Exception:
                pass

        return self._icon_provider.icon(QFileInfo(path))

    @staticmethod
    def _extract_icon_fast(exe_path):
        """Extract icon from .exe using ExtractIconExW (fast, no shell timeout)."""
        try:
            import ctypes
            import ctypes.wintypes
            from PyQt5.QtGui import QPixmap
            from PyQt5.QtWinExtras import QtWin

            _ExtractIconExW = ctypes.windll.shell32.ExtractIconExW
            _ExtractIconExW.argtypes = [ctypes.c_wchar_p, ctypes.c_int,
                                        ctypes.POINTER(ctypes.wintypes.HICON),
                                        ctypes.POINTER(ctypes.wintypes.HICON),
                                        ctypes.c_uint]
            _ExtractIconExW.restype = ctypes.c_uint
            _DestroyIcon = ctypes.windll.user32.DestroyIcon

            hicon_large = ctypes.wintypes.HICON()
            hicon_small = ctypes.wintypes.HICON()
            count = _ExtractIconExW(exe_path, 0,
                                    ctypes.byref(hicon_large),
                                    ctypes.byref(hicon_small), 1)
            if count and hicon_large.value:
                try:
                    pixmap = QtWin.fromHICON(hicon_large.value)
                    if not pixmap.isNull():
                        return QIcon(pixmap)
                finally:
                    _DestroyIcon(hicon_large)
                    if hicon_small.value:
                        _DestroyIcon(hicon_small)
            elif hicon_small.value:
                _DestroyIcon(hicon_small)
        except Exception:
            pass
        return None

    def _rebuild_buttons(self):
        while self._layout.count() > 0:
            item = self._layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        # 提示文字固定在最左侧
            self._hint_label = QLabel(tr("+ 拖入应用或快捷方式"))
        self._hint_label.setStyleSheet("QLabel { color: #888; font-size: 9pt; padding: 0 4px; }")
        self._hint_label.setAcceptDrops(True)
        self._hint_label.installEventFilter(self)
        self._layout.addWidget(self._hint_label)

        if not self._paths:
            return

        self._layout.addSpacing(6)

        for path in self._paths:
            btn = QToolButton(self)
            btn.setAutoRaise(True)
            btn.setFixedSize(self._button_size, self._button_size)
            icon_px = max(12, min(self._icon_size + 1, self._button_size - 4))
            btn.setIconSize(QSize(icon_px, icon_px))
            btn.setToolTip(f"{os.path.basename(path)}\n{path}\n左键启动应用/快捷方式，右键移除，拖拽可排序")
            # Use cached icon or defer loading to avoid SHGetFileInfo blocking startup
            cached_icon = self._icon_cache.get(path)
            if cached_icon is not None:
                if not cached_icon.isNull():
                    btn.setIcon(cached_icon)
                else:
                    btn.setText("↗")
            else:
                btn.setText("…")
                # Defer icon loading well after startup to avoid SHGetFileInfo blocking
                QTimer.singleShot(1500, lambda b=btn, p=path: self._load_icon_deferred(b, p))
            btn.setStyleSheet(
                "QToolButton { background: transparent; border: none; border-radius: 4px; padding: 0px; margin: 0px; }"
                "QToolButton:hover { background: #e0e0e0; }"
                "QToolButton:pressed { background: #d0d0d0; }"
            )
            btn.setAcceptDrops(True)
            btn.installEventFilter(self)
            btn.setProperty("shortcut_path", path)
            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(lambda pos, b=btn: self._show_context_menu_for_button(b, pos))
            btn.clicked.connect(lambda _checked=False, p=path: self.shortcutClicked.emit(p))
            self._layout.addWidget(btn)

    def _load_icon_deferred(self, btn, path):
        """Load shortcut icon asynchronously and update button."""
        try:
            if not btn or not btn.isVisible():
                return
            icon = self._resolve_shortcut_icon(path)
            self._icon_cache[path] = icon
            if not icon.isNull():
                btn.setIcon(icon)
                btn.setText("")
            else:
                btn.setText("↗")
        except RuntimeError:
            pass  # Button already deleted

    def _show_context_menu_for_button(self, btn, pos):
        path = btn.property("shortcut_path")
        if not path:
            return
        menu = QMenu(self)
        menu.setStyleSheet(
            """
            QMenu {
                background-color: #ffffff;
                border: 1px solid #c7c7c7;
                border-radius: 6px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 24px 6px 12px;
                background: transparent;
                border-radius: 4px;
                color: #303030;
                margin: 2px 4px;
            }
            QMenu::item:selected {
                background: #3f5f7a;
                color: #f6f8fa;
            }
            QMenu::item:pressed {
                background: #324c62;
                color: #ffffff;
            }
            """
        )
        remove_action = menu.addAction(tr("移除该快捷方式"))
        action = menu.exec_(btn.mapToGlobal(pos))
        if action == remove_action:
            self._remove_shortcut(path)

    def _remove_shortcut(self, path):
        if path not in self._paths:
            return
        self._paths = [p for p in self._paths if p != path]
        self._rebuild_buttons()
        self.shortcutsChanged.emit(list(self._paths))

    def _button_index_at(self, pos):
        widget = self.childAt(pos)
        while widget and widget is not self:
            if isinstance(widget, QToolButton):
                path = widget.property("shortcut_path")
                if path in self._paths:
                    return self._paths.index(path)
            widget = widget.parentWidget()
        return -1

    def _get_insert_index(self, x_pos):
        button_centers = []
        for i in range(self._layout.count()):
            item = self._layout.itemAt(i)
            widget = item.widget()
            if isinstance(widget, QToolButton):
                g = widget.geometry()
                button_centers.append((g.left() + g.right()) // 2)
        if not button_centers:
            return 0
        for idx, center in enumerate(button_centers):
            if x_pos < center:
                return idx
        return len(button_centers)

    def _reorder_shortcut(self, source_index, target_index):
        if source_index < 0 or source_index >= len(self._paths):
            return
        if target_index > source_index:
            target_index -= 1
        if target_index < 0:
            target_index = 0
        if target_index >= len(self._paths):
            target_index = len(self._paths) - 1
        if target_index == source_index:
            return
        moving = self._paths.pop(source_index)
        self._paths.insert(target_index, moving)
        self._rebuild_buttons()
        self.shortcutsChanged.emit(list(self._paths))

    def _start_drag_from_index(self, source_index):
        if source_index < 0 or source_index >= len(self._paths):
            return
        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setData("application/x-tabex-shortcut-index", str(source_index).encode("utf-8"))
        drag.setMimeData(mime_data)
        drag.exec_(Qt.MoveAction)

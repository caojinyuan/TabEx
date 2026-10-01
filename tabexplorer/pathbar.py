"""面包屑地址栏。"""

import os

from PyQt5.QtCore import pyqtSignal, QDir, QEvent, QMimeData, Qt, QUrl
from PyQt5.QtGui import QDrag
from PyQt5.QtWidgets import QApplication, QLineEdit, QMenu, QWidget

from . import debuglog as _debuglog
from . import theme as _theme
from .i18n import tr
from .debuglog import debug_print


class SimplePathBar(QWidget):
    """Windows Explorer 风格面包屑地址栏（单控件自绘实现）。
    - 普通模式：paintEvent 一次性绘制可点击的路径分段 (C: › project › EOL)，
      通过命中测试判断点击了哪一段；不再为每一段创建/销毁子控件。
    - 编辑模式：覆盖显示常驻 QLineEdit，按 Enter 导航，按 Escape 取消。
    - 点击路径段之间的空白区域或调用 enter_edit_mode() 进入编辑模式。
    接口与旧实现兼容：set_path / pathChanged /
    enter_edit_mode / exit_edit_mode / get_path_for_copy。

    重要：旧实现每次导航都销毁并重建 N 个 QPushButton/QToolButton，叠加多轮
    强制重排与自愈循环，导致最小化恢复后偶发“地址栏不刷新/卡顿”（数据正确、
    像素滞后）。单控件自绘从根本上消除该问题——一个 paintEvent 一定会被 Qt
    正确绘制，且没有子控件析构/创建开销。
    """
    pathChanged = pyqtSignal(str)
    activated = pyqtSignal()

    _FONT   = "font-family: 'Segoe UI', 'Microsoft YaHei UI', sans-serif; font-size: 11pt;"
    _S_BAR  = ("SimplePathBar { background: #ffffff; border: none; }")
    _S_EDIT = ("QLineEdit { background: white; border: 1px solid #ccc; border-radius: 3px;"
               " font-family: 'Segoe UI', 'Microsoft YaHei UI', sans-serif; font-size: 10pt;"
               " color: #202020; padding: 2px 6px; selection-background-color: #0078d4; }")

    # 绘制配色（与旧样式表一致）
    _COL_SEG   = '#003d7a'   # 路径段文字
    _COL_SEP   = '#888888'   # 分隔符 ›
    _COL_HOVER = '#cce5ff'   # 悬停背景
    _CHEVRON   = '\u203a'    # ›
    _PAD       = 6           # 段内文字左右内边距
    _SEP_W     = 16          # 分隔符宽度

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_path = ''
        self._segments = []            # [(显示名, 完整路径), …] 全部层级
        self._display_regions = []     # 绘制时记录的命中区: dict(kind,x0,x1,payload,label,is_current)
        self._hover_idx = -1
        self._in_edit = False
        self._pane_active = False
        self._split_indicator = False
        self._completer = None         # 编辑模式路径自动补全（首次进入编辑时惰性创建）
        self._press_pos = None         # 左键按下位置（用于拖拽判定）
        self._press_idx = -1           # 按下时命中的区索引
        self._dragging = False
        self.setFixedHeight(30)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMouseTracking(True)    # hover 高亮需要
        self.setAttribute(Qt.WA_StyledBackground, True)
        _theme.bind_style(self, self._S_BAR)

        from PyQt5.QtGui import QFont
        self._font = QFont("Segoe UI", 11)
        try:
            self._font.setWeight(QFont.Medium)  # 500
        except Exception:
            pass
        # 让 self.fontMetrics()（_choose_display_parts / _elide_crumb_text 使用）
        # 与 paintEvent 绘制字体一致，避免宽度测算与实际绘制不符导致的折叠/裁剪偏差。
        self.setFont(self._font)
        # 编辑框：常驻子控件，编辑模式下覆盖显示，不参与常态绘制
        self._edit = QLineEdit(self)
        _theme.bind_style(self._edit, self._S_EDIT)
        self._edit.setFrame(False)
        self._edit.returnPressed.connect(self._commit_edit)
        self._edit.installEventFilter(self)
        self._edit.hide()

    # ── 公共 API ─────────────────────────────────────────────────────────────

    def set_path(self, path):
        new_path = path or ''
        path_changed = (new_path != self._current_path)
        if _debuglog._DEBUG_MODE:
            debug_print(f"[SimplePathBar] set_path: new='{new_path}', old='{self._current_path}', changed={path_changed}, in_edit={self._in_edit}")
        self._current_path = new_path
        if self._in_edit:
            if path_changed:
                # 路径已变化（用户通过 Explorer 导航了）→ 退出编辑模式并显示新路径
                self.exit_edit_mode()
            # 路径未变化时保持编辑模式不中断（用户可能在复制路径）
            return
        # 单控件自绘：只更新分段数据并请求重绘，绝不创建/销毁子控件。
        if path_changed:
            # 空路径时保留现有分段，避免瞬时空白。
            if new_path:
                self._segments = self._split_path(new_path)
            self._hover_idx = -1
            self.update()
        else:
            # 路径未变化（来自轮询/保活的重复回写）：轻量异步刷新即可，无需强制同步。
            self.update()

    def force_refresh(self):
        """同步重绘：供窗口恢复/手动兜底调用，强制像素立即跟上当前路径。
        单控件 repaint() 一定被 Qt 立即执行，规避恢复后异步 update() 被合成器丢弃。"""
        if self._in_edit:
            return
        if self._current_path:
            self._segments = self._split_path(self._current_path)
        self.repaint()

    def set_pane_active(self, active, split):
        state = (bool(active), bool(split))
        if state == (self._pane_active, self._split_indicator):
            return
        self._pane_active, self._split_indicator = state
        border = '#2f6fdb' if active else '#a0a5ad'
        _theme.bind_style(self._edit, self._S_EDIT + (f'QLineEdit {{ border: 2px solid {border}; }}' if split else ''))
        self.update()

    def enter_edit_mode(self):
        if self._in_edit:
            return
        debug_print(f"[SimplePathBar] enter_edit_mode: path='{self._current_path}'")
        self._ensure_completer()
        self._in_edit = True
        self._edit.setGeometry(self.rect())
        self._edit.setText(self._current_path)
        self._edit.show()
        self._edit.raise_()
        self._edit.setFocus()
        self._edit.selectAll()
        # 安装应用级事件过滤器，监听点击外部区域时退出编辑模式
        QApplication.instance().installEventFilter(self)
        self.update()

    def exit_edit_mode(self):
        if not self._in_edit:
            return  # 已经不在编辑模式
        self._in_edit = False
        self._edit.hide()
        # 移除应用级事件过滤器
        try:
            QApplication.instance().removeEventFilter(self)
        except Exception:
            pass
        self.update()

    def changeEvent(self, event):
        """窗口激活时请求重绘（单控件自绘，无子控件丢失问题）。"""
        if event.type() == QEvent.WindowActivate and not self._in_edit:
            self.update()
        super().changeEvent(event)

    # ── 自绘 + 命中测试 ────────────────────────────────────────────────────────

    def _region_at(self, x):
        """返回横坐标 x 命中的显示区索引；未命中返回 -1。"""
        for i, r in enumerate(self._display_regions):
            if r['x0'] <= x < r['x1']:
                return i
        return -1

    def paintEvent(self, event):
        if self._in_edit:
            return
        from PyQt5.QtGui import QPainter, QColor, QFont
        from PyQt5.QtCore import QRectF
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(_theme.bg('#ffffff')))
        if self._split_indicator:
            color = '#2f6fdb' if self._pane_active else '#b7bcc4'
            painter.fillRect(0, self.height() - 2, self.width(), 2, QColor(_theme.line(color)))
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.TextAntialiasing, True)
        painter.setFont(self._font)
        fm = painter.fontMetrics()
        h = self.height()
        y = int((h + fm.ascent() - fm.descent()) / 2)
        regions = []
        parts = self._segments
        if not parts:
            self._display_regions = regions
            painter.end()
            return
        display_parts = self._choose_display_parts(parts)
        left_collapsed = len(display_parts) < len(parts)
        col_seg = QColor(_theme.fg(self._COL_SEG))
        col_sep = QColor(_theme.fg(self._COL_SEP))
        col_hover = QColor(_theme.bg(self._COL_HOVER))
        pad = self._PAD
        sep_w = self._SEP_W
        chevron = self._CHEVRON
        x = 2

        def hover_bg(x0, x1):
            painter.save()
            painter.setPen(Qt.NoPen)
            painter.setBrush(col_hover)
            painter.drawRoundedRect(QRectF(float(x0), 4.0, float(x1 - x0), float(h - 8)), 3.0, 3.0)
            painter.restore()

        idx = 0
        if left_collapsed:
            # 折叠省略号（点击进入编辑模式，与旧行为一致）
            ell = '...'
            w = fm.horizontalAdvance(ell) + pad * 2
            if self._hover_idx == idx:
                hover_bg(x, x + w)
            painter.setPen(col_sep)
            painter.drawText(x + pad, y, ell)
            regions.append({'kind': 'ellipsis', 'x0': x, 'x1': x + w,
                            'payload': None, 'label': self._current_path, 'is_current': False})
            x += w
            idx += 1
            # 折叠后的分隔符：列出首个显示段的同级文件夹
            first_parent = os.path.dirname(display_parts[0][1]) if display_parts else None
            if self._hover_idx == idx and first_parent:
                hover_bg(x, x + sep_w)
            painter.setPen(col_sep)
            painter.drawText(x + (sep_w - fm.horizontalAdvance(chevron)) // 2, y, chevron)
            regions.append({'kind': 'separator', 'x0': x, 'x1': x + sep_w,
                            'payload': first_parent, 'label': None, 'is_current': False})
            x += sep_w
            idx += 1

        for i, (label, full) in enumerate(display_parts):
            if i > 0:
                # 分隔符下拉列出左侧段的子文件夹（即当前段的同级）
                parent = display_parts[i - 1][1]
                if self._hover_idx == idx and parent:
                    hover_bg(x, x + sep_w)
                painter.setPen(col_sep)
                painter.drawText(x + (sep_w - fm.horizontalAdvance(chevron)) // 2, y, chevron)
                regions.append({'kind': 'separator', 'x0': x, 'x1': x + sep_w,
                                'payload': parent, 'label': None, 'is_current': False})
                x += sep_w
                idx += 1
            is_current = (full == self._current_path)
            max_text_w = max(0, (self.width() - 2) - x - pad * 2)
            shown = self._elide_crumb_text(label, is_current=is_current, max_text_width=max_text_w)
            w = fm.horizontalAdvance(shown) + pad * 2
            hovered = (self._hover_idx == idx)
            if hovered:
                hover_bg(x, x + w)
            f = QFont(self._font)
            f.setUnderline(hovered)
            painter.setFont(f)
            painter.setPen(col_seg)
            painter.drawText(x + pad, y, shown)
            painter.setFont(self._font)
            regions.append({'kind': 'segment', 'x0': x, 'x1': x + w,
                            'payload': full, 'label': label, 'is_current': is_current})
            x += w
            idx += 1

        self._display_regions = regions
        painter.end()

    def mouseMoveEvent(self, event):
        if self._in_edit:
            return super().mouseMoveEvent(event)
        # 拖拽判定：在某个路径段上按下并拖动 → 拖到标签栏打开新 tab
        if self._press_pos is not None and (event.buttons() & Qt.LeftButton):
            if (0 <= self._press_idx < len(self._display_regions) and
                    (event.pos() - self._press_pos).manhattanLength() >= QApplication.startDragDistance()):
                r = self._display_regions[self._press_idx]
                if r['kind'] == 'segment' and r['payload']:
                    self._dragging = True
                    drag = QDrag(self)
                    mime = QMimeData()
                    mime.setUrls([QUrl.fromLocalFile(r['payload'])])
                    mime.setText(r['payload'])
                    drag.setMimeData(mime)
                    self._press_pos = None
                    self._press_idx = -1
                    drag.exec_(Qt.CopyAction | Qt.MoveAction)
                    return
        idx = self._region_at(int(event.pos().x()))
        if idx != self._hover_idx:
            self._hover_idx = idx
            if idx >= 0:
                r = self._display_regions[idx]
                clickable = (r['kind'] == 'segment' or
                             (r['kind'] == 'separator' and r['payload']) or
                             r['kind'] == 'ellipsis')
                self.setCursor(Qt.PointingHandCursor if clickable else Qt.IBeamCursor)
                self.setToolTip(r.get('label') or '')
            else:
                self.setCursor(Qt.IBeamCursor)   # 空白处点击可进入编辑模式
                self.setToolTip('')
            self.update()
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and not self._in_edit:
            self.activated.emit()
            self._press_pos = event.pos()
            self._press_idx = self._region_at(int(event.pos().x()))
            self._dragging = False
        super().mousePressEvent(event)

    def focusInEvent(self, event):
        self.activated.emit()
        super().focusInEvent(event)

    def mouseReleaseEvent(self, event):
        if self._in_edit or event.button() != Qt.LeftButton:
            return super().mouseReleaseEvent(event)
        was_dragging = self._dragging
        press_idx = self._press_idx
        self._press_pos = None
        self._press_idx = -1
        self._dragging = False
        if was_dragging:
            return super().mouseReleaseEvent(event)
        idx = self._region_at(int(event.pos().x()))
        if idx < 0:
            # 空白区域 → 进入编辑模式（与旧行为一致）
            if press_idx < 0:
                self.enter_edit_mode()
            return super().mouseReleaseEvent(event)
        if idx != press_idx:
            # 按下与释放不在同一区，视为取消
            return super().mouseReleaseEvent(event)
        r = self._display_regions[idx]
        kind = r['kind']
        if kind == 'segment':
            if r.get('is_current'):
                self.enter_edit_mode()        # 点击当前(末级)目录 → 编辑
            elif r['payload']:
                self.pathChanged.emit(r['payload'])
        elif kind == 'separator':
            parent = r['payload']
            if parent and os.path.isdir(parent):
                self._show_sibling_menu(parent)
        elif kind == 'ellipsis':
            self.enter_edit_mode()
        super().mouseReleaseEvent(event)

    def leaveEvent(self, event):
        if self._hover_idx != -1:
            self._hover_idx = -1
            self.setToolTip('')
            self.update()
        super().leaveEvent(event)

    def get_path_for_copy(self, separator='\\'):
        p = self._current_path
        if separator != '\\':
            p = p.replace('\\', separator)
        return p

    def _ensure_completer(self):
        """惰性创建路径自动补全器（仅补全目录，类似 Explorer 地址栏）。
        QCompleter 对 QFileSystemModel 有内置的路径拆分支持，能逐级补全 D:\\a\\b。"""
        if self._completer is not None:
            return  # 已创建或已尝试失败
        try:
            from PyQt5.QtWidgets import QCompleter, QFileSystemModel
            model = QFileSystemModel(self)
            model.setRootPath('')
            model.setFilter(QDir.Dirs | QDir.NoDotAndDotDot | QDir.Drives)
            comp = QCompleter(model, self)
            comp.setCaseSensitivity(Qt.CaseInsensitive)
            comp.setCompletionMode(QCompleter.PopupCompletion)
            self._edit.setCompleter(comp)
            self._completer = comp
        except Exception as e:
            debug_print(f"[SimplePathBar] completer init failed: {e}")
            self._completer = False  # 标记已尝试，不重试

    def _show_sibling_menu(self, parent_path):
        """弹出 parent_path 下的子文件夹菜单，选择后导航到该文件夹。"""
        try:
            from PyQt5.QtGui import QCursor
            entries = []
            try:
                with os.scandir(parent_path) as it:
                    for e in it:
                        try:
                            if e.is_dir():
                                entries.append(e.name)
                        except Exception:
                            pass
            except Exception as e:
                debug_print(f"[SimplePathBar] sibling scandir error: {e}")
                return
            entries.sort(key=lambda s: s.lower())
            menu = QMenu(self)
            if not entries:
                act = menu.addAction(tr("无子文件夹"))
                act.setEnabled(False)
            else:
                cur_norm = os.path.normcase(os.path.normpath(self._current_path or ''))
                for name in entries[:300]:
                    full = os.path.join(parent_path, name)
                    act = menu.addAction(name)
                    try:
                        full_norm = os.path.normcase(os.path.normpath(full))
                        if cur_norm == full_norm or cur_norm.startswith(full_norm + os.sep):
                            f = act.font(); f.setBold(True); act.setFont(f)
                    except Exception:
                        pass
                    act.triggered.connect(lambda _=False, p=full: self.pathChanged.emit(p))
            menu.exec_(QCursor.pos())
        except Exception as e:
            debug_print(f"[SimplePathBar] sibling menu error: {e}")

    def _elide_crumb_text(self, text, is_current=False, max_text_width=None):
        try:
            _ = is_current  # 保留参数以兼容现有调用语义
            s = str(text)
            if max_text_width is None:
                return s
            width = int(max_text_width)
            if width <= 0:
                return ''
            fm = self.fontMetrics()
            if fm.horizontalAdvance(s) <= width:
                return s
            return fm.elidedText(s, Qt.ElideMiddle, width)
        except Exception:
            return text

    def _choose_display_parts(self, parts):
        """根据可用宽度选择要显示的路径尾部，优先显示当前目录。"""
        try:
            if not parts:
                return []
            if len(parts) <= 2:
                return parts

            fm = self.fontMetrics()
            avail = max(120, int(self.width() or 0) - 12)
            sep_w = 18
            ellipsis_w = max(24, fm.horizontalAdvance('...') + 8)

            def part_width(label, is_current=False):
                shown = self._elide_crumb_text(label, is_current=is_current)
                return max(24, fm.horizontalAdvance(shown) + 16)

            def total_parts_width(items):
                total = 0
                for idx, (label, _full_path) in enumerate(items):
                    total += part_width(label, is_current=(idx == len(items) - 1))
                    if idx > 0:
                        total += sep_w
                return total

            # 先判断完整路径是否本来就放得下。之前的增量算法在尝试从右向左保留时，
            # 会过早为左侧 "..." 预留宽度，导致“其实足够宽却被折叠”。
            if total_parts_width(parts) <= avail:
                return parts

            # 从最后一级开始向左保留，确保当前目录优先可见。
            selected_rev = []
            used = 0
            for rev_idx, (label, full_path) in enumerate(reversed(parts)):
                is_current = (rev_idx == 0)
                w = part_width(label, is_current=is_current)
                extra = w if not selected_rev else (sep_w + w)
                if selected_rev and used + extra + ellipsis_w > avail:
                    break
                if not selected_rev and w > avail:
                    selected_rev.append((label, full_path))
                    used = min(w, avail)
                    break
                selected_rev.append((label, full_path))
                used += extra

            selected = list(reversed(selected_rev))
            if len(selected) == len(parts):
                return parts

            # 若仍有空间，尽量把根目录也保留在 ... 后面，提升定位感。
            root = parts[0]
            if selected and selected[0][1] != root[1]:
                root_w = part_width(root[0], is_current=False)
                needed = used + ellipsis_w + sep_w + root_w + sep_w
                if needed <= avail:
                    return [root] + selected

            return selected
        except Exception:
            return parts

    def _split_path(self, path):
        """将路径拆分为 [(显示名, 完整路径), …]。"""
        if not path:
            return []
        if path.startswith('shell:') or '::' in path:
            return [(path, path)]
        norm = os.path.normpath(path)
        # splitdrive 可正确识别本地盘符（'C:'）与 UNC 共享根（'\\\\server\\share'），
        # 直接按 os.sep 拆分会丢失 UNC 的 '\\\\' 前缀，导致面包屑各段路径失效。
        drive, tail = os.path.splitdrive(norm)
        result = []
        if drive:
            root = drive + os.sep
            result.append((drive, root))
            current = root
            parts = [p for p in tail.split(os.sep) if p]
        else:
            parts = [p for p in norm.split(os.sep) if p]
            if not parts:
                return []
            current = parts[0] + os.sep
            result.append((parts[0], current))
            parts = parts[1:]
        for part in parts:
            current = os.path.join(current, part)
            result.append((part, current))
        return result

    def _commit_edit(self):
        path = self._edit.text().strip()
        self.exit_edit_mode()
        if path:
            self.pathChanged.emit(path)

    def resizeEvent(self, event):
        # 单控件自绘：尺寸变化时重定位编辑框覆盖层并请求重绘（无子控件重建）
        if self._in_edit:
            self._edit.setGeometry(self.rect())
        self.update()
        super().resizeEvent(event)

    def eventFilter(self, obj, event):
        if obj is self._edit and event.type() in (QEvent.FocusIn, QEvent.MouseButtonPress):
            self.activated.emit()
        if obj is self._edit and event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Escape:
                debug_print(f"[SimplePathBar] eventFilter: Escape pressed, exit edit mode")
                self.exit_edit_mode()
                return True
        # 应用级事件过滤：点击编辑框外部时退出编辑模式
        if self._in_edit and event.type() == QEvent.MouseButtonPress:
            try:
                # 自动补全下拉框可见时，点击下拉项不应退出编辑模式
                if self._completer and self._completer.popup() and self._completer.popup().isVisible():
                    return super().eventFilter(obj, event)
                global_pos = event.globalPos()
                edit_rect = self._edit.rect()
                local_pos = self._edit.mapFromGlobal(global_pos)
                if not edit_rect.contains(local_pos):
                    # 额外检查：点击是否在整个 SimplePathBar 区域内（如果在路径栏区域内不退出）
                    bar_local = self.mapFromGlobal(global_pos)
                    if self.rect().contains(bar_local):
                        debug_print(f"[SimplePathBar] eventFilter: click inside path bar area but outside edit, ignoring exit")
                    else:
                        debug_print(f"[SimplePathBar] eventFilter: click outside path bar, exit edit mode. global_pos={global_pos.x()},{global_pos.y()}")
                        self.exit_edit_mode()
            except Exception:
                pass
        return super().eventFilter(obj, event)

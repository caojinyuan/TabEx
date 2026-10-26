"""Screenshot overlay: selection geometry, annotation rendering and session guards."""

import unittest
from unittest.mock import patch

from PyQt5.QtCore import QEvent, QPoint, QPointF, QRect, Qt
from PyQt5.QtGui import QColor, QMouseEvent, QPixmap
from PyQt5.QtWidgets import QApplication

from tabexplorer import screenshot
from tabexplorer.hotkeys import _hotkey_bindings, _parse_hotkey, _validate_hotkey_bindings


class _FakeScreen:
    def __init__(self, geometry, dpr):
        self._geometry, self._dpr = geometry, dpr

    def geometry(self):
        return self._geometry

    def devicePixelRatio(self):
        return self._dpr


def _send(widget, kind, pos, button=Qt.LeftButton, buttons=Qt.LeftButton):
    event = QMouseEvent(kind, QPointF(pos), QPointF(widget.mapToGlobal(pos)), button, buttons, Qt.NoModifier)
    QApplication.sendEvent(widget, event)


def _drag(widget, start, end):
    _send(widget, QEvent.MouseButtonPress, start)
    _send(widget, QEvent.MouseMove, QPoint((start.x() + end.x()) // 2, (start.y() + end.y()) // 2), Qt.NoButton)
    _send(widget, QEvent.MouseMove, end, Qt.NoButton)
    _send(widget, QEvent.MouseButtonRelease, end, Qt.LeftButton, Qt.NoButton)


class ScreenshotTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def make_overlay(self, dpr=1.0, windows=()):
        geometry = QRect(100, 50, 400, 300)
        physical = QPixmap(round(400 * dpr), round(300 * dpr))
        physical.fill(QColor('#ffffff'))
        physical.setDevicePixelRatio(dpr)
        overlay = screenshot.ScreenshotOverlay(self, _FakeScreen(geometry, dpr), physical, list(windows))
        self.addCleanup(overlay.deleteLater)
        return overlay

    def notify(self, message, level='info'):
        self.messages.append(message)

    def setUp(self):
        self.messages = []

    def test_screen_map_converts_between_native_pixels_and_widget_coordinates(self):
        """150% 缩放且屏幕不在原点时，物理窗口矩形与叠加层逻辑坐标互相换算一致。"""
        mapping = screenshot._ScreenMap(QRect(1920, 0, 1280, 720), 1.5)
        self.assertEqual(mapping.rect_to_widget(QRect(1920 + 150, 30, 300, 150)), QRect(100, 20, 200, 100))
        self.assertEqual(mapping.point_to_physical(QPoint(100, 20)), (1920 + 150, 30))

    def test_drag_selection_matches_the_dragged_size_and_shows_toolbar(self):
        """拖拽 (20,30)→(220,130) 得到 200×100 的选区，并出现工具栏。"""
        overlay = self.make_overlay()
        _drag(overlay, QPoint(20, 30), QPoint(220, 130))
        self.assertEqual(overlay._sel, QRect(20, 30, 200, 100))
        self.assertIsNotNone(overlay._toolbar)

    def test_click_selects_the_detected_window_and_wheel_walks_the_hierarchy(self):
        """单击选中悬停的窗口矩形；滚轮在窗口与子控件层级间切换。"""
        windows = [(1, QRect(100 + 40, 50 + 40, 200, 150))]
        children = [QRect(100 + 60, 50 + 60, 80, 40)]
        overlay = self.make_overlay(windows=windows)
        with patch.object(screenshot, '_child_rects', return_value=children):
            _send(overlay, QEvent.MouseMove, QPoint(70, 70), Qt.NoButton, Qt.NoButton)
            self.assertEqual(overlay._chain, [QRect(40, 40, 200, 150), QRect(60, 60, 80, 40)])
            self.assertEqual(overlay._chain_index, 1)
            overlay._chain_index = 0
            _send(overlay, QEvent.MouseButtonPress, QPoint(70, 70))
            _send(overlay, QEvent.MouseButtonRelease, QPoint(70, 70), Qt.LeftButton, Qt.NoButton)
        self.assertEqual(overlay._sel, QRect(40, 40, 200, 150))

    def test_render_result_uses_physical_pixels_and_burns_annotations(self):
        """125% 缩放下输出为物理像素尺寸，矩形标注边框被烧录进图片。"""
        overlay = self.make_overlay(dpr=1.25)
        _drag(overlay, QPoint(40, 40), QPoint(240, 140))
        overlay._tool, overlay._color, overlay._width = 'rect', QColor('#FF0000'), 4
        _drag(overlay, QPoint(60, 60), QPoint(160, 120))
        image = overlay.render_result().toImage()
        self.assertEqual((image.width(), image.height()), (250, 125))
        self.assertEqual(image.pixelColor(125, 80), QColor('#ffffff'))
        edge = image.pixelColor(round(20 * 1.25), 50)
        self.assertEqual((edge.red(), edge.green(), edge.blue()), (255, 0, 0))

    def test_undo_removes_last_annotation_and_numbers_stay_sequential(self):
        """序号标注依次编号；撤销后下一个序号顺延回收的编号。"""
        overlay = self.make_overlay()
        _drag(overlay, QPoint(10, 10), QPoint(200, 200))
        overlay._tool = 'number'
        for x in (30, 60, 90):
            _send(overlay, QEvent.MouseButtonPress, QPoint(x, 100))
            _send(overlay, QEvent.MouseButtonRelease, QPoint(x, 100), Qt.LeftButton, Qt.NoButton)
        self.assertEqual([item.number for item in overlay._items], [1, 2, 3])
        overlay._undo()
        _send(overlay, QEvent.MouseButtonPress, QPoint(120, 100))
        _send(overlay, QEvent.MouseButtonRelease, QPoint(120, 100), Qt.LeftButton, Qt.NoButton)
        self.assertEqual([item.number for item in overlay._items], [1, 2, 3])

    def test_mosaic_pixelates_only_the_brushed_area(self):
        """马赛克只改变笔刷经过的区域，选区其余像素保持原样。"""
        overlay = self.make_overlay()
        pixmap = overlay._shot
        from PyQt5.QtGui import QPainter
        painter = QPainter(pixmap)
        painter.fillRect(0, 0, 205, 300, QColor('#000000'))
        painter.end()
        overlay._image = pixmap.toImage()
        _drag(overlay, QPoint(0, 0), QPoint(400, 300))
        overlay._tool, overlay._width = 'mosaic', 2
        _drag(overlay, QPoint(190, 150), QPoint(210, 150))
        image = overlay.render_result().toImage()
        self.assertEqual(image.pixelColor(20, 20), QColor('#000000'))
        self.assertEqual(image.pixelColor(380, 20), QColor('#ffffff'))
        blended = image.pixelColor(204, 150)
        self.assertTrue(0 < blended.red() < 255, blended.name())

    def test_right_click_resets_selection_then_closes(self):
        """右键先清除选区与标注，再次右键退出截图。"""
        overlay = self.make_overlay()
        closed = []
        overlay.closed.connect(lambda: closed.append(True))
        _drag(overlay, QPoint(10, 10), QPoint(100, 100))
        _send(overlay, QEvent.MouseButtonPress, QPoint(50, 50), Qt.RightButton, Qt.RightButton)
        self.assertIsNone(overlay._sel)
        self.assertEqual(closed, [])
        _send(overlay, QEvent.MouseButtonPress, QPoint(50, 50), Qt.RightButton, Qt.RightButton)
        self.assertEqual(closed, [True])

    def test_default_screenshot_hotkey_is_valid_and_global_capable(self):
        """默认截图键可解析、不与其他命令冲突，且带修饰键以便注册为全局热键。"""
        bindings = _hotkey_bindings({})
        self.assertEqual(_validate_hotkey_bindings(bindings), [])
        ctrl, shift, alt, _vk = _parse_hotkey(bindings['screenshot'])
        self.assertTrue(ctrl or shift or alt)

    def test_hotkey_registrations_are_independent_per_controller(self):
        """一个控制器注销热键不会释放另一个控制器的注册（各自使用独立的热键 ID）。"""
        first, second, probe = (screenshot._GlobalHotkey(lambda: None) for _ in range(3))
        combo = (True, True, True, 0x7B)  # Ctrl+Shift+Alt+F12
        self.addCleanup(lambda: [item.unregister() for item in (first, second, probe)])
        self.assertTrue(first.register(*combo))
        self.assertFalse(second.register(*combo))
        second.unregister()
        self.assertFalse(probe.register(*combo))
        first.unregister()
        self.assertTrue(probe.register(*combo))

    def test_controller_ignores_a_second_start_while_a_session_is_pending(self):
        """热键与菜单同时触发时只会开启一个截图会话。"""
        controller = screenshot.ScreenshotController(self.notify)
        starts = []
        with patch.object(controller, '_start_now', side_effect=lambda: (starts.append(1), setattr(controller, '_pending', False))):
            controller.start(0)
            controller.start(0)
            self.app.processEvents()
            self.app.processEvents()
        self.assertEqual(starts, [1])
        controller.shutdown()


if __name__ == '__main__':
    unittest.main()

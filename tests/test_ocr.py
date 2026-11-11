"""RapidOCR: offline mixed text, row ordering, process cancellation and Qt lifetime."""

import asyncio
import builtins
import multiprocessing
from pathlib import Path
import runpy
import sys
import time
import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock, patch

from PyQt5.QtCore import QPointF, QRect, Qt
from PyQt5.QtGui import QColor, QFont, QImage, QPainter, QPixmap
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication, QDialog

from tabexplorer import ocr, ocr_backend, screenshot
from tabexplorer.i18n import _LANG_EN


def _sample_image():
    image = QImage(1100, 220, QImage.Format_ARGB32)
    image.fill(Qt.white)
    painter = QPainter(image)
    painter.setFont(QFont('Microsoft YaHei', 28))
    painter.drawText(30, 75, '离线文字识别 中文测试')
    painter.drawText(30, 155, 'Hello World 12345')
    painter.end()
    return image


class OcrTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def wait_for(self, condition):
        deadline = time.monotonic() + 20
        while not condition() and time.monotonic() < deadline:
            QTest.qWait(10)
        self.assertTrue(condition(), 'Timed out waiting for OCR worker')

    def make_dialog(self):
        dialog = screenshot.OcrDialog(_sample_image())
        self.addCleanup(dialog.deleteLater)
        self.addCleanup(dialog.shutdown)
        return dialog

    def test_native_chinese_mixed_and_english_recognition(self):
        image = _sample_image()
        text = asyncio.run(ocr.recognize_image(image))
        self.assertIn('离线文字识别', text)
        self.assertIn('中文测试', text)
        self.assertIn('Hello World 12345', text)
        dialog = self.make_dialog()
        self.wait_for(lambda: bool(dialog.status.text()) and dialog._worker is None)
        self.assertEqual(dialog.text.toPlainText(), text)

    def test_image_scaling_preserves_aspect_ratio_and_physical_pixels(self):
        image = QImage(4000, 2000, QImage.Format_RGB888)
        image.setDevicePixelRatio(2)
        prepared = ocr._prepare_image(image, 2600)
        self.assertEqual((prepared.width(), prepared.height()), (2600, 1300))
        self.assertEqual(prepared.format(), QImage.Format_RGB888)
        self.assertEqual(ocr._prepare_image(_sample_image(), 2600).size(), _sample_image().size())
        with self.assertRaises(ocr.OcrError):
            ocr._prepare_image(QImage(), 2600)

    def test_unsupported_language_is_rejected_before_loading_native_code(self):
        with patch.object(ocr_backend, '_load_engine') as load:
            with self.assertRaisesRegex(ocr.OcrError, 'xx'):
                asyncio.run(ocr.recognize_image(_sample_image(), 'xx'))
            load.assert_not_called()

    def test_missing_models_are_explicit_and_never_downloaded(self):
        import rapidocr
        with patch.object(ocr_backend.Path, 'is_file', return_value=False), \
                patch.object(rapidocr, 'RapidOCR') as engine:
            with self.assertRaisesRegex(ocr.OcrError, 'PP-OCRv6_det_small'):
                ocr_backend._load_engine()
            engine.assert_not_called()

    def test_missing_bindings_do_not_prevent_screenshot_module_loading(self):
        original_import = builtins.__import__

        def missing(name, *args, **kwargs):
            if name == 'rapidocr':
                raise ModuleNotFoundError(name)
            return original_import(name, *args, **kwargs)

        with patch('builtins.__import__', side_effect=missing):
            with self.assertRaisesRegex(ocr.OcrError, 'EXE'):
                ocr_backend._load_engine()

    def test_all_engine_models_have_explicit_existing_paths(self):
        import rapidocr
        with patch.object(rapidocr, 'RapidOCR') as engine:
            ocr_backend._load_engine()
        params = engine.call_args.kwargs['params']
        for stage in ocr_backend.MODEL_FILES:
            self.assertTrue(Path(params[f'{stage}.model_path']).is_file())
        self.assertEqual(params['EngineConfig.onnxruntime.intra_op_num_threads'], 2)

    def test_frozen_runtime_hook_only_initializes_onnx_in_ocr_child(self):
        hook = Path(__file__).resolve().parents[1] / 'tools' / 'pyi_rth_ocr.py'
        original_import = builtins.__import__

        for arguments, expected_calls in ((['TabExplorer.exe'], 0),
                                          (['TabExplorer.exe', '--multiprocessing-fork'], 1)):
            native_imports = []

            def track(name, *args, **kwargs):
                if name == 'onnxruntime':
                    native_imports.append(name)
                    return Mock()
                return original_import(name, *args, **kwargs)

            with patch.object(sys, 'argv', arguments), patch('builtins.__import__', side_effect=track):
                runpy.run_path(str(hook))
            self.assertEqual(len(native_imports), expected_calls)

    def test_columns_are_joined_by_row_despite_small_vertical_offsets(self):
        def box(x, y):
            return [[x, y], [x + 80, y], [x + 80, y + 14], [x, y + 14]]

        result = SimpleNamespace(
            boxes=[box(300, 11), box(20, 10), box(300, 30), box(20, 32)],
            txts=['2026/11/11', ' a_file.c ', '5 KB', '中文.docx'])
        self.assertEqual(ocr_backend._format_result(result), 'a_file.c\t2026/11/11\n中文.docx\t5 KB')
        self.assertEqual(ocr_backend._format_result(SimpleNamespace(boxes=None, txts=None)), '')

    def test_cancelled_inference_does_not_leave_a_child_process(self):
        before = {process.pid for process in multiprocessing.active_children()}

        async def cancel():
            task = asyncio.create_task(ocr.recognize_image(_sample_image()))
            await asyncio.sleep(0.1)
            task.cancel()
            with self.assertRaises(asyncio.CancelledError):
                await task

        asyncio.run(cancel())
        self.assertEqual({process.pid for process in multiprocessing.active_children()}, before)

    def test_broken_pipe_terminates_live_child_before_joining(self):
        receiver, sender = Mock(), Mock()
        receiver.poll.side_effect = BrokenPipeError('closed')
        process = Mock(pid=123, exitcode=None)
        process.is_alive.return_value = True
        context = Mock()
        context.Pipe.return_value = (receiver, sender)
        context.Process.return_value = process
        calls = Mock()
        calls.attach_mock(process.terminate, 'terminate')
        calls.attach_mock(process.join, 'join')
        with patch.object(ocr.multiprocessing, 'get_context', return_value=context):
            with self.assertRaisesRegex(ocr.OcrError, 'OCR'):
                asyncio.run(ocr.recognize_image(_sample_image()))
        self.assertEqual([call[0] for call in calls.mock_calls], ['terminate', 'join'])
        receiver.close.assert_called_once()
        process.close.assert_called_once()

    def test_toolbar_and_shortcut_send_rendered_physical_selection_to_ocr(self):
        controller = Mock()
        screen = SimpleNamespace(geometry=lambda: QRect(0, 0, 800, 600), devicePixelRatio=lambda: 1.5)
        shot = QPixmap(1200, 900)
        shot.fill(Qt.white)
        shot.setDevicePixelRatio(1.5)
        overlay = screenshot.ScreenshotOverlay(controller, screen, shot, [])
        self.addCleanup(overlay.deleteLater)
        overlay._sel = QRect(20, 20, 300, 100)
        item = screenshot._Item('rect', QColor(Qt.red), 4, QPointF(30, 30))
        item.p2 = QPointF(150, 80)
        overlay._items.append(item)
        overlay._show_toolbar()
        self.assertIn('ocr', overlay._toolbar._action_buttons)
        closed = Mock()
        overlay.closed.connect(closed)
        QTest.keyClick(overlay, Qt.Key_T, Qt.ControlModifier)
        closed.assert_called_once()
        controller.recognize.assert_called_once()
        image = controller.recognize.call_args.args[0]
        self.assertEqual((image.width(), image.height()), (450, 150))
        self.assertEqual(image.pixelColor(15, 50), QColor(Qt.red))

    def test_dialog_is_editable_and_only_copies_on_request(self):
        recognize = AsyncMock(return_value='中文测试\nHello World')
        with patch.object(ocr, 'recognize_image', recognize):
            dialog = self.make_dialog()
            with patch.object(screenshot.QApplication, 'clipboard') as clipboard:
                self.wait_for(lambda: recognize.await_count == 1 and dialog._worker is None)
                self.assertEqual(dialog.text.toPlainText(), '中文测试\nHello World')
                self.assertFalse(dialog.text.isReadOnly())
                self.assertTrue(dialog.copy_button.isEnabled())
                clipboard.assert_not_called()
                dialog.text.setPlainText('已校对\nHello')
                dialog.copy_button.click()
                clipboard.return_value.setText.assert_called_once_with('已校对\nHello')

    def test_dialog_can_retry_after_engine_error(self):
        recognize = AsyncMock(side_effect=[ocr.OcrError('OCR model failure'), 'English 123'])
        with patch.object(ocr, 'recognize_image', recognize):
            dialog = self.make_dialog()
            self.wait_for(lambda: recognize.await_count == 1 and dialog._worker is None)
            self.assertEqual(dialog.status.text(), 'OCR model failure')
            self.assertFalse(dialog.copy_button.isEnabled())
            dialog.recognize_button.click()
            self.wait_for(lambda: recognize.await_count == 2 and dialog._worker is None)
            self.assertEqual(recognize.call_args.args[1], 'auto')
            self.assertEqual(dialog.text.toPlainText(), 'English 123')

    def test_blank_result_is_reported_without_enabling_copy(self):
        with patch.object(ocr, 'recognize_image', AsyncMock(return_value='')):
            dialog = self.make_dialog()
            self.wait_for(lambda: bool(dialog.status.text()) and dialog._worker is None)
            self.assertIn('未识别到文字', dialog.status.text())
            self.assertFalse(dialog.copy_button.isEnabled())

    def test_close_cancels_running_recognition_before_releasing_dialog(self):
        cancelled = []

        async def slow(*args):
            try:
                await asyncio.sleep(10)
            finally:
                cancelled.append(True)

        recognize = AsyncMock(side_effect=slow)
        with patch.object(ocr, 'recognize_image', recognize):
            dialog = self.make_dialog()
            dialog.show()
            finished = Mock()
            dialog.finished.connect(finished)
            self.wait_for(lambda: recognize.await_count == 1)
            self.assertFalse(dialog.recognize_button.isEnabled())
            self.assertFalse(dialog.language.isEnabled())
            dialog.recognize()
            self.assertEqual(recognize.await_count, 1)
            dialog.reject()
            self.assertFalse(dialog.isVisible())
            self.wait_for(lambda: dialog._worker is None)
            self.assertEqual(cancelled, [True])
            finished.assert_called_once_with(QDialog.Rejected)
            self.assertEqual(dialog.text.toPlainText(), '')

    def test_controller_shutdown_waits_for_running_worker_before_deleting_dialog(self):
        async def slow(*args):
            await asyncio.sleep(10)

        recognize = AsyncMock(side_effect=slow)
        controller = screenshot.ScreenshotController()
        self.addCleanup(controller.shutdown)
        with patch.object(ocr, 'recognize_image', recognize):
            controller.recognize(_sample_image())
            dialog = controller._ocr_dialog
            self.wait_for(lambda: recognize.await_count == 1)
            worker = dialog._worker
            controller.shutdown()
            self.assertFalse(worker.isRunning())
            self.assertFalse(dialog.isVisible())
            self.assertIsNone(controller._ocr_dialog)
            QApplication.processEvents()

    def test_timeout_and_native_failure_are_visible(self):
        async def slow(*args):
            await asyncio.sleep(10)

        for effect, expected in ((slow, '超时'), (OSError('native failure'), 'native failure')):
            with self.subTest(effect=effect):
                recognize = AsyncMock(side_effect=effect)
                with patch.object(ocr, 'recognize_image', recognize), patch.object(ocr, 'OCR_TIMEOUT', 0.1):
                    dialog = self.make_dialog()
                    self.wait_for(lambda: recognize.await_count == 1 and dialog._worker is None)
                    self.assertIn(expected, dialog.status.text())
                    self.assertFalse(dialog.copy_button.isEnabled())

    def test_controller_replaces_dialog_and_shutdown_cancels_pending_start(self):
        controller = screenshot.ScreenshotController()
        self.addCleanup(controller.shutdown)
        with patch.object(screenshot.QTimer, 'singleShot'):
            controller.recognize(_sample_image())
            first = controller._ocr_dialog
            controller.recognize(_sample_image())
            self.assertTrue(first._closing)
            self.assertFalse(first.isVisible())
            second = controller._ocr_dialog
            second.reject()
            self.assertIsNone(controller._ocr_dialog)
            controller.recognize(_sample_image())
            third = controller._ocr_dialog
            controller.shutdown()
            third.recognize()
            self.assertIsNone(third._worker)
            self.assertIsNone(controller._ocr_dialog)

    def test_dynamic_ocr_labels_are_translated(self):
        labels = [label for _code, label in ocr.LANGUAGES]
        labels.extend(label for _name, _glyph, label in screenshot.ACTIONS)
        self.assertTrue(all(label in _LANG_EN for label in labels))


if __name__ == '__main__':
    unittest.main()

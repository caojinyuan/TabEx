import os
from pathlib import Path
import tempfile
import time
import types
import unittest
from unittest.mock import Mock, patch
import zipfile

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QImage, QTextDocument
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication

from tabexplorer import preview
from tabexplorer.mainwindow import MainWindow


class PreviewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)

    def test_text_encodings_binary_and_truncated_character(self):
        path = self.root / 'text.txt'
        path.write_bytes('中文 hello'.encode('gb18030'))
        self.assertIn('中文', preview.load_preview(str(path))[1])
        path.write_bytes(b'\xff\xfe' + 'hello 中文'.encode('utf-16-le'))
        self.assertIn('中文', preview.load_preview(str(path))[1])
        path.write_bytes(b'a' * 8 + '中'.encode('utf-8'))
        with patch.object(preview, 'TEXT_LIMIT', 10):
            _, text, notice = preview.load_preview(str(path))
        self.assertEqual(text, 'a' * 8)
        self.assertTrue(notice)
        path.write_bytes(b'MZ\x00binary')
        with self.assertRaises(preview.PreviewError):
            preview.load_preview(str(path))

    def test_image_and_archive_listing(self):
        image = QImage(3, 2, QImage.Format_RGB32)
        image.fill(Qt.red)
        path = self.root / 'image.png'
        self.assertTrue(image.save(str(path)))
        kind, result, _ = preview.load_preview(str(path))
        self.assertEqual((kind, result.width(), result.height()), ('image', 3, 2))
        archive = self.root / 'data.zip'
        with zipfile.ZipFile(archive, 'w') as stream:
            stream.writestr('folder/example.txt', 'not extracted')
        self.assertIn('folder/example.txt', preview.load_preview(str(archive))[1])
        self.assertFalse((self.root / 'folder').exists())

    def test_office_text(self):
        from docx import Document
        image = QImage(3, 2, QImage.Format_RGB32)
        image.fill(Qt.red)
        image_path = self.root / 'embedded.png'
        self.assertTrue(image.save(str(image_path)))
        docx = self.root / 'example.docx'
        document = Document()
        document.add_heading('Hello document', level=1)
        document.add_paragraph().add_run('bold text').bold = True
        document.add_picture(str(image_path))
        document.save(docx)
        kind, (html, images), notice = preview.load_preview(str(docx))
        self.assertEqual(kind, 'word')
        self.assertIn('<h1>Hello document</h1>', html)
        self.assertIn('<strong>bold text</strong>', html)
        self.assertEqual(len(images), 1)
        self.assertFalse(next(iter(images.values())).isNull())
        self.assertTrue(notice)
        dialog = preview.PreviewDialog(str(docx))
        dialog.show()
        self.addCleanup(dialog.close)
        for _ in range(200):
            self.app.processEvents()
            if dialog._kind:
                break
            QTest.qWait(10)
        self.assertEqual(dialog._kind, 'word')
        self.assertFalse(dialog.word.document().resource(
            QTextDocument.ImageResource, QUrl(next(iter(images)))).isNull())
        self.assertTrue(dialog.word.loadResource(
            QTextDocument.ImageResource, QUrl('https://example.invalid/a.png')).isNull())
        dialog.search.setText('bold text')
        dialog.find_next()
        self.assertEqual(dialog.word.textCursor().selectedText(), 'bold text')
        pptx = self.root / 'example.pptx'
        with zipfile.ZipFile(pptx, 'w') as stream:
            stream.writestr('ppt/slides/slide1.xml',
                            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                            '<a:r><a:t>Hello slide</a:t></a:r></a:p>')
        self.assertIn('Hello slide', preview.load_preview(str(pptx))[1])

    def test_excel_real_workbooks_are_read_as_sheets(self):
        from openpyxl import Workbook
        book = Workbook()
        book.active.title = 'First'
        book.active['B2'] = 'Hello cell'
        book.create_sheet('Second')['A1'] = '=1+2'
        xlsx = self.root / 'example.xlsx'
        book.save(xlsx)
        kind, sheets, _ = preview.load_preview(str(xlsx))
        self.assertEqual(kind, 'sheet')
        self.assertEqual([sheet[0] for sheet in sheets], ['First', 'Second'])
        self.assertIn((1, 1, 'Hello cell'), sheets[0][3])
        self.assertIn((0, 0, '=1+2'), sheets[1][3])

        import xlwt
        old = xlwt.Workbook()
        old.add_sheet('Legacy').write(0, 0, 'old Excel')
        xls = self.root / 'old.xls'
        old.save(str(xls))
        self.assertIn((0, 0, 'old Excel'), preview.load_preview(str(xls))[1][0][3])

    def test_pdf_renders_page_and_rejects_encrypted_file(self):
        from pypdf import PdfWriter
        from pypdf.generic import DecodedStreamObject, NameObject
        writer = PdfWriter()
        page = writer.add_blank_page(width=100, height=100)
        stream = DecodedStreamObject()
        stream.set_data(b'1 0 0 rg 10 10 50 50 re f')
        page[NameObject('/Contents')] = writer._add_object(stream)
        writer.add_blank_page(width=100, height=100)
        pdf = self.root / 'example.pdf'
        with pdf.open('wb') as stream:
            writer.write(stream)
        kind, (image, count, scale), _ = preview.load_preview(str(pdf))
        self.assertEqual((kind, count), ('pdf', 2))
        self.assertAlmostEqual(scale, 1.5)
        self.assertGreater(image.pixelColor(30, 90).red(), 200)
        self.assertLess(image.pixelColor(30, 90).green(), 50)
        second, _, _ = preview._render_pdf_page(str(pdf), 1, 1.0)
        self.assertGreater(second.pixelColor(30, 60).green(), 200)
        writer.encrypt('password')
        with pdf.open('wb') as stream:
            writer.write(stream)
        with self.assertRaisesRegex(preview.PreviewError, 'PDF'):
            preview.load_preview(str(pdf))

    def test_space_guard_and_launch_use_isolated_preview_process(self):
        path = str(self.root / 'test.txt')
        pane = types.SimpleNamespace(_get_selected_paths=lambda: [path])
        host = types.SimpleNamespace(config={'language': 'en', 'theme': 'monokai_black'}, get_active_pane=lambda: pane,
                                     _last_typing_key_at=time.monotonic())
        with patch('tabexplorer.mainwindow.file_view_has_focus', return_value=True), \
                patch('tabexplorer.mainwindow.subprocess.Popen') as launch:
            MainWindow._preview_selection(host)
            launch.assert_not_called()
            host._last_typing_key_at = 0
            MainWindow._preview_selection(host)
            command = launch.call_args.args[0]
            self.assertEqual(command[-4:], ['--preview', path, 'en', 'monokai_black'])
        with patch('tabexplorer.mainwindow.file_view_has_focus', return_value=False), \
                patch('tabexplorer.mainwindow.subprocess.Popen') as launch:
            MainWindow._preview_selection(host)
            launch.assert_not_called()

    def test_dialog_search_and_space_inside_search(self):
        path = self.root / 'text.txt'
        path.write_text('first hello world\nsecond hello world', encoding='utf-8')
        dialog = preview.PreviewDialog(str(path))
        dialog.show()
        self.addCleanup(dialog.close)
        for _ in range(100):
            self.app.processEvents()
            if dialog.text.toPlainText():
                break
            QTest.qWait(10)
        self.assertIn('second hello world', dialog.text.toPlainText())
        dialog._loaded(('text', dialog.text.toPlainText(), 'truncated'))
        dialog.search.setFocus()
        QTest.keyClicks(dialog.search, 'hello world')
        self.assertEqual(dialog.search.text(), 'hello world')
        self.assertTrue(dialog.isVisible())
        dialog.find_next()
        self.assertEqual(dialog.text.textCursor().selectedText(), 'hello world')
        self.assertEqual(dialog.status.text(), 'truncated')
        QTest.keyClick(dialog.search, Qt.Key_Escape)
        self.assertFalse(dialog.isVisible())

    def test_code_highlighting_handles_preprocessor_multiline_and_unicode(self):
        from tabexplorer.syntax_preview import highlight_spans
        text = '#define SIZE 4\n/* comment\ncontinued */\nconst char *s = "😀"; int value = 2;\n'
        spans = highlight_spans('example.c', text)
        self.assertTrue(any(category == 'preprocessor' for _, _, category in spans[0]))
        self.assertTrue(any(category == 'comment' for _, _, category in spans[2]))
        int_start = len('const char *s = "😀"; '.encode('utf-16-le')) // 2
        self.assertTrue(any(start == int_start and category == 'keyword' for start, _, category in spans[3]))
        self.assertTrue(highlight_spans('example.py', 'def hello():\n    return 3\n'))

    def test_code_preview_shows_line_numbers_and_current_theme(self):
        from tabexplorer import theme
        previous = theme.current_theme_id()
        theme.set_theme('monokai_black')
        self.addCleanup(theme.set_theme, previous)
        path = self.root / 'example.h'
        path.write_text('#define VALUE 1\nint function(void);\n', encoding='utf-8')
        dialog = preview.PreviewDialog(str(path))
        dialog.show()
        self.addCleanup(dialog.close)
        for _ in range(200):
            self.app.processEvents()
            if dialog._kind:
                break
            QTest.qWait(10)
        self.assertEqual(dialog._kind, 'code')
        self.assertGreater(dialog.text._numbers.width(), 0)
        self.assertIn(theme.color('editor-bg'), dialog.styleSheet())
        self.app.processEvents()
        self.assertTrue(dialog.text.document().firstBlock().layout().formats())

    def test_pdf_dialog_navigation_and_excel_dialog_search(self):
        from pypdf import PdfWriter
        writer = PdfWriter()
        writer.add_blank_page(width=100, height=100)
        writer.add_blank_page(width=100, height=100)
        pdf = self.root / 'pages.pdf'
        with pdf.open('wb') as stream:
            writer.write(stream)
        dialog = preview.PreviewDialog(str(pdf))
        dialog.show()
        self.addCleanup(dialog.close)
        for _ in range(200):
            self.app.processEvents()
            if dialog._kind == 'pdf':
                break
            QTest.qWait(10)
        self.assertEqual(dialog._kind, 'pdf')
        self.assertEqual(dialog._pdf_count, 2)
        dialog.pdf_next.click()
        for _ in range(200):
            self.app.processEvents()
            if not dialog._pdf_rendering:
                break
            QTest.qWait(10)
        self.assertIn('2 / 2', dialog.pdf_page.text())
        self.assertIsNotNone(dialog.image.pixmap())
        dialog.close()

        from openpyxl import Workbook
        book = Workbook()
        book.active['C3'] = 'target'
        path = self.root / 'cells.xlsx'
        book.save(path)
        sheet_dialog = preview.PreviewDialog(str(path))
        sheet_dialog.show()
        self.addCleanup(sheet_dialog.close)
        for _ in range(200):
            self.app.processEvents()
            if sheet_dialog._kind == 'sheet':
                break
            QTest.qWait(10)
        self.assertEqual(sheet_dialog.table.item(2, 2).text(), 'target')
        sheet_dialog.search.setText('target')
        sheet_dialog.find_next()
        self.assertEqual((sheet_dialog.table.currentRow(), sheet_dialog.table.currentColumn()), (2, 2))

    def test_audio_preview_does_not_play_automatically(self):
        import wave
        from PyQt5.QtMultimedia import QMediaPlayer
        path = self.root / 'audio.wav'
        with wave.open(str(path), 'wb') as audio:
            audio.setnchannels(1)
            audio.setsampwidth(2)
            audio.setframerate(8000)
            audio.writeframes(b'\x00\x00' * 800)
        dialog = preview.PreviewDialog(str(path))
        dialog.show()
        self.addCleanup(dialog.close)
        for _ in range(200):
            self.app.processEvents()
            if dialog._kind == 'media':
                break
            QTest.qWait(10)
        self.assertEqual(dialog._kind, 'media')
        self.assertEqual(dialog._media_player.state(), QMediaPlayer.StoppedState)
        self.assertFalse(dialog.media_play.isHidden())

    def test_markdown_is_rendered_and_external_images_are_blocked(self):
        image = QImage(3, 2, QImage.Format_RGB32)
        image.fill(Qt.red)
        self.assertTrue(image.save(str(self.root / 'local.png')))
        path = self.root / 'readme.md'
        path.write_text('# Heading\n\n![local](local.png)\n\n**Important** text', encoding='utf-8')
        dialog = preview.PreviewDialog(str(path))
        dialog.show()
        self.addCleanup(dialog.close)
        for _ in range(200):
            self.app.processEvents()
            if dialog._kind:
                break
            QTest.qWait(10)
        self.assertEqual(dialog._kind, 'markdown')
        self.assertIn('font-weight:600', dialog.markdown.toHtml())
        local = dialog.markdown.loadResource(QTextDocument.ImageResource, QUrl('local.png'))
        self.assertFalse(local.isNull())
        remote = dialog.markdown.loadResource(QTextDocument.ImageResource, QUrl('https://example.invalid/a.png'))
        self.assertTrue(remote.isNull())
        outside = dialog.markdown.loadResource(QTextDocument.ImageResource, QUrl('../outside.png'))
        self.assertTrue(outside.isNull())
        dialog.search.setText('Important')
        dialog.find_next()
        self.assertEqual(dialog.markdown.textCursor().selectedText(), 'Important')

"""Read-only, process-isolated file preview."""

import ctypes
import ctypes.wintypes
import codecs
import math
import os
import re
import sys
import tarfile
import threading
import zipfile
from datetime import datetime
from xml.etree import ElementTree

from PyQt5.QtCore import QByteArray, QBuffer, QEvent, QFileInfo, QIODevice, QObject, QSize, Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QColor, QFontDatabase, QImage, QImageReader, QKeySequence, QPixmap, QTextCursor, QTextDocument
from PyQt5.QtWidgets import (
    QAbstractButton, QApplication, QComboBox, QDialog, QFileIconProvider, QFormLayout, QHBoxLayout, QLabel, QLineEdit,
    QHeaderView, QPlainTextEdit, QPushButton, QScrollArea, QShortcut, QSizePolicy, QTableWidget, QTableWidgetItem, QTextBrowser,
    QVBoxLayout, QWidget,
)

from .debuglog import debug_print
from . import theme
from .i18n import _set_app_language, tr
from .firmware_preview import (
    BINARY_EXTENSIONS, HEX_PAGE_BYTES, IHEX_EXTENSIONS, SREC_EXTENSIONS, read_hex_page, read_records,
)
from .syntax_preview import CodeHighlighter, HIGHLIGHT_LIMIT, PreviewTextEdit, highlight_spans
from .system import format_file_size

TEXT_LIMIT = 5 * 1024 * 1024
DOCUMENT_LIMIT = 32 * 1024 * 1024
PDF_FILE_LIMIT = 256 * 1024 * 1024
XML_LIMIT = 8 * 1024 * 1024
ITEM_LIMIT = 1000
PAGE_LIMIT = 30
SHEET_LIMIT = 30
ROW_LIMIT = 200
COLUMN_LIMIT = 50
PDF_PIXEL_LIMIT = 12_000_000
TYPING_GUARD_SECONDS = 0.8
IMAGE_EXTENSIONS = {
    '.bmp', '.cur', '.gif', '.icns', '.ico', '.jpeg', '.jpg', '.pbm', '.pgm', '.png',
    '.ppm', '.svg', '.svgz', '.tga', '.tif', '.tiff', '.wbmp', '.webp', '.xbm', '.xpm',
}
VIDEO_EXTENSIONS = {'.mp4', '.m4v', '.mov', '.avi', '.mkv', '.wmv', '.mpg', '.mpeg', '.webm', '.mts', '.m2ts',
                    '.3gp', '.flv', '.vob', '.ogv'}
AUDIO_EXTENSIONS = {'.mp3', '.wav', '.flac', '.ogg', '.m4a', '.aac', '.wma', '.opus', '.aif', '.aiff', '.mka',
                    '.mid', '.midi', '.ac3', '.ape', '.amr'}
PE_EXTENSIONS = {'.exe', '.dll', '.sys', '.ocx', '.cpl', '.scr'}
INFO_EXTENSIONS = {'.elf', '.axf', '.msi', '.msix', '.iso', '.dmp', '.lnk', '.com'}
CODE_EXTENSIONS = {
    '.c', '.h', '.cc', '.cpp', '.cxx', '.hpp', '.hxx', '.py', '.js', '.ts', '.java',
    '.cs', '.go', '.rs', '.xml', '.json', '.yaml', '.yml', '.html', '.css', '.sh', '.ps1',
}


class PreviewError(Exception):
    pass


def is_typing_key(vk, ctrl, alt):
    return not (ctrl or alt) and (0x30 <= vk <= 0x5A or 0xBA <= vk <= 0xC0 or 0xDB <= vk <= 0xDF)


class _GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.wintypes.DWORD), ('flags', ctypes.wintypes.DWORD),
        ('hwndActive', ctypes.wintypes.HWND), ('hwndFocus', ctypes.wintypes.HWND),
        ('hwndCapture', ctypes.wintypes.HWND), ('hwndMenuOwner', ctypes.wintypes.HWND),
        ('hwndMoveSize', ctypes.wintypes.HWND), ('hwndCaret', ctypes.wintypes.HWND),
        ('rcCaret', ctypes.wintypes.RECT),
    ]


def file_view_has_focus():
    try:
        user32 = ctypes.windll.user32
        user32.GetForegroundWindow.restype = ctypes.wintypes.HWND
        user32.GetWindowThreadProcessId.argtypes = [ctypes.wintypes.HWND, ctypes.POINTER(ctypes.wintypes.DWORD)]
        user32.GetWindowThreadProcessId.restype = ctypes.wintypes.DWORD
        user32.GetGUIThreadInfo.argtypes = [ctypes.wintypes.DWORD, ctypes.POINTER(_GUITHREADINFO)]
        user32.GetClassNameW.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.LPWSTR, ctypes.c_int]
        foreground = user32.GetForegroundWindow()
        if not foreground:
            return False
        thread = user32.GetWindowThreadProcessId(foreground, None)
        info = _GUITHREADINFO()
        info.cbSize = ctypes.sizeof(info)
        if not user32.GetGUIThreadInfo(thread, ctypes.byref(info)) or not info.hwndFocus:
            return False
        name = ctypes.create_unicode_buffer(128)
        user32.GetClassNameW(info.hwndFocus, name, len(name))
        return name.value.lower() in ('directuihwnd', 'syslistview32', 'shelldll_defview')
    except (OSError, AttributeError) as error:
        debug_print(f"[Preview] focus lookup failed: {error}")
        return False


def _read_xml(archive, name):
    info = archive.getinfo(name)
    if info.file_size > XML_LIMIT:
        raise PreviewError(tr("文档内容过大，无法预览。"))
    return ElementTree.fromstring(archive.read(info))


def _office_text(path, ext):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if ext == '.pptx':
            slides = sorted((name for name in names if re.fullmatch(r'ppt/slides/slide\d+\.xml', name)),
                            key=lambda name: int(re.search(r'\d+', name).group()))
            result = []
            for name in slides[:PAGE_LIMIT]:
                root = _read_xml(archive, name)
                texts = [node.text for node in root.iter(
                    '{http://schemas.openxmlformats.org/drawingml/2006/main}t') if node.text]
                result.append(f'--- {name} ---\n' + '\n'.join(texts))
                if sum(len(item) for item in result) >= TEXT_LIMIT:
                    result.append(tr("内容已截断。"))
                    break
            if len(slides) > PAGE_LIMIT:
                result.append(tr("内容已截断。"))
            return '\n\n'.join(result)
    raise PreviewError(tr("此文件格式不支持预览。"))


def _word_document(path):
    import mammoth

    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        if sum(item.file_size for item in infos) > 128 * 1024 * 1024 or any(
                item.file_size > 32 * 1024 * 1024 for item in infos if item.filename.endswith('.xml')):
            raise PreviewError(tr("文档内容过大，无法预览。"))

    images = {}
    image_bytes = 0

    def convert_image(image):
        nonlocal image_bytes
        if len(images) >= 50:
            return {}
        with image.open() as stream:
            data = stream.read(10 * 1024 * 1024 + 1)
        image_bytes += len(data)
        if len(data) > 10 * 1024 * 1024 or image_bytes > 20 * 1024 * 1024:
            return {}
        buffer = QBuffer()
        buffer.setData(QByteArray(data))
        buffer.open(QIODevice.ReadOnly)
        reader = QImageReader(buffer)
        size = reader.size()
        if not size.isValid() or size.width() * size.height() > 20_000_000:
            return {}
        if max(size.width(), size.height()) > 2000:
            size.scale(QSize(2000, 2000), Qt.KeepAspectRatio)
            reader.setScaledSize(size)
        decoded = reader.read()
        if decoded.isNull():
            return {}
        url = f'tabex-preview-image:{len(images)}'
        images[url] = decoded
        return {'src': url}

    with open(path, 'rb') as stream:
        result = mammoth.convert_to_html(stream, convert_image=mammoth.images.img_element(convert_image),
                                         external_file_access=False, include_embedded_style_map=False)
    if len(result.value) > TEXT_LIMIT:
        raise PreviewError(tr("文档内容过大，无法预览。"))
    return result.value, images


def _spreadsheet(path, ext):
    if os.path.getsize(path) > DOCUMENT_LIMIT:
        raise PreviewError(tr("文档内容过大，无法预览。"))
    sheets, total = [], 0
    if ext == '.xls':
        import xlrd
        book = xlrd.open_workbook(path, on_demand=True)
        try:
            for name in book.sheet_names()[:SHEET_LIMIT]:
                sheet = book.sheet_by_name(name)
                cells = []
                for row in range(min(sheet.nrows, ROW_LIMIT)):
                    for column in range(min(sheet.ncols, COLUMN_LIMIT)):
                        cell = sheet.cell(row, column)
                        if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                            continue
                        value = cell.value
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            value = xlrd.xldate.xldate_as_datetime(value, book.datemode)
                        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                            value = bool(value)
                        elif cell.ctype == xlrd.XL_CELL_ERROR:
                            value = xlrd.error_text_from_code.get(value, str(value))
                        text = str(value)[:10000]
                        total += len(text)
                        if total > TEXT_LIMIT:
                            raise PreviewError(tr("电子表格内容过大，无法预览。"))
                        cells.append((row, column, text))
                sheets.append((name, min(sheet.nrows, ROW_LIMIT), min(sheet.ncols, COLUMN_LIMIT), cells))
                book.unload_sheet(name)
        finally:
            book.release_resources()
    else:
        from openpyxl import load_workbook
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if sum(item.file_size for item in infos) > 128 * 1024 * 1024 or any(
                    item.file_size > 32 * 1024 * 1024 for item in infos if item.filename.endswith('.xml')):
                raise PreviewError(tr("电子表格内容过大，无法预览。"))
        book = load_workbook(path, read_only=True, data_only=False, keep_links=False)
        try:
            cached = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        except Exception:
            book.close()
            raise
        try:
            for sheet in book.worksheets[:SHEET_LIMIT]:
                rows = min(sheet.max_row or ROW_LIMIT, ROW_LIMIT)
                columns = min(sheet.max_column or COLUMN_LIMIT, COLUMN_LIMIT)
                cells = []
                source_rows = sheet.iter_rows(max_row=rows, max_col=columns, values_only=True)
                cached_rows = cached[sheet.title].iter_rows(max_row=rows, max_col=columns, values_only=True)
                for row_index, (row, cached_row) in enumerate(zip(source_rows, cached_rows)):
                    for column, (value, cached_value) in enumerate(zip(row, cached_row)):
                        if value is None:
                            continue
                        if isinstance(value, str) and value.startswith('=') and cached_value is not None:
                            value = cached_value
                        text = str(value)[:10000]
                        total += len(text)
                        if total > TEXT_LIMIT:
                            raise PreviewError(tr("电子表格内容过大，无法预览。"))
                        cells.append((row_index, column, text))
                sheets.append((sheet.title, rows, columns, cells))
        finally:
            cached.close()
            book.close()
    if not sheets:
        raise PreviewError(tr("电子表格没有可预览的工作表。"))
    return sheets


def _render_pdf_page(path, index, zoom):
    import pypdfium2
    try:
        document = pypdfium2.PdfDocument(path)
    except pypdfium2.PdfiumError as error:
        if 'password' in str(error).lower():
            raise PreviewError(tr("加密 PDF 无法预览。")) from error
        raise PreviewError(tr("无法打开 PDF：{}").format(error)) from error
    try:
        count = len(document)
        if not count:
            raise PreviewError(tr("PDF 没有可预览的页面。"))
        if not 0 <= index < count:
            raise PreviewError(tr("PDF 页码无效。"))
        page = document[index]
        try:
            width, height = page.get_size()
            if not math.isfinite(width) or not math.isfinite(height) or width <= 0 or height <= 0:
                raise PreviewError(tr("PDF 页面尺寸无效。"))
            scale = min(zoom, math.sqrt(PDF_PIXEL_LIMIT / (width * height)))
            bitmap = page.render(scale=scale, rev_byteorder=True, draw_annots=True)
            try:
                image = QImage(bytes(bitmap.buffer), bitmap.width, bitmap.height, bitmap.stride, QImage.Format_RGB888).copy()
                if image.isNull():
                    raise PreviewError(tr("无法渲染 PDF 页面。"))
                return image, count, scale
            finally:
                bitmap.close()
        finally:
            page.close()
    finally:
        document.close()


def _pdf_text(path):
    from pypdf import PdfReader
    reader = PdfReader(path)
    if reader.is_encrypted:
        raise PreviewError(tr("加密 PDF 无法预览。"))
    pages = []
    for index, page in enumerate(reader.pages):
        if index >= PAGE_LIMIT:
            break
        pages.append(f'--- {index + 1} ---\n' + (page.extract_text() or ''))
        if sum(len(item) for item in pages) >= TEXT_LIMIT:
            break
    return 'text', '\n\n'.join(pages)[:TEXT_LIMIT], tr("仅显示前 30 页提取的文本，不保留原始排版。")


def _decode_text(data, truncated=False):
    if data.startswith((b'\xff\xfe', b'\xfe\xff')):
        return codecs.getincrementaldecoder('utf-16')().decode(data, final=not truncated)
    if b'\x00' in data:
        raise PreviewError(tr("此文件格式不支持预览。"))
    for encoding in ('utf-8-sig', 'gb18030'):
        try:
            text = codecs.getincrementaldecoder(encoding)().decode(data, final=not truncated)
            if sum(ord(c) < 32 and c not in '\r\n\t\f' for c in text[:4096]) > 20:
                raise PreviewError(tr("此文件格式不支持预览。"))
            return text
        except UnicodeDecodeError:
            continue
    raise PreviewError(tr("无法识别文本编码。"))


def _file_fields(path):
    details = os.stat(path)
    return [
        (tr("文件名"), os.path.basename(path)),
        (tr("类型"), os.path.splitext(path)[1].lstrip('.').upper() or tr("文件")),
        (tr("文件路径"), os.path.abspath(path)),
        (tr("大小"), format_file_size(details.st_size)),
        (tr("修改时间"), datetime.fromtimestamp(details.st_mtime).strftime('%Y-%m-%d %H:%M:%S')),
    ]


def _exe_information(path):
    import pywintypes
    import win32api

    fields = _file_fields(path)
    notice = tr("仅显示文件信息，不会执行程序。")
    try:
        translations = win32api.GetFileVersionInfo(path, '\\VarFileInfo\\Translation') or ()
        found = set()
        names = (
            ('FileDescription', tr("文件说明")),
            ('CompanyName', tr("公司名称")),
            ('FileVersion', tr("文件版本")),
            ('ProductVersion', tr("产品版本")),
        )
        for language, codepage in translations:
            for name, label in names:
                if name in found:
                    continue
                subblock = f'\\StringFileInfo\\{language:04x}{codepage:04x}\\{name}'
                value = win32api.GetFileVersionInfo(path, subblock)
                if value:
                    fields.append((label, str(value)[:1024]))
                    found.add(name)
        if not found:
            notice += ' ' + tr("该文件没有可用的版本信息。")
    except pywintypes.error as error:
        debug_print(f"[Preview] executable version information unavailable: {error}")
        if error.winerror in (1812, 1813, 1814, 1815):
            notice += ' ' + tr("该文件没有可用的版本信息。")
        else:
            notice += ' ' + tr("无法读取版本信息：{}").format(error.strerror)
    return fields, notice


def _media_information(path):
    import pythoncom
    from win32com.propsys import propsys, pscon

    fields = _file_fields(path)
    notice = tr("仅显示媒体信息，不会播放文件。")
    store = None
    initialized = False
    extra = []
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
        store = propsys.SHGetPropertyStoreFromParsingName(os.path.abspath(path))

        def number(key):
            try:
                value = store.GetValue(key).GetValue()
            except pythoncom.com_error as error:
                debug_print(f"[Preview] media property unavailable: {error}")
                return 0
            return value if isinstance(value, int) and value > 0 else 0

        duration = number(pscon.PKEY_Media_Duration)
        if duration:
            ms = duration // 10_000
            extra.append((tr("时长"), f'{ms // 3_600_000:02}:{ms // 60_000 % 60:02}:{ms // 1000 % 60:02}.{ms % 1000:03}'))
        if os.path.splitext(path)[1].lower() in VIDEO_EXTENSIONS:
            width, height = number(pscon.PKEY_Video_FrameWidth), number(pscon.PKEY_Video_FrameHeight)
            if width and height:
                extra.append((tr("分辨率"), f'{width} × {height}'))
            rate = number(pscon.PKEY_Video_FrameRate)
            if rate:
                extra.append((tr("帧率"), f'{rate / 1000:g} fps'))
        for label, key, unit in ((tr("采样率"), pscon.PKEY_Audio_SampleRate, ' Hz'),
                                 (tr("声道数"), pscon.PKEY_Audio_ChannelCount, ''),
                                 (tr("音频码率"), pscon.PKEY_Audio_EncodingBitrate, ' bit/s')):
            value = number(key)
            if value:
                extra.append((label, f'{value}{unit}'))
    except pythoncom.com_error as error:
        debug_print(f"[Preview] media properties unavailable: {error}")
    finally:
        store = None
        if initialized:
            pythoncom.CoUninitialize()
    if not extra:
        notice += ' ' + tr("系统未提供更多媒体属性。")
    return fields + extra, notice


def load_preview(path):
    """Return (kind, content, notice); never execute or modify the selected file."""
    if os.path.isdir(path):
        names = []
        with os.scandir(path) as entries:
            for entry in entries:
                names.append(entry.name + ('/' if entry.is_dir(follow_symlinks=False) else ''))
                if len(names) >= ITEM_LIMIT + 1:
                    break
        return 'text', '\n'.join(names[:ITEM_LIMIT]), tr("内容已截断。") if len(names) > ITEM_LIMIT else ''
    ext = os.path.splitext(path)[1].lower()
    if ext in PE_EXTENSIONS:
        fields, notice = _exe_information(path)
        return 'exe', fields, notice
    if ext in VIDEO_EXTENSIONS | AUDIO_EXTENSIONS:
        fields, notice = _media_information(path)
        return 'media', fields, notice
    if ext in INFO_EXTENSIONS:
        return 'info', _file_fields(path), tr("仅显示文件信息，不会执行程序。")
    if ext in BINARY_EXTENSIONS:
        return 'hex', read_hex_page(path), ''
    if ext in SREC_EXTENSIONS | IHEX_EXTENSIONS:
        rows, notice = read_records(path, ext in SREC_EXTENSIONS)
        return 'records', rows, notice
    if ext in IMAGE_EXTENSIONS:
        reader = QImageReader(path)
        reader.setAutoTransform(True)
        size = reader.size()
        if size.isValid() and size.width() * size.height() > 80_000_000:
            raise PreviewError(tr("图片过大，无法预览。"))
        if size.isValid() and max(size.width(), size.height()) > 3000:
            size.scale(QSize(3000, 3000), Qt.KeepAspectRatio)
            reader.setScaledSize(size)
        image = reader.read()
        if image.isNull():
            raise PreviewError(tr("无法读取图片：{}").format(reader.errorString()))
        return 'image', image, ''
    if ext in ('.xls', '.xlsx', '.xlsm', '.xltx', '.xltm'):
        return 'sheet', _spreadsheet(path, ext), tr("只显示单元格内容；样式、图表和图片可能不同。")
    if ext == '.xlsb':
        raise PreviewError(tr("不支持 XLSB 文件。"))
    if ext in ('.docx', '.pptx'):
        if os.path.getsize(path) > DOCUMENT_LIMIT:
            raise PreviewError(tr("文档内容过大，无法预览。"))
        if ext == '.docx':
            return 'word', _word_document(path), tr("文档为简化排版，部分样式和图片可能不同。")
        content = _office_text(path, ext)
        notice = tr("仅显示提取的文本，不保留原始排版。")
        if len(content) > TEXT_LIMIT:
            notice += ' ' + tr("内容已截断。")
        return 'text', content[:TEXT_LIMIT], notice
    if ext == '.pdf':
        if os.path.getsize(path) > PDF_FILE_LIMIT:
            raise PreviewError(tr("文档内容过大，无法预览。"))
        image, count, scale = _render_pdf_page(path, 0, 1.5)
        return 'pdf', (image, count, scale), ''
    if ext in ('.zip', '.jar', '.apk', '.whl'):
        with zipfile.ZipFile(path) as archive:
            entries = archive.infolist()
            return 'text', '\n'.join(item.filename for item in entries[:ITEM_LIMIT]), \
                tr("内容已截断。") if len(entries) > ITEM_LIMIT else ''
    if ext in ('.tar', '.tgz', '.gz', '.bz2', '.xz') and tarfile.is_tarfile(path):
        names = []
        with tarfile.open(path, 'r:*') as archive:
            for item in archive:
                names.append(item.name)
                if len(names) > ITEM_LIMIT:
                    break
        return 'text', '\n'.join(names[:ITEM_LIMIT]), tr("内容已截断。") if len(names) > ITEM_LIMIT else ''
    with open(path, 'rb') as stream:
        data = stream.read(TEXT_LIMIT + 1)
    truncated = len(data) > TEXT_LIMIT
    if truncated:
        data = data[:TEXT_LIMIT]
    kind = 'markdown' if ext in ('.md', '.markdown') else 'text'
    try:
        text = _decode_text(data, truncated)
    except (PreviewError, UnicodeDecodeError):
        if kind == 'markdown' or ext in CODE_EXTENSIONS or ext in ('.txt', '.log', '.csv', '.tsv'):
            raise
        return 'info', _file_fields(path), tr("没有适用的内容预览，仅显示文件信息。")
    return kind, text, tr("内容已截断。") if truncated else ''


class _MarkdownBrowser(QTextBrowser):
    def __init__(self, path, parent=None):
        super().__init__(parent)
        self._directory = os.path.realpath(os.path.dirname(os.path.abspath(path)))
        self.document().setBaseUrl(QUrl.fromLocalFile(self._directory + os.sep))
        self.setOpenLinks(False)
        self.setOpenExternalLinks(False)

    def loadResource(self, resource_type, name):
        if resource_type != QTextDocument.ImageResource:
            return None
        url = self.document().baseUrl().resolved(name)
        if not url.isLocalFile():
            return QImage()
        path = os.path.realpath(url.toLocalFile())
        try:
            if os.path.commonpath((self._directory, path)) != self._directory or os.path.getsize(path) > 10 * 1024 * 1024:
                return QImage()
            reader = QImageReader(path)
            size = reader.size()
            if not size.isValid() or size.width() * size.height() > 20_000_000:
                return QImage()
            return reader.read()
        except (OSError, ValueError) as error:
            debug_print(f"[Preview] Markdown image unavailable: {error}")
            return QImage()


class _WordBrowser(QTextBrowser):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setOpenLinks(False)
        self.setOpenExternalLinks(False)

    def loadResource(self, resource_type, name):
        return QImage() if resource_type == QTextDocument.ImageResource else None


class _Loader(QObject):
    completed = pyqtSignal(object)
    failed = pyqtSignal(str)

    def run(self, path):
        try:
            result = load_preview(path)
            kind, content, notice = result
            if kind == 'text' and os.path.isfile(path) and os.path.splitext(path)[1].lower() not in (
                    '.docx', '.pptx', '.zip', '.jar', '.apk', '.whl', '.tar', '.tgz', '.gz', '.bz2', '.xz'):
                content = content.replace('\r\n', '\n').replace('\r', '\n')
                spans = highlight_spans(path, content)
                if spans is None and len(content) > HIGHLIGHT_LIMIT and os.path.splitext(path)[1].lower() in CODE_EXTENSIONS:
                    notice = ' '.join(filter(None, (notice, tr("文件过大，已关闭语法高亮。"))))
                result = ('code', (content, spans), notice) if spans else ('text', content, notice)
        except Exception as error:
            debug_print(f"[Preview] load failed: {type(error).__name__}: {error}")
            self.failed.emit(str(error) or type(error).__name__)
        else:
            self.completed.emit(result)

    def run_pdf_text(self, path):
        try:
            result = _pdf_text(path)
        except Exception as error:
            debug_print(f"[Preview] PDF text failed: {type(error).__name__}: {error}")
            self.failed.emit(str(error) or type(error).__name__)
        else:
            self.completed.emit(result)


class _PageLoader(QObject):
    completed = pyqtSignal(object, object)
    failed = pyqtSignal(object, str)

    def run(self, read_page, path, request):
        try:
            result = read_page(path, *request)
        except Exception as error:
            debug_print(f"[Preview] page load failed: {type(error).__name__}: {error}")
            self.failed.emit(request, str(error) or type(error).__name__)
        else:
            self.completed.emit(request, result)


class PreviewDialog(QDialog):
    def __init__(self, path):
        super().__init__()
        self.setWindowFlags(
            Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint |
            Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowSystemMenuHint |
            Qt.WindowStaysOnTopHint
        )
        self._path = path
        self._kind = ''
        self._pdf_index, self._pdf_zoom, self._pdf_count = 0, 1.5, 0
        self._pdf_requested = None
        self._pdf_rendering = False
        self._pdf_text_loaded = False
        self._pdf_text_mode = False
        self._pdf_search_backward = False
        self._search_fresh = True
        self._sheets = []
        filename_text = os.path.basename(os.path.normpath(path)) or path
        self.setWindowTitle(filename_text)
        file_info = QFileInfo(path)
        icon = QFileIconProvider().icon(file_info)
        if icon.isNull():
            debug_print(f"[Preview] file icon unavailable: {path}")
        self.setWindowIcon(icon)
        self.resize(900, 650)
        theme.bind_style(self, """
            QDialog { background: @editor-bg; color: @fg; }
            QPlainTextEdit, QTextBrowser, QTableWidget, QScrollArea { background: @editor-bg; color: @fg;
                border: 1px solid @widget-border; selection-background-color: @selection-bg;
                selection-color: @selection-fg; }
            QLineEdit, QComboBox { background: @input-bg; color: @input-fg;
                border: 1px solid @input-border; border-radius: 4px; padding: 5px 8px; }
            QPushButton { background: @input-bg; color: @fg; border: 1px solid @input-border;
                border-radius: 4px; padding: 5px 10px; }
            QPushButton:hover { border-color: @focus-border; }
            QPushButton:checked { border-color: @focus-border; background: @selection-bg; color: @selection-fg; }
            QLabel { color: @fg; }
            QLabel#previewType { background: @chrome-bg; color: @fg-muted; border-radius: 4px; padding: 3px 8px; }
            QLabel#previewStatus { color: @fg-muted; }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        heading = QHBoxLayout()
        filename = QLabel(filename_text, self)
        filename.setTextFormat(Qt.PlainText)
        filename.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        filename.setToolTip(path)
        title_font = filename.font()
        title_font.setBold(True)
        title_font.setPointSize(max(11, title_font.pointSize()))
        filename.setFont(title_font)
        heading.addWidget(filename, 1)
        file_type = tr("文件夹") if file_info.isDir() else os.path.splitext(path)[1].upper().lstrip('.') or tr("文件")
        badge = QLabel(self)
        badge.setTextFormat(Qt.PlainText)
        badge.setText(file_type[:16])
        badge.setObjectName('previewType')
        heading.addWidget(badge)
        self.copy_path = QPushButton(tr("复制路径"), self)
        self.copy_path.clicked.connect(lambda: QApplication.clipboard().setText(os.path.abspath(self._path)))
        heading.addWidget(self.copy_path)
        layout.addLayout(heading)
        self.search_bar = QWidget(self)
        search_row = QHBoxLayout(self.search_bar)
        search_row.setContentsMargins(0, 0, 0, 0)
        self.search = QLineEdit(self.search_bar)
        self.search.setClearButtonEnabled(True)
        self.search.setPlaceholderText(tr("搜索文本"))
        self.search.setToolTip('Ctrl+F · F3 / Shift+F3')
        previous = QPushButton(tr("上一个"), self)
        following = QPushButton(tr("下一个"), self)
        search_row.addWidget(self.search)
        search_row.addWidget(previous)
        search_row.addWidget(following)
        self.wrap = QPushButton(tr("自动换行"), self.search_bar)
        self.wrap.setCheckable(True)
        search_row.addWidget(self.wrap)
        self.errors_only = QPushButton(tr("只看错误"), self.search_bar)
        self.errors_only.setCheckable(True)
        self.errors_only.toggled.connect(self._filter_records)
        search_row.addWidget(self.errors_only)
        layout.addWidget(self.search_bar)
        self.search_bar.hide()
        self.search_bar.setEnabled(False)
        self.controls = QWidget(self)
        controls = QHBoxLayout(self.controls)
        controls.setContentsMargins(0, 0, 0, 0)
        self.sheets = QComboBox(self.controls)
        self.sheets.currentIndexChanged.connect(self._show_sheet)
        controls.addWidget(self.sheets)
        self.pdf_previous = QPushButton(tr("上一页"), self.controls)
        self.pdf_next = QPushButton(tr("下一页"), self.controls)
        self.pdf_page = QLabel(self.controls)
        self.pdf_smaller = QPushButton(tr("缩小"), self.controls)
        self.pdf_larger = QPushButton(tr("放大"), self.controls)
        self.pdf_text_button = QPushButton(tr("查看 PDF 文本"), self.controls)
        for widget in (self.pdf_previous, self.pdf_page, self.pdf_next, self.pdf_smaller,
                       self.pdf_larger, self.pdf_text_button):
            controls.addWidget(widget)
        self.hex_tools = QWidget(self.controls)
        pager = QHBoxLayout(self.hex_tools)
        pager.setContentsMargins(0, 0, 0, 0)
        self.hex_previous = QPushButton(tr("上一页"), self.hex_tools)
        self.hex_next = QPushButton(tr("下一页"), self.hex_tools)
        self.hex_page_label = QLabel(self.hex_tools)
        self.hex_offset = QLineEdit(self.hex_tools)
        self.hex_offset.setPlaceholderText(tr("十六进制偏移"))
        self.hex_offset.setToolTip(tr("十六进制偏移，例如 0x1000"))
        self.hex_offset.setMaximumWidth(155)
        self.hex_offset.setMaxLength(18)
        go = QPushButton(tr("跳转"), self.hex_tools)
        for widget in (self.hex_previous, self.hex_page_label, self.hex_next, self.hex_offset, go):
            pager.addWidget(widget)
        self.hex_previous.clicked.connect(lambda: self._request_hex(self._hex_page.start - HEX_PAGE_BYTES))
        self.hex_next.clicked.connect(lambda: self._request_hex(self._hex_page.start + HEX_PAGE_BYTES))
        self.hex_offset.returnPressed.connect(self._jump_hex)
        go.clicked.connect(self._jump_hex)
        controls.addWidget(self.hex_tools)
        controls.addStretch()
        layout.addWidget(self.controls)
        self.controls.hide()
        self.pdf_previous.clicked.connect(lambda: self._change_pdf_page(-1))
        self.pdf_next.clicked.connect(lambda: self._change_pdf_page(1))
        self.pdf_smaller.clicked.connect(lambda: self._change_pdf_zoom(1 / 1.25))
        self.pdf_larger.clicked.connect(lambda: self._change_pdf_zoom(1.25))
        self.pdf_text_button.clicked.connect(self._toggle_pdf_text)
        self.status = QLabel(tr("正在加载预览…"), self)
        self.status.setObjectName('previewStatus')
        self.status.setTextFormat(Qt.PlainText)
        self._notice = ''
        self.status.setWordWrap(True)
        self.text = PreviewTextEdit(self)
        self.text.setReadOnly(True)
        self.wrap.toggled.connect(lambda enabled: self.text.setLineWrapMode(
            QPlainTextEdit.WidgetWidth if enabled else QPlainTextEdit.NoWrap))
        layout.addWidget(self.text, 1)
        self.info_panel = QScrollArea(self)
        self.info_panel.setWidgetResizable(True)
        card = QWidget()
        info_layout = QVBoxLayout(card)
        info_layout.setContentsMargins(24, 18, 24, 18)
        self.info_icon = QLabel(card)
        info_layout.addWidget(self.info_icon)
        self.info_fields = QFormLayout()
        self.info_fields.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        self.info_fields.setVerticalSpacing(10)
        info_layout.addLayout(self.info_fields)
        self.copy_info = QPushButton(tr("复制文件信息"), card)
        info_layout.addWidget(self.copy_info, 0, Qt.AlignLeft)
        info_layout.addStretch()
        self.info_panel.setWidget(card)
        self.info_panel.hide()
        layout.addWidget(self.info_panel, 1)
        self.image = QLabel(self)
        self.image.setAlignment(Qt.AlignCenter)
        self.image_scroll = QScrollArea(self)
        self.image_scroll.setWidgetResizable(True)
        self.image_scroll.setWidget(self.image)
        self.image_scroll.hide()
        layout.addWidget(self.image_scroll, 1)
        self.table = QTableWidget(self)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.hide()
        layout.addWidget(self.table, 1)
        self.markdown = _MarkdownBrowser(path, self)
        self.markdown.hide()
        layout.addWidget(self.markdown, 1)
        self.word = _WordBrowser(self)
        self.word.hide()
        layout.addWidget(self.word, 1)
        layout.addWidget(self.status)
        self.search.returnPressed.connect(self.find_next)
        self.search.textChanged.connect(self._reset_search)
        following.clicked.connect(self.find_next)
        previous.clicked.connect(lambda: self.find_next(backward=True))
        shortcut = QShortcut(QKeySequence.Find, self)
        shortcut.activated.connect(self._focus_search)
        self._search_shortcut = shortcut
        self._next_shortcut = QShortcut(QKeySequence('F3'), self, activated=self.find_next)
        self._previous_shortcut = QShortcut(QKeySequence('Shift+F3'), self,
                                            activated=lambda: self.find_next(backward=True))
        for button in self.findChildren(QPushButton):
            button.setAutoDefault(False)
        QApplication.instance().installEventFilter(self)
        self._loader = _Loader()
        self._loader.completed.connect(self._loaded)
        self._loader.failed.connect(self._failed)
        self._thread = threading.Thread(target=self._loader.run, args=(path,), daemon=True)
        self._thread.start()

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress:
            if event.key() == Qt.Key_Escape:
                self.close()
                return True
            if (event.key() == Qt.Key_Space and not isinstance(obj, (QLineEdit, QAbstractButton, QComboBox))
                    and event.modifiers() == Qt.NoModifier):
                self.close()
                return True
        return super().eventFilter(obj, event)

    def keyPressEvent(self, event):
        if (self.search.hasFocus() or self.hex_offset.hasFocus()) and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            event.accept()
            return
        super().keyPressEvent(event)

    def closeEvent(self, event):
        QApplication.instance().removeEventFilter(self)
        super().closeEvent(event)

    def _loaded(self, result):
        kind, content, notice = result
        self._kind = kind
        self._reset_search()
        self._notice = notice
        self.status.setText(notice)
        searchable = kind not in ('image', 'media', 'exe', 'info')
        self.search_bar.setEnabled(searchable)
        self.search_bar.setVisible(searchable)
        for shortcut in (self._search_shortcut, self._previous_shortcut, self._next_shortcut):
            shortcut.setEnabled(searchable)
        self.wrap.setVisible(kind in ('text', 'code'))
        self.wrap.setChecked(kind == 'text')
        self.errors_only.setVisible(kind == 'records')
        self.controls.setVisible(kind in ('hex', 'sheet', 'pdf'))
        self.hex_tools.setVisible(kind == 'hex')
        self.sheets.setVisible(kind == 'sheet')
        for widget in (self.pdf_previous, self.pdf_page, self.pdf_next, self.pdf_smaller,
                       self.pdf_larger, self.pdf_text_button):
            widget.setVisible(kind == 'pdf')
        if kind in ('image', 'pdf'):
            self.text.hide()
            image = content
            if kind == 'pdf':
                image, self._pdf_count, actual_zoom = content
                self._update_pdf_page_label(actual_zoom)
            self.image.setPixmap(QPixmap.fromImage(image))
            self.image_scroll.show()
        elif kind == 'sheet':
            self.text.hide()
            self._sheets = content
            self.sheets.addItems([sheet[0] for sheet in content])
            self._show_sheet(0)
            self.table.show()
        elif kind == 'records':
            self._show_records(content)
        elif kind in ('code', 'hex'):
            font = QFontDatabase.systemFont(QFontDatabase.FixedFont)
            font.setPointSizeF(max(10.0, font.pointSizeF()))
            self.text.setFont(font)
            self.text.setLineWrapMode(QPlainTextEdit.NoWrap)
            if kind == 'hex':
                self.search.setPlaceholderText(tr("搜索本页（十六进制 / ASCII）"))
                self._show_hex_page(content)
            else:
                content, spans = content
                self.text.set_line_numbers_visible(True)
                self.text.setPlainText(content)
                self._highlighter = CodeHighlighter(self.text.document(), spans)
        elif kind == 'markdown':
            self.text.hide()
            self.markdown.setMarkdown(content)
            self.markdown.show()
        elif kind == 'word':
            self.text.hide()
            html, images = content
            for url, image in images.items():
                self.word.document().addResource(QTextDocument.ImageResource, QUrl(url), image)
            self.word.setHtml(html)
            self.word.show()
        elif kind in ('media', 'exe', 'info'):
            self.text.hide()
            self.info_icon.setPixmap(self.windowIcon().pixmap(64, 64))
            for label, value in content:
                field = QLabel(value, self.info_panel)
                field.setTextFormat(Qt.PlainText)
                field.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
                field.setWordWrap(True)
                self.info_fields.addRow(label + ':', field)
            details = '\n'.join(f'{label}: {value}' for label, value in content)
            self.copy_info.clicked.connect(lambda: QApplication.clipboard().setText(details))
            self.info_panel.show()
        else:
            self.text.setPlainText(content)

    def _show_records(self, rows):
        self.text.hide()
        headers = [tr("源行"), tr("记录"), tr("地址 / 参数"), tr("字节数"), tr("数据（HEX）"), tr("检查")]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.verticalHeader().hide()
        self.table.setRowCount(len(rows))
        self.table.setWordWrap(False)
        self.table.setShowGrid(False)
        self.table.setFont(QFontDatabase.systemFont(QFontDatabase.FixedFont))
        roles = {'S0': tr("文件头"), 'S5': tr("记录数"), 'S6': tr("记录数"), '01': tr("文件结束"),
                 '02': tr("段基址"), '04': tr("线性基址")}
        roles.update(dict.fromkeys(('S1', 'S2', 'S3', '00'), tr("数据")))
        roles.update(dict.fromkeys(('S7', 'S8', 'S9', '03', '05'), tr("入口地址")))
        self.errors_only.setEnabled(any(row[-1] != 'OK' for row in rows))
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                item = QTableWidgetItem(f'{value} · {roles[value]}' if c == 1 and value in roles else value)
                if c == 4:
                    item.setToolTip(Qt.convertFromPlainText(value))
                if c == 5 and value != 'OK':
                    item.setForeground(QColor(theme.status_color('#d32f2f')))
                self.table.setItem(r, c, item)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.show()

    def _filter_records(self, errors_only):
        self.table.clearSelection()
        self._reset_search()
        for row in range(self.table.rowCount()):
            self.table.setRowHidden(row, errors_only and self.table.item(row, 5).text() == 'OK')

    def _show_hex_page(self, page):
        self._hex_page = page
        self._reset_search()
        self.text.setPlainText(page.text())
        self.hex_tools.setEnabled(True)
        self.hex_previous.setEnabled(page.start > 0)
        self.hex_next.setEnabled(page.start + len(page.data) < page.size)
        pages = max(1, (page.size + HEX_PAGE_BYTES - 1) // HEX_PAGE_BYTES)
        self.hex_page_label.setText(f'{page.start // HEX_PAGE_BYTES + 1} / {pages}')
        self._notice = (f'0x{page.start:X} – 0x{page.start + len(page.data) - 1:X}  ·  '
                        f'{format_file_size(page.size)}') if page.data else tr("空文件")
        self.status.setText(self._notice)

    def _jump_hex(self):
        value = self.hex_offset.text().strip()
        if not re.fullmatch(r'(?:0[xX])?[0-9A-Fa-f]{1,16}', value):
            self.status.setText(tr("请输入十六进制偏移，例如 0x1000。"))
            return
        self._request_hex(int(value, 16))

    def _request_hex(self, offset):
        if not 0 <= offset < self._hex_page.size:
            self.status.setText(tr("偏移超出文件范围。"))
            return
        if self._hex_page.start <= offset < self._hex_page.start + len(self._hex_page.data):
            self._select_hex_row(offset)
            return
        self.hex_tools.setEnabled(False)
        self.status.setText(tr("正在加载预览…"))
        self._hex_loader = _PageLoader()
        self._hex_loader.completed.connect(self._hex_ready)
        self._hex_loader.failed.connect(self._hex_failed)
        threading.Thread(target=self._hex_loader.run, args=(read_hex_page, self._path, (offset,)), daemon=True).start()

    def _hex_ready(self, request, page):
        self._show_hex_page(page)
        self._select_hex_row(request[0])

    def _select_hex_row(self, offset):
        block = self.text.document().findBlockByNumber(1 + (offset - self._hex_page.start) // 16)
        cursor = QTextCursor(block)
        cursor.select(QTextCursor.LineUnderCursor)
        self.text.setTextCursor(cursor)
        self.text.ensureCursorVisible()
        self.status.setText(self._notice)

    def _hex_failed(self, _request, error):
        self.hex_tools.setEnabled(True)
        self.status.setText(tr("预览失败：{}").format(error))

    def _show_sheet(self, index):
        if not 0 <= index < len(self._sheets):
            return
        self._reset_search()
        from openpyxl.utils import get_column_letter
        _name, rows, columns, cells = self._sheets[index]
        self.table.clear()
        self.table.setRowCount(rows)
        self.table.setColumnCount(columns)
        self.table.setHorizontalHeaderLabels([get_column_letter(index + 1) for index in range(columns)])
        for row, column, value in cells:
            self.table.setItem(row, column, QTableWidgetItem(value))
        self._notice = (tr("只显示单元格内容；样式、图表和图片可能不同。") + ' ' +
                        tr("每表最多显示前 200 行、50 列，最多 30 个工作表。"))
        self.status.setText(self._notice)

    def _update_pdf_page_label(self, actual_zoom=None):
        zoom = self._pdf_zoom if actual_zoom is None else actual_zoom
        self.pdf_page.setText(f'{self._pdf_index + 1} / {self._pdf_count}  ·  {round(zoom * 100)}%')
        self.pdf_previous.setEnabled(self._pdf_index > 0)
        self.pdf_next.setEnabled(self._pdf_index + 1 < self._pdf_count)

    def _change_pdf_page(self, delta):
        self._pdf_index = max(0, min(self._pdf_count - 1, self._pdf_index + delta))
        self._update_pdf_page_label()
        self._request_pdf_page()

    def _change_pdf_zoom(self, factor):
        self._pdf_zoom = max(0.5, min(3.0, self._pdf_zoom * factor))
        self._update_pdf_page_label()
        self._request_pdf_page()

    def _request_pdf_page(self):
        self._pdf_requested = (self._pdf_index, self._pdf_zoom)
        if self._pdf_rendering:
            return
        self._pdf_rendering = True
        self.status.setText(tr("正在加载预览…"))
        self._pdf_loader = _PageLoader()
        self._pdf_loader.completed.connect(self._pdf_page_loaded)
        self._pdf_loader.failed.connect(self._pdf_page_failed)
        self._pdf_thread = threading.Thread(target=self._pdf_loader.run,
                                            args=(_render_pdf_page, self._path, self._pdf_requested), daemon=True)
        self._pdf_thread.start()

    def _pdf_page_loaded(self, request, result):
        self._pdf_rendering = False
        if request == self._pdf_requested:
            image, self._pdf_count, actual_zoom = result
            self.image.setPixmap(QPixmap.fromImage(image))
            self._update_pdf_page_label(actual_zoom)
            self.status.setText(self._pdf_text_notice if self._pdf_text_mode else self._notice)
        else:
            self._request_pdf_page()

    def _pdf_page_failed(self, request, error):
        self._pdf_rendering = False
        if request == self._pdf_requested:
            self.status.setText(tr("预览失败：{}").format(error))
        else:
            self._request_pdf_page()

    def _focus_search(self):
        if self._kind == 'pdf' and not self._pdf_text_mode:
            self._toggle_pdf_text()
        else:
            self.search.setFocus()

    def _toggle_pdf_text(self):
        if self._pdf_text_mode:
            self._pdf_text_mode = False
            self.text.hide()
            self.image_scroll.show()
            self.pdf_text_button.setText(tr("查看 PDF 文本"))
            self.status.setText('')
        elif self._pdf_text_loaded:
            self._show_pdf_text()
        else:
            self.pdf_text_button.setEnabled(False)
            self.status.setText(tr("正在加载预览…"))
            self._pdf_text_loader = _Loader()
            self._pdf_text_loader.completed.connect(self._pdf_text_ready)
            self._pdf_text_loader.failed.connect(self._pdf_text_failed)
            self._pdf_text_thread = threading.Thread(target=self._pdf_text_loader.run_pdf_text,
                                                     args=(self._path,), daemon=True)
            self._pdf_text_thread.start()

    def _pdf_text_ready(self, result):
        _kind, content, notice = result
        self.text.setPlainText(content)
        self._pdf_text_notice = notice
        self._pdf_text_loaded = True
        self.pdf_text_button.setEnabled(True)
        self._show_pdf_text()
        if self.search.text():
            self.find_next(backward=self._pdf_search_backward)
        self._pdf_search_backward = False

    def _pdf_text_failed(self, error):
        self.pdf_text_button.setEnabled(True)
        self.status.setText(tr("预览失败：{}").format(error))

    def _show_pdf_text(self):
        self._pdf_text_mode = True
        self.image_scroll.hide()
        self.text.show()
        self.pdf_text_button.setText(tr("查看 PDF 页面"))
        self.status.setText(self._pdf_text_notice)
        self.search.setFocus()

    def _failed(self, message):
        self.status.setText(tr("预览失败：{}").format(message))
        self.search_bar.hide()
        self.search_bar.setEnabled(False)

    def _reset_search(self):
        self._search_fresh = True

    def find_next(self, backward=False):
        query = self.search.text()
        if not query:
            return
        if self._kind in ('sheet', 'records'):
            matches = sorted((item for item in self.table.findItems(query, Qt.MatchContains)
                              if not self.table.isRowHidden(item.row())), key=lambda item: (item.row(), item.column()))
            if not matches:
                self.status.setText(' '.join(filter(None, (self._notice, tr("未找到匹配内容。")))))
                return
            current = self.table.currentItem()
            position = (matches.index(current) if not self._search_fresh and current in matches
                        else (-1 if not backward else 0))
            self.table.setCurrentItem(matches[(position + (-1 if backward else 1)) % len(matches)])
            self._search_fresh = False
            self.status.setText(self._notice)
            return
        if self._kind == 'pdf' and not self._pdf_text_mode:
            self._pdf_search_backward = backward
            if self.pdf_text_button.isEnabled():
                self._toggle_pdf_text()
            if not self._pdf_text_mode:
                return
        target = self.markdown if self._kind == 'markdown' else self.word if self._kind == 'word' else self.text
        if self._search_fresh:
            cursor = target.textCursor()
            cursor.movePosition(cursor.End if backward else cursor.Start)
            target.setTextCursor(cursor)
            self._search_fresh = False
        flags = QTextDocument.FindBackward if backward else QTextDocument.FindFlags()
        if not target.find(query, flags):
            cursor = target.textCursor()
            cursor.movePosition(cursor.End if backward else cursor.Start)
            target.setTextCursor(cursor)
            if not target.find(query, flags):
                notice = self._pdf_text_notice if self._kind == 'pdf' else self._notice
                self.status.setText(' '.join(filter(None, (notice, tr("未找到匹配内容。")))))
                return
        self.status.setText(self._pdf_text_notice if self._kind == 'pdf' else self._notice)


def main():
    if len(sys.argv) != 5:
        raise ValueError('Expected --preview PATH LANGUAGE THEME')
    _set_app_language(sys.argv[3])
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv[:1])
    from PyQt5.QtGui import QFont
    app.setFont(QFont("Microsoft YaHei UI", 9))
    theme.set_theme(sys.argv[4])
    dialog = PreviewDialog(sys.argv[2])
    dialog.show()
    theme.apply_window_frame(dialog)
    dialog.raise_()
    dialog.activateWindow()
    sys.exit(app.exec_())

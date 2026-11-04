"""Capture real native TabEx UI steps with isolated, disposable example data."""

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time


ROOT = Path(__file__).resolve().parents[1]
DEMOS = ('tabs', 'tab-basics', 'bookmarks', 'folder-drop', 'launcher', 'screenshot',
         'groups', 'split', 'content-search', 'command-palette', 'file-basket')
WIDTH = 1120
HEIGHT = 700
INTERVAL_MS = 200
MAX_BYTES = 3 * 1024 * 1024


def gif_frame_duration(count):
    if count < 1:
        raise ValueError('Cannot encode an empty animation')
    return min(INTERVAL_MS, max(20, 18000 // count // 10 * 10))


def exposed_regions(rect, overlays):
    regions = [rect]
    for cover in overlays:
        remaining = []
        for left, top, right, bottom in regions:
            x1, y1 = max(left, cover[0]), max(top, cover[1])
            x2, y2 = min(right, cover[2]), min(bottom, cover[3])
            if x1 >= x2 or y1 >= y2:
                remaining.append((left, top, right, bottom))
            else:
                remaining.extend(box for box in (
                    (left, top, right, y1), (left, y2, right, bottom),
                    (left, y1, x1, y2), (x2, y1, right, y2),
                ) if box[0] < box[2] and box[1] < box[3])
        regions = remaining
    return regions


def prepare_fixture(data):
    for name in ('Project-A', 'Project-B', 'Documents'):
        folder = data / name
        (folder / 'src').mkdir(parents=True)
        (folder / 'README.txt').write_text(f'{name}: example files for TabEx demos.\n', encoding='utf-8')
        (folder / 'src' / 'main.py').write_text(
            '# Demo source\nmessage = "TabEx demo match"\nprint(message)\n', encoding='utf-8')
        (folder / 'src' / 'notes.txt').write_text('Example notes.\n', encoding='utf-8')
    (data / 'config').mkdir()
    (data / 'config' / 'config.json').write_text(json.dumps({
        'enable_explorer_monitor': False, 'auto_update_check': False,
        'enable_cache_tabs': True, 'cached_tabs': [{'path': str(data / 'Project-A')}],
        'pinned_tabs': [], 'debug_mode': False, 'theme': 'light', 'language': 'zh',
        'show_bottom_statusbar': False, 'hotkeys': {'screenshot': False},
    }), encoding='utf-8')


def prepare_launcher(data):
    compiler = Path(os.environ['WINDIR']) / 'Microsoft.NET' / 'Framework64' / 'v4.0.30319' / 'csc.exe'
    if not compiler.is_file():
        raise RuntimeError('Launcher demo requires the Windows .NET Framework 4 C# compiler')
    folder = data / 'Demo-Tools'
    folder.mkdir()
    executable = folder / 'DemoApp.exe'
    subprocess.run([str(compiler), '/nologo', '/target:winexe', f'/out:{executable}',
                    '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll',
                    str(ROOT / 'tools' / 'demo_application.cs')], check=True, timeout=30)
    import win32com.client
    shortcut = win32com.client.Dispatch('WScript.Shell').CreateShortcut(str(folder / 'DemoApp.lnk'))
    shortcut.TargetPath = str(executable)
    shortcut.WorkingDirectory = str(folder)
    shortcut.Save()


def inspect_gif(path):
    from PIL import Image
    with Image.open(path) as image:
        duration = 0
        for frame in range(image.n_frames):
            image.seek(frame)
            duration += image.info.get('duration', 0)
        if image.n_frames < 2 or not 8000 <= duration <= 20000:
            raise AssertionError(f'{path.name}: expected animation lasting 8-20 seconds, got {duration} ms')
        if image.size != (WIDTH, HEIGHT + 84) or image.info.get('loop') != 0:
            raise AssertionError(f'{path.name}: unexpected dimensions or loop setting')
    size = path.stat().st_size
    if size > MAX_BYTES:
        raise AssertionError(f'{path.name}: {size} bytes exceeds the 3 MiB budget')
    return f'{path.name}: {size / 1024:.0f} KiB, {duration / 1000:.1f}s'


def capture(name, output):
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    from PIL import Image, ImageDraw, ImageFont
    import ctypes
    from ctypes import wintypes
    import win32gui
    import win32api
    import win32con
    import win32process
    from PyQt5.QtCore import QBuffer, QCoreApplication, QEvent, QIODevice, QMimeData, QPoint, Qt, QTimer, QUrl
    from PyQt5.QtGui import QDrag, QFont
    from PyQt5.QtTest import QTest
    from PyQt5.QtWidgets import QApplication, QLabel, QMenu, QInputDialog, QTreeWidgetItemIterator
    from tabexplorer.shellview import _ieb_keyboard_filter
    from tabexplorer.widgets import _build_te_icon
    from tabexplorer.mainwindow import MainWindow

    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication([])
    app.setQuitOnLastWindowClosed(False)
    app.setApplicationName('TabExplorer')
    app.setFont(QFont('Microsoft YaHei UI', 9))
    app.setWindowIcon(_build_te_icon())
    app.installNativeEventFilter(_ieb_keyboard_filter)
    window = MainWindow()
    window.setWindowFlag(Qt.WindowStaysOnTopHint, True)
    screen = app.primaryScreen()
    if screen.devicePixelRatio() != 1:
        raise RuntimeError('Set Windows display scaling to 100% before recording (required for privacy masks)')
    available = screen.availableGeometry()
    if available.width() < WIDTH or available.height() < HEIGHT:
        raise RuntimeError('Recording requires an unlocked desktop of at least 1120 x 700 logical pixels')
    window.setGeometry(available.x() + 40, available.y() + 40, WIDTH, HEIGHT)
    window.show()
    window.raise_()
    window.activateWindow()
    data = Path(os.environ['TABEX_DATA_DIR'])
    frames = []
    caption = ['', '']
    errors = []
    cursor_visible = [False]
    saved_cursor = win32gui.GetCursorPos()
    font_path = Path(os.environ['WINDIR']) / 'Fonts' / 'msyh.ttc'
    title_font = ImageFont.truetype(str(font_path), 22)
    step_font = ImageFont.truetype(str(font_path), 17)
    timer = QTimer()

    def is_demo_process(pid):
        marker = data / 'Demo-Tools' / 'launched.txt'
        if name != 'launcher' or not marker.is_file() or marker.read_text().strip() != str(pid):
            return False
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel.OpenProcess.restype = wintypes.HANDLE
        kernel.QueryFullProcessImageNameW.argtypes = [
            wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        handle = kernel.OpenProcess(0x1000, False, pid)
        if not handle:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            path = ctypes.create_unicode_buffer(32768)
            length = wintypes.DWORD(len(path))
            if not kernel.QueryFullProcessImageNameW(handle, 0, path, ctypes.byref(length)):
                raise ctypes.WinError(ctypes.get_last_error())
            return os.path.normcase(path.value) == os.path.normcase(str(data / 'Demo-Tools' / 'DemoApp.exe'))
        finally:
            kernel.CloseHandle(handle)

    def frame():
        try:
            origin = window.mapToGlobal(window.rect().topLeft())
            if (window.width(), window.height()) != (WIDTH, HEIGHT):
                raise RuntimeError(f'Recording window resized to {window.width()} x {window.height()}')
            rect = (origin.x(), origin.y(), origin.x() + window.width(), origin.y() + window.height())
            blockers = []
            overlays = []

            def check_overlay(other, _):
                if other == int(window.winId()):
                    return False
                if win32gui.IsWindowVisible(other):
                    bounds = win32gui.GetWindowRect(other)
                    left, top, right, bottom = bounds
                    if left < rect[2] and right > rect[0] and top < rect[3] and bottom > rect[1]:
                        pid = win32process.GetWindowThreadProcessId(other)[1]
                        if pid != os.getpid() and not is_demo_process(pid):
                            blockers.append(other)
                        else:
                            visible = wintypes.RECT()
                            result = ctypes.windll.dwmapi.DwmGetWindowAttribute(
                                wintypes.HWND(other), 9, ctypes.byref(visible), ctypes.sizeof(visible))
                            if result != 0:
                                raise OSError(f'Cannot obtain visible overlay bounds: HRESULT {result}')
                            overlays.append((visible.left, visible.top, visible.right, visible.bottom))
                return True

            # EnumWindows is front-to-back; only foreign windows above ours matter.
            win32gui.EnumWindows(check_overlay, None)
            if blockers:
                raise RuntimeError('A foreign window overlaps the recording area; close notifications and retry')
            pixmap = screen.grabWindow(0, origin.x(), origin.y(), window.width(), window.height())
            if pixmap.isNull():
                raise RuntimeError('Desktop capture returned an empty image')
            buffer = QBuffer()
            buffer.open(QIODevice.WriteOnly)
            if not pixmap.save(buffer, 'PNG'):
                raise RuntimeError('Could not encode captured frame')
            import io
            with Image.open(io.BytesIO(bytes(buffer.data()))) as source:
                if source.size != (WIDTH, HEIGHT):
                    raise RuntimeError(f'Desktop returned unexpected capture dimensions: {source.size}')
                image = source.convert('RGB')
            redactions = []
            def find_navigation(child, _):
                if (win32gui.IsWindowVisible(child)
                        and win32gui.GetClassName(child) in ('SysTreeView32', 'NamespaceTreeControl')):
                    redactions.append(win32gui.GetWindowRect(child))
            win32gui.EnumChildWindows(int(window.winId()), find_navigation, None)
            draw = ImageDraw.Draw(image)
            regions = [region for bounds in redactions for region in exposed_regions(bounds, overlays)]
            for left, top, right, bottom in regions:
                box = (max(0, left - origin.x()), max(0, top - origin.y()),
                       min(WIDTH, right - origin.x()), min(HEIGHT, bottom - origin.y()))
                if box[0] < box[2] and box[1] < box[3]:
                    draw.rectangle(box, fill='#eef2f6')
                    if box[2] - box[0] >= 140 and box[3] - box[1] >= 40:
                        draw.text((box[0] + 8, box[1] + 12), '系统目录已遮挡', font=step_font, fill='#52657b')
            if name == 'bookmarks':
                modal = QApplication.activeModalWidget()
                tree = getattr(modal, 'tree', None)
                if tree is not None:
                    iterator = QTreeWidgetItemIterator(tree)
                    while iterator.value():
                        item = iterator.value()
                        url = item.text(1)
                        if url and data.as_posix().casefold() not in url.replace('\\', '/').casefold():
                            area = tree.visualItemRect(item)
                            area.setLeft(tree.columnViewportPosition(1))
                            area.setWidth(tree.columnWidth(1))
                            area = area.intersected(tree.viewport().rect())
                            if not area.isEmpty():
                                point = tree.viewport().mapToGlobal(area.topLeft()) - origin
                                draw.rectangle((point.x(), point.y(), point.x() + area.width(),
                                                point.y() + area.height()), fill='#eef2f6')
                                draw.text((point.x() + 2, point.y()), '路径已遮挡',
                                          font=step_font, fill='#52657b')
                        iterator += 1
            if cursor_visible[0]:
                x, y = win32gui.GetCursorPos()
                x, y = x - origin.x(), y - origin.y()
                if 0 <= x < WIDTH and 0 <= y < HEIGHT:
                    draw.ellipse((x - 10, y - 10, x + 10, y + 10), outline='#ff7200', width=3)
                    draw.line((x, y, x + 10, y + 15), fill='#ff7200', width=3)
            canvas = Image.new('RGB', (WIDTH, HEIGHT + 84), '#14243b')
            canvas.paste(image, (0, 0))
            draw = ImageDraw.Draw(canvas)
            draw.text((20, HEIGHT + 8), caption[0], font=title_font, fill='#ffffff')
            draw.text((20, HEIGHT + 43), caption[1], font=step_font, fill='#bce0ff')
            frames.append(canvas)
        except Exception as error:
            timer.stop()
            errors.append(error)

    def check_errors():
        if errors:
            raise errors[0]

    def wait_until(predicate, timeout=15):
        deadline = time.monotonic() + timeout
        while not predicate():
            check_errors()
            if time.monotonic() > deadline:
                raise TimeoutError(f'{name}: UI did not reach expected state: {caption[1]}')
            QTest.qWait(50)
        check_errors()

    def native_ready(tab, path):
        explorer = getattr(tab, 'explorer', None)
        return (getattr(explorer, '_init_ok', False)
                and os.path.normcase(explorer._url_to_path(explorer.property('LocationURL')) or '')
                == os.path.normcase(str(path)))

    def hold(text, action=None, ready=None, seconds=2.4):
        caption[1] = text
        if action is not None:
            action()
        if ready is not None:
            wait_until(ready)
        QTest.qWait(int(seconds * 1000))
        check_errors()

    def open_tab(folder):
        window.add_new_tab(str(data / folder))
        wait_until(lambda: native_ready(window.get_current_tab_widget(), data / folder))

    def modal_sequence(action, steps):
        singles = []
        def later(delay, callback):
            single = QTimer()
            single.setSingleShot(True)
            def invoke():
                try:
                    callback()
                except Exception as error:
                    errors.append(error)
                    popup = QApplication.activePopupWidget()
                    modal = QApplication.activeModalWidget()
                    if popup is not None:
                        popup.close()
                    if modal is not None:
                        modal.reject()
            single.timeout.connect(invoke)
            single.start(delay)
            singles.append(single)
        for delay, callback in steps:
            later(delay, callback)
        def watchdog():
            raise TimeoutError(f'{name}: modal operation did not finish')
        later(14000, watchdog)
        try:
            action()
        finally:
            for single in singles:
                single.stop()
        check_errors()

    def click_menu(label):
        menu = QApplication.activePopupWidget()
        if not isinstance(menu, QMenu):
            raise AssertionError(f'Expected menu containing {label}')
        action = next((action for action in menu.actions() if label in action.text()), None)
        if action is None:
            raise AssertionError(f'Missing menu action: {label}')
        QTest.mouseClick(menu, Qt.LeftButton, pos=menu.actionGeometry(action).center())

    def move_cursor(point):
        left, top = win32api.GetSystemMetrics(76), win32api.GetSystemMetrics(77)
        width, height = win32api.GetSystemMetrics(78), win32api.GetSystemMetrics(79)
        x = round((point.x() - left) * 65535 / max(1, width - 1))
        y = round((point.y() - top) * 65535 / max(1, height - 1))
        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE | win32con.MOUSEEVENTF_ABSOLUTE | 0x4000,
                             x, y, 0, 0)

    def pointer_move(start, end, pressed=False):
        cursor_visible[0] = True
        move_cursor(start)
        QTest.qWait(50)
        if pressed:
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        try:
            for step in range(1, 13):
                point = start + (end - start) * (step / 12)
                move_cursor(point)
                QTest.qWait(70)
                check_errors()
        finally:
            if pressed:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        QTest.qWait(100)

    def native_gesture(start, end):
        cursor_visible[0] = True
        move_cursor(start)
        QTest.qWait(50)
        stop = threading.Event()
        points = [start + (end - start) * (step / 16) for step in range(1, 17)]
        def feed_input():
            try:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                for point in points:
                    if stop.wait(.08):
                        return
                    move_cursor(point)
                stop.wait(.25)
            except Exception as error:
                errors.append(error)
            finally:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        driver = threading.Thread(target=feed_input, name='demo-mouse-input')
        watchdog = QTimer()
        watchdog.setSingleShot(True)
        def cancel_drag():
            errors.append(TimeoutError(f'{name}: native drag did not finish'))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
            QDrag.cancel()
        watchdog.timeout.connect(cancel_drag)
        watchdog.start(6000)
        try:
            driver.start()
            QTest.qWait(2100)
        finally:
            stop.set()
            driver.join(timeout=2)
            watchdog.stop()
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
            if driver.is_alive():
                raise RuntimeError('Mouse input driver failed to stop')
        check_errors()

    class DemoDragSource(QLabel):
        payload = None
        result = None

        def mousePressEvent(self, event):
            event.accept()

        def mouseMoveEvent(self, event):
            if self.payload is not None and event.buttons() & Qt.LeftButton:
                mime, label = self.payload
                self.payload = None
                drag = QDrag(self)
                drag.setMimeData(mime)
                badge = QLabel(label)
                badge.setStyleSheet('background: #fff4d4; color: #202020; border: 2px solid #f59e0b; padding: 8px;')
                badge.adjustSize()
                drag.setPixmap(badge.grab())
                drag.setHotSpot(QPoint(8, 8))
                self.result = drag.exec_(Qt.CopyAction)
                badge.deleteLater()
            event.accept()

    def drag_mime(source, target, mime, label, target_pos=None):
        source.payload = (mime, label)
        source.result = None
        end = target.mapToGlobal(target_pos or QPoint(min(180, target.width() - 8), target.height() // 2))
        native_gesture(source.mapToGlobal(source.rect().center()), end)
        if source.result != Qt.CopyAction:
            raise AssertionError(f'{label}: native source drag was not accepted')

    def source_card(text):
        card = DemoDragSource(text, window)
        card.setAttribute(Qt.WA_NativeWindow, True)
        card.setGeometry(610, 360, 360, 90)
        card.setStyleSheet('background: #fff4d4; border: 2px solid #f59e0b; padding: 12px; color: #202020;')
        card.show()
        card.raise_()
        return card

    def file_mime(path):
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(path))])
        return mime

    try:
        wait_until(lambda: native_ready(window.get_current_tab_widget(), data / 'Project-A'))
        if name in ('groups', 'split', 'command-palette'):
            open_tab('Project-B')
        QTest.qWait(800)
        window._bring_to_front()
        timer.setInterval(INTERVAL_MS)
        timer.timeout.connect(frame)
        titles = {
            'tabs': '多标签浏览 · 在多个目录间快速切换',
            'groups': '标签分组 · 用颜色整理相关目录',
            'split': '左右分屏 · 同步进入对应子目录',
            'content-search': '内容搜索 · 查找源码中的文字并预览匹配行',
            'command-palette': '功能搜索 · 不记菜单，也能找到操作',
            'file-basket': '文件收集篮 · 跨目录收集引用，不移动原文件',
            'tab-basics': '标签基础操作 · 新建、排序、固定与恢复',
            'bookmarks': '书签管理 · 分类保存常用目录，一键打开',
            'folder-drop': '拖入文件夹 · 拖到标签栏，自动打开新标签',
            'launcher': '标题栏启动区 · 拖入 EXE / 快捷方式，一键启动',
            'screenshot': '屏幕截图 · 拖动选区、标注、撤销与贴图',
        }
        caption[0] = titles[name]
        timer.start()
        if name == 'tabs':
            hold('打开 Project-A；每个标签保留自己的目录')
            hold('新建标签并打开 Project-B（Ctrl+T 可新建标签）', lambda: open_tab('Project-B'))
            hold('再打开 Documents；多个目录同时保留', lambda: open_tab('Documents'))
            hold('点击第一个标签，返回 Project-A',
                 lambda: window.tab_widget.setCurrentIndex(0),
                 lambda: native_ready(window.get_current_tab_widget(), data / 'Project-A'), seconds=3)
        elif name == 'tab-basics':
            hold('新建标签并打开 Project-B；Ctrl+T 可新建', lambda: open_tab('Project-B'), seconds=1.6)
            bar = window.tab_widget.tabBar()
            first = window.content_stack.widget(0)
            second = window.content_stack.widget(1)
            hold('按住 Project-B 标签拖到左边，调整顺序',
                 lambda: pointer_move(bar.mapToGlobal(bar.tabRect(1).center()),
                                      bar.mapToGlobal(QPoint(bar.tabRect(0).left() + 8,
                                                            bar.tabRect(0).center().y())), pressed=True),
                 lambda: window.content_stack.widget(0) is second, seconds=1.2)
            caption[1] = '右键 Project-B 标签，选择“固定”；常用目录留在左侧'
            modal_sequence(lambda: window.tab_context_menu(bar.tabRect(0).center()),
                           [(1200, lambda: click_menu('固定'))])
            hold('Project-B 已固定；点击另一个标签继续浏览',
                 lambda: window.tab_widget.setCurrentIndex(window.content_stack.indexOf(first)),
                 lambda: second.is_pinned, seconds=1.6)
            hold('Ctrl+W 关闭当前 Project-A 标签', window.close_current_tab,
                 lambda: window.tab_widget.count() == 1, seconds=1.6)
            hold('Ctrl+Shift+T 恢复误关的 Project-A 标签', window.reopen_closed_tab,
                 lambda: window.tab_widget.count() == 2
                 and native_ready(window.get_current_tab_widget(), data / 'Project-A'), seconds=2.2)
        elif name == 'bookmarks':
            add_dialog = []
            def rename_bookmark():
                dialog = QApplication.activeModalWidget()
                if dialog is None or not hasattr(dialog, 'name_edit'):
                    raise AssertionError('Add bookmark dialog did not open')
                add_dialog.append(dialog)
                dialog.name_edit.setText('Project-A')
                caption[1] = '添加当前目录；点击“新建文件夹”创建分类'
                dialog._create_folder()
            def name_folder():
                dialog = QApplication.activeModalWidget()
                if not isinstance(dialog, QInputDialog):
                    raise AssertionError('New bookmark folder dialog did not open')
                dialog.setTextValue('Demo Projects')
                dialog.accept()
                caption[1] = '选择 Demo Projects 分类，确认保存 Project-A'
            caption[1] = 'Ctrl+D 将当前目录添加为书签'
            modal_sequence(window.add_current_tab_bookmark, [
                (1800, rename_bookmark), (3000, name_folder),
                (5100, lambda: add_dialog[0]._accept_if_valid())])
            hold('打开书签管理，查看分类与保存的路径', seconds=1.4)
            def inspect_bookmarks():
                dialog = QApplication.activeModalWidget()
                if dialog is None or not hasattr(dialog, 'tree'):
                    raise AssertionError('Bookmark manager did not open')
                dialog.resize(900, 540)
                dialog.move(window.x() + 100, window.y() + 80)
                dialog.tree.setColumnWidth(0, 220)
                dialog.tree.expandAll()
                found = []
                iterator = QTreeWidgetItemIterator(dialog.tree)
                while iterator.value():
                    item = iterator.value()
                    if item.text(0) == 'Project-A':
                        found.append(item)
                    iterator += 1
                if not found or found[0].parent().text(0) != 'Demo Projects':
                    raise AssertionError('Bookmark was not saved inside the chosen folder')
                dialog.tree.setCurrentItem(found[0])
                dialog.tree.scrollToItem(found[0])
            modal_sequence(window.show_bookmark_dialog, [
                (300, inspect_bookmarks), (2800, lambda: QApplication.activeModalWidget().accept())])
            # Execute the actual bookmark-bar menu item, not a direct navigation call.
            def open_bookmark():
                folder = next((action for action in window.menu_bar.actions()
                               if action.menu() and 'Demo Projects' in action.text()), None)
                if folder is None:
                    raise AssertionError('New bookmark folder missing from bookmark bar')
                menu = folder.menu()
                def choose():
                    click_menu('Project-A')
                modal_sequence(lambda: menu.exec_(window.menu_bar.mapToGlobal(QPoint(180, 20))),
                               [(1400, choose)])
            hold('点击书签栏分类中的 Project-A，一键打开目录', open_bookmark,
                 lambda: native_ready(window.get_current_tab_widget(), data / 'Project-A'), seconds=1.8)
        elif name == 'folder-drop':
            card = source_card('示例拖拽源：Project-B 文件夹\n拖动到上方标签栏（不是工具栏）')
            hold('将文件夹拖到标签栏，不必手动输入路径', seconds=2)
            hold('按住文件夹，移动到标签栏并松开',
                 lambda: drag_mime(card, window.tab_widget, file_mime(data / 'Project-B'), 'Project-B'),
                 lambda: window.tab_widget.count() == 2
                 and native_ready(window.get_current_tab_widget(), data / 'Project-B'), seconds=2)
            card.hide()
            hold('Project-B 自动打开为新标签；原目录仍然保留', seconds=2)
            hold('点击 Project-A 标签，返回原目录', lambda: window.tab_widget.setCurrentIndex(0),
                 lambda: native_ready(window.get_current_tab_widget(), data / 'Project-A'), seconds=2)
        elif name == 'launcher':
            executable = data / 'Demo-Tools' / 'DemoApp.exe'
            shortcut = executable.with_suffix('.lnk')
            card = source_card('示例拖拽源：DemoApp.exe\n拖到标题栏“+ 拖入应用或快捷方式”')
            hold('程序和快捷方式可固定在标题栏，一键启动', seconds=1.4)
            hold('拖入 DemoApp.exe：标题栏出现启动按钮',
                 lambda: drag_mime(card, window.title_shortcut_bar._hint_label,
                                   file_mime(executable), 'DemoApp.exe',
                                   window.title_shortcut_bar._hint_label.rect().center()),
                 lambda: executable in map(Path, window.title_shortcut_bar.get_shortcuts()), seconds=1.2)
            card.setText('示例拖拽源：DemoApp.lnk\n快捷方式也可拖入同一个启动区')
            hold('再拖入 DemoApp.lnk：快捷方式同样支持',
                 lambda: drag_mime(card, window.title_shortcut_bar._hint_label,
                                   file_mime(shortcut), 'DemoApp.lnk',
                                   window.title_shortcut_bar._hint_label.rect().center()),
                 lambda: len(window.title_shortcut_bar.get_shortcuts()) == 2, seconds=1.2)
            card.hide()
            def button_for(path):
                from PyQt5.QtWidgets import QToolButton
                return next(button for button in window.title_shortcut_bar.findChildren(QToolButton)
                            if Path(button.property('shortcut_path')) == path)
            def launch():
                button_for(executable).click()
                wait_until(lambda: (executable.parent / 'launched.txt').is_file())
                handles = []
                def find_demo(hwnd, _):
                    if win32gui.IsWindowVisible(hwnd) and is_demo_process(
                            win32process.GetWindowThreadProcessId(hwnd)[1]):
                        handles.append(hwnd)
                    return True
                win32gui.EnumWindows(find_demo, None)
                if not handles:
                    raise AssertionError('Demo executable started without a visible window')
                win32gui.SetWindowPos(handles[0], win32con.HWND_TOPMOST,
                                     window.x() + 600, window.y() + 220, 360, 160,
                                     win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)
            hold('点击 EXE 按钮：真实启动示例程序（稍后自动退出）', launch, seconds=3.2)
            hold('按住快捷方式按钮拖到另一按钮右侧，调整顺序',
                 lambda: native_gesture(button_for(shortcut).mapToGlobal(button_for(shortcut).rect().center()),
                                        window.title_shortcut_bar.mapToGlobal(QPoint(
                                            window.title_shortcut_bar.width() - 2,
                                            window.title_shortcut_bar.height() // 2))),
                 lambda: list(map(Path, window.title_shortcut_bar.get_shortcuts())) == [executable, shortcut],
                 seconds=.8)
            caption[1] = '右键启动按钮，选择“移除该快捷方式”；不删除原文件'
            button = button_for(shortcut)
            modal_sequence(lambda: window.title_shortcut_bar._show_context_menu_for_button(
                button, button.rect().center()), [(1200, lambda: click_menu('移除该快捷方式'))])
            hold('快捷方式已从启动区移除；原 EXE 和 LNK 仍在',
                 ready=lambda: list(map(Path, window.title_shortcut_bar.get_shortcuts())) == [executable]
                 and executable.is_file() and shortcut.is_file(), seconds=1)
        elif name == 'screenshot':
            from tabexplorer import screenshot as screenshot_module
            board = QLabel(window)
            board.setAttribute(Qt.WA_NativeWindow, True)
            board.setGeometry(0, 125, WIDTH, HEIGHT - 125)
            board.setText('TabEx 截图示例\n\n拖动选择区域\n添加矩形、箭头、文字等标注\n不满意可撤销，完成后可保存、复制或贴图')
            board.setAlignment(Qt.AlignCenter)
            board.setStyleSheet('background: #edf5ff; color: #21496c; font-size: 22px;')
            board.show()
            board.raise_()
            hold('Ctrl+Alt+Z 开始截图（演示中通过软件入口启动）', seconds=1.6)
            # The board covers the native tree before the real desktop snapshot.
            cursor_visible[0] = True
            win32api.SetCursorPos((window.x() + 450, window.y() + 300))
            window.start_screenshot()
            wait_until(lambda: window.screenshot._overlay is not None)
            overlay = window.screenshot._overlay
            origin = window.mapToGlobal(QPoint())
            hold('按住左键拖动，选择要截取的区域',
                 lambda: pointer_move(origin + QPoint(260, 180), origin + QPoint(920, 460), pressed=True),
                 lambda: overlay._sel is not None, seconds=1.2)
            toolbar = overlay._toolbar
            def rectangle():
                toolbar._tool_buttons['rect'].click()
                pointer_move(origin + QPoint(400, 245), origin + QPoint(760, 400), pressed=True)
            hold('选择矩形工具，在选区内拖动添加标注', rectangle,
                 lambda: len(overlay._items) == 1 and overlay._items[0].kind == 'rect', seconds=1.2)
            hold('点击撤销：移除刚才的标注', toolbar._action_buttons['undo'].click,
                 lambda: not overlay._items, seconds=1.4)
            def arrow():
                toolbar._tool_buttons['arrow'].click()
                pointer_move(origin + QPoint(330, 260), origin + QPoint(540, 340), pressed=True)
            hold('选择箭头工具，指出需要说明的位置', arrow,
                 lambda: len(overlay._items) == 1 and overlay._items[0].kind == 'arrow', seconds=1.2)
            def pin_screenshot():
                toolbar._action_buttons['pin'].click()
                wait_until(lambda: window.screenshot._overlay is None and len(screenshot_module._pins) == 1)
                QTest.qWait(150)
                pin = screenshot_module._pins[0]
                win32gui.SetWindowPos(int(pin.winId()), win32con.HWND_TOPMOST,
                                     window.x() + 440, window.y() + 200, pin.width(), pin.height(),
                                     win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE)
            hold('点击贴图：截图悬浮桌面；不修改系统剪贴板', pin_screenshot, seconds=2)
            pin = screenshot_module._pins[0]
            if pin._base.isNull():
                raise AssertionError('Pinned screenshot is empty')
            pin.close()
            board.hide()
        elif name == 'groups':
            hold('先打开两个相关项目目录')
            hold('按 F4：当前标签及左侧连续未分组标签组成一组', window.insert_group_btn.click,
                 lambda: bool(window.get_current_tab_widget().bookmark_group_color))
            hold('点击分组色块折叠：标签栏更简洁',
                 lambda: window.toggle_tab_group_collapsed(0, window.tab_widget),
                 lambda: bool(window.content_stack.widget(0).tab_group_collapsed))
            hold('再次点击色块展开：相关标签仍然保留',
                 lambda: window.toggle_tab_group_collapsed(0, window.tab_widget),
                 lambda: not window.content_stack.widget(0).tab_group_collapsed, seconds=3)
        elif name == 'split':
            hold('左右两侧先准备相同结构的 Project-A / Project-B')
            hold('按 F3：当前 Project-B 标签移到右侧', window.split_view_btn.click,
                 lambda: window._split_active)
            hold('点击“同步导航”：开启相对路径跟随', window.sync_browse_button.click,
                 lambda: window._sync_browse)
            left = window.content_stack.widget(0)
            right = window.split_content_stack.widget(0)
            hold('左侧进入 src；右侧自动进入自己的 src，不联动选中项',
                 lambda: left.navigate_to(str(data / 'Project-A' / 'src')),
                 lambda: native_ready(left, data / 'Project-A' / 'src')
                 and native_ready(right, data / 'Project-B' / 'src'), seconds=3)
        elif name == 'content-search':
            def open_search():
                window.search_button.click()
                dialog = window.search_dialogs[-1]
                dialog.resize(1000, 580)
                dialog.move(window.x() + 55, window.y() + 50)
            hold('打开项目目录；Ctrl+F 搜索当前文件夹', open_search)
            dialog = window.search_dialogs[-1]
            def configure_search():
                dialog.use_everything_cb.setChecked(False)
                dialog.search_filename_cb.setChecked(False)
                dialog.search_content_cb.setChecked(True)
                dialog.file_type_input.setText('py')
                dialog.search_input.setEditText('TabEx demo match')
            hold('关闭 Everything，选择文件内容；限定 py 文件', configure_search)
            hold('开始搜索：结果显示匹配文件与行号', dialog.start_search,
                 lambda: dialog.search_thread is not None and not dialog.search_thread.is_alive()
                 and dialog.current_result_count == 1)
            hold('选中结果，直接预览匹配内容',
                 ready=lambda: 'TabEx demo match' in dialog.match_preview.toPlainText(), seconds=3)
            dialog.close()
        elif name == 'command-palette':
            hold('Ctrl+Shift+P 打开功能搜索；也可从“工具”进入')
            modal_timers = []
            def later(delay, callback):
                single = QTimer()
                single.setSingleShot(True)
                def invoke():
                    try:
                        callback()
                    except Exception as error:
                        errors.append(error)
                        modal = QApplication.activeModalWidget()
                        if modal is not None:
                            modal.reject()
                single.timeout.connect(invoke)
                single.start(delay)
                modal_timers.append(single)
            def query():
                modal = QApplication.activeModalWidget()
                if modal is None or not hasattr(modal, 'query'):
                    raise RuntimeError('Command palette did not open')
                caption[1] = '输入“分屏”：筛选匹配功能，并显示可用状态'
                modal.query.setText('分屏')
            def choose():
                modal = QApplication.activeModalWidget()
                if modal is None or not hasattr(modal, 'results'):
                    raise RuntimeError('Command palette closed unexpectedly')
                row = next((modal.results.topLevelItem(i) for i in range(modal.results.topLevelItemCount())
                            if '分屏对比' in modal.results.topLevelItem(i).text(0)), None)
                if row is None:
                    raise AssertionError('Split command missing from palette')
                modal.results.setCurrentItem(row)
                caption[1] = '选择“分屏对比”并回车：直接执行已有功能'
                modal.choose_current()
            caption[1] = '浏览可搜索的功能列表，不需要逐个翻菜单'
            later(2400, query)
            later(5200, choose)
            def watchdog():
                raise TimeoutError('Command palette watchdog')
            later(12000, watchdog)
            window.show_command_palette()
            for single in modal_timers:
                single.stop()
            check_errors()
            hold('功能立即执行：当前标签已移入右侧分屏',
                 ready=lambda: window._split_active, seconds=3)
        elif name == 'file-basket':
            hold('从不同目录收集文件，适合一次整理多个项目')
            paths = [str(data / folder / 'README.txt') for folder in ('Project-A', 'Project-B')]
            def collect():
                window.add_paths_to_basket(paths)
                window.show_file_basket()
                basket = window.get_file_basket()
                basket.move(window.x() + 130, window.y() + 100)
            hold('添加两个目录的 README.txt：统一查看路径引用', collect,
                 lambda: len(window.get_file_basket().all_paths()) == 2)
            hold('重复添加同一路径：自动去重，仍为 2 项',
                 lambda: window.add_paths_to_basket(paths),
                 lambda: len(window.get_file_basket().all_paths()) == 2)
            hold('点击“清空收集篮”：只清引用，原文件仍在',
                 window.get_file_basket().clear_button.click,
                 lambda: not window.get_file_basket().all_paths()
                 and all(Path(path).is_file() for path in paths), seconds=3)
            window.get_file_basket().close()
        timer.stop()
        check_errors()
        if not frames:
            raise RuntimeError('No frames captured')
        palette = frames[0].quantize(colors=256, method=Image.Quantize.MEDIANCUT)
        encoded = [image.quantize(palette=palette, dither=Image.Dither.NONE) for image in frames]
        playback_ms = gif_frame_duration(len(encoded))
        if playback_ms != INTERVAL_MS:
            print(f'{name}: playback accelerated by {INTERVAL_MS / playback_ms:.2f}x '
                  'to keep the demonstration concise (all frames retained)', flush=True)
        with tempfile.TemporaryDirectory(prefix='encode-', dir=output.parent) as stage:
            candidate = Path(stage) / output.name
            encoded[0].save(candidate, save_all=True, append_images=encoded[1:],
                            duration=playback_ms, loop=0, optimize=True, disposal=1)
            print(inspect_gif(candidate), flush=True)
            candidate.replace(output)
    finally:
        timer.stop()
        try:
            if cursor_visible[0]:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                win32api.SetCursorPos(saved_cursor)
        finally:
            window.close()
            window.deleteLater()
            QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
            app.removeNativeEventFilter(_ieb_keyboard_filter)
            QTest.qWait(200)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--only', choices=DEMOS, nargs='+', default=list(DEMOS))
    parser.add_argument('--output', type=Path, default=ROOT / 'docs' / 'demos')
    parser.add_argument('--check', action='store_true', help='Validate existing GIFs without opening TabEx')
    parser.add_argument('--capture', choices=DEMOS, help=argparse.SUPPRESS)
    options = parser.parse_args()
    output = options.output.resolve()
    if options.capture:
        capture(options.capture, output / f'{options.capture}.gif')
        return
    if options.check:
        for name in options.only:
            print(inspect_gif(output / f'{name}.gif'))
        return
    if os.name != 'nt':
        raise RuntimeError('Native TabEx demonstrations require Windows')
    output.mkdir(parents=True, exist_ok=True)
    staging_root = ROOT / 'artifacts'
    staging_root.mkdir(exist_ok=True)
    for name in options.only:
        with tempfile.TemporaryDirectory(prefix='demo-', dir=staging_root) as temporary:
            data = Path(temporary)
            prepare_fixture(data)
            if name == 'launcher':
                prepare_launcher(data)
            environment = dict(os.environ, TABEX_DATA_DIR=str(data), LOCALAPPDATA=str(data / 'local'),
                               PYTHONUTF8='1', QT_LOGGING_RULES='*.debug=false;qt.qpa.*=false')
            subprocess.run([sys.executable, '-u', '-X', 'utf8', str(Path(__file__).resolve()),
                            '--capture', name, '--output', str(output)],
                           cwd=ROOT, env=environment, check=True, timeout=90)


if __name__ == '__main__':
    main()

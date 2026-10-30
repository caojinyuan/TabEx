"""QuickLook 预览集成：首次使用时下载便携版到数据目录，再通过命令行调用预览。

QuickLook 为 GPL v3 程序，不随 TabExplorer 分发，仅在用户同意后从官方 GitHub 发布页下载。"""

import ctypes
import ctypes.wintypes
import hashlib
import os
import shutil
import subprocess
import time
import urllib.request

from PyQt5.QtCore import pyqtSignal, QObject, QThread, QTimer

from .debuglog import debug_print
from .i18n import tr
from .paths import get_app_data_path, get_bundled_path

QUICKLOOK_VERSION = '4.5.0'
QUICKLOOK_URL = f'https://github.com/QL-Win/QuickLook/releases/download/{QUICKLOOK_VERSION}/QuickLook-{QUICKLOOK_VERSION}.7z'
QUICKLOOK_SHA256 = 'd1d578deb75b83ea09255b89382b088d35d38125c0c2b309825b1e20ac962b91'
QUICKLOOK_SIZE = 62569358
_CHUNK_SIZE = 256 * 1024
_NETWORK_TIMEOUT_S = 30

# 焦点控件类名含这些片段时（重命名框、路径栏、目录树、Qt 自身控件）不触发预览
_NON_FILE_VIEW_CLASS_PARTS = ('edit', 'combobox', 'systreeview', 'qt', 'scintilla')
# 连续输入文件名（含空格）时，空格不应当作预览键
TYPING_GUARD_SECONDS = 0.8

_owned_process = None
_last_preview_at = None
# 本程序启动的 QuickLook 常驻后台且内存随预览累积，空闲超过该时长后结束，下次预览时自动重新启动
QUICKLOOK_IDLE_TIMEOUT_S = 20 * 60


class QuickLookCancelled(Exception):
    pass


class QuickLookInstallError(Exception):
    pass


def quicklook_dir():
    return get_app_data_path('QuickLook')


def find_quicklook():
    """返回 QuickLook.exe 路径：优先使用本程序下载的便携版，其次是系统中已安装的版本。"""
    candidates = [os.path.join(quicklook_dir(), 'QuickLook.exe')]
    for variable, subdirectory in (('LOCALAPPDATA', 'Programs'), ('ProgramFiles', ''), ('ProgramFiles(x86)', '')):
        base = os.environ.get(variable)
        if base:
            candidates.append(os.path.join(base, subdirectory, 'QuickLook', 'QuickLook.exe'))
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    return None


def is_file_view_focus_class(class_name):
    """焦点控件类名是否属于文件列表（其余情况按文本输入处理，不抢空格键）。"""
    lowered = str(class_name or '').lower()
    return bool(lowered) and not any(part in lowered for part in _NON_FILE_VIEW_CLASS_PARTS)


def is_typing_key(vk, ctrl, alt):
    """字母、数字和符号键：Explorer 会把它们当作按名称定位的输入。"""
    if ctrl or alt:
        return False
    return 0x30 <= vk <= 0x5A or 0xBA <= vk <= 0xC0 or 0xDB <= vk <= 0xDF


class _GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.wintypes.DWORD), ('flags', ctypes.wintypes.DWORD),
        ('hwndActive', ctypes.wintypes.HWND), ('hwndFocus', ctypes.wintypes.HWND),
        ('hwndCapture', ctypes.wintypes.HWND), ('hwndMenuOwner', ctypes.wintypes.HWND),
        ('hwndMoveSize', ctypes.wintypes.HWND), ('hwndCaret', ctypes.wintypes.HWND),
        ('rcCaret', ctypes.wintypes.RECT),
    ]


def foreground_focus_class_name():
    """前台窗口当前获得键盘焦点的子控件类名；无法获取时返回空字符串。"""
    try:
        user32 = ctypes.windll.user32
        user32.GetForegroundWindow.restype = ctypes.wintypes.HWND
        foreground = user32.GetForegroundWindow()
        if not foreground:
            return ''
        thread_id = user32.GetWindowThreadProcessId(foreground, None)
        info = _GUITHREADINFO()
        info.cbSize = ctypes.sizeof(info)
        if not user32.GetGUIThreadInfo(thread_id, ctypes.byref(info)) or not info.hwndFocus:
            return ''
        name = ctypes.create_unicode_buffer(128)
        user32.GetClassNameW(info.hwndFocus, name, len(name))
        return name.value
    except Exception as error:
        debug_print(f"[QuickLook] focus class lookup failed: {error}")
        return ''


def preview(path, executable=None):
    """让 QuickLook 预览（或再次调用时关闭预览）指定路径；QuickLook 未运行时会先启动。"""
    global _owned_process, _last_preview_at
    executable = executable or find_quicklook()
    if not executable or not path:
        return False
    try:
        process = subprocess.Popen([executable, path], cwd=os.path.dirname(executable), close_fds=True)
    except OSError as error:
        debug_print(f"[QuickLook] launch failed: {error}")
        return False
    _last_preview_at = time.monotonic()
    # 第一个实例会常驻托盘；之后再启动的进程只是把路径转交给它并立即退出
    if _owned_process is None or _owned_process.poll() is not None:
        _owned_process = process
    return True


_WS_EX_APPWINDOW = 0x00040000
_GWL_EXSTYLE = -20
# SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED
_SWP_REFRESH_FRAME = 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020
_PREVIEW_CLASS_PREFIX = 'HwndWrapper[QuickLook.exe'


def is_preview_window(class_name, visible):
    """可见的 QuickLook WPF 顶层窗口即预览窗口（其余同类窗口都是隐藏的消息窗口）。"""
    return bool(visible) and str(class_name or '').startswith(_PREVIEW_CLASS_PREFIX)


class _GUID(ctypes.Structure):
    _fields_ = [('a', ctypes.c_ulong), ('b', ctypes.c_ushort), ('c', ctypes.c_ushort), ('d', ctypes.c_ubyte * 8)]


def _guid(text):
    value = _GUID()
    ctypes.windll.ole32.CLSIDFromString(text, ctypes.byref(value))
    return value


def _add_taskbar_tab(hwnd):
    """用 ITaskbarList::AddTab 让任务栏立即为 hwnd 建按钮；失败返回 False。

    任务栏只在窗口显示的瞬间判断是否建按钮，窗口已显示后再改样式不会被重新评估。"""
    ole32 = ctypes.windll.ole32
    taskbar = ctypes.c_void_p()
    if ole32.CoCreateInstance(ctypes.byref(_guid('{56FDF344-FD6D-11d0-958A-006097C9A090}')), None, 1,
                              ctypes.byref(_guid('{56FDF342-FD6D-11d0-958A-006097C9A090}')),
                              ctypes.byref(taskbar)) != 0 or not taskbar.value:
        return False
    # ITaskbarList 虚函数表：QueryInterface(0) AddRef(1) Release(2) HrInit(3) AddTab(4)
    methods = ctypes.cast(ctypes.cast(taskbar, ctypes.POINTER(ctypes.c_void_p))[0], ctypes.POINTER(ctypes.c_void_p))
    this = ctypes.c_void_p
    try:
        if ctypes.WINFUNCTYPE(ctypes.c_long, this)(methods[3])(taskbar) != 0:
            return False
        return ctypes.WINFUNCTYPE(ctypes.c_long, this, ctypes.wintypes.HWND)(methods[4])(taskbar, hwnd) == 0
    finally:
        ctypes.WINFUNCTYPE(ctypes.c_ulong, this)(methods[2])(taskbar)


def show_preview_in_taskbar():
    """让正在显示的 QuickLook 预览窗口出现在任务栏（QuickLook 默认给它挂了隐藏的 owner 窗口，不显示任务栏按钮）。

    只改窗口扩展样式并通知任务栏，可重复调用；返回 (可见的预览窗口数, 本次新加任务栏按钮的窗口数)。"""
    visible, updated = [], []
    try:
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        get_long = getattr(user32, 'GetWindowLongPtrW', None) or user32.GetWindowLongW
        set_long = getattr(user32, 'SetWindowLongPtrW', None) or user32.SetWindowLongW
        get_long.argtypes = [ctypes.wintypes.HWND, ctypes.c_int]
        get_long.restype = ctypes.c_ssize_t
        set_long.argtypes = [ctypes.wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
        set_long.restype = ctypes.c_ssize_t
        user32.SetWindowPos.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                        ctypes.c_int, ctypes.c_int, ctypes.c_uint]
        user32.IsWindowVisible.argtypes = [ctypes.wintypes.HWND]
        user32.ShowWindow.argtypes = [ctypes.wintypes.HWND, ctypes.c_int]
        user32.GetClassNameW.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.LPWSTR, ctypes.c_int]
        callback_type = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        name = ctypes.create_unicode_buffer(128)

        def visit(hwnd, _lparam):
            if user32.GetClassNameW(hwnd, name, len(name)) and is_preview_window(name.value, user32.IsWindowVisible(hwnd)):
                visible.append(hwnd)
                style = get_long(hwnd, _GWL_EXSTYLE)
                if not style & _WS_EX_APPWINDOW:
                    set_long(hwnd, _GWL_EXSTYLE, style | _WS_EX_APPWINDOW)
                    user32.SetWindowPos(hwnd, None, 0, 0, 0, 0, _SWP_REFRESH_FRAME)
                    if not _add_taskbar_tab(hwnd):
                        # 兜底：隐藏再显示会让任务栏重新评估该窗口
                        user32.ShowWindow(hwnd, 0)
                        user32.ShowWindow(hwnd, 8)
                    updated.append(hwnd)
            return True

        user32.EnumWindows.argtypes = [callback_type, ctypes.wintypes.LPARAM]
        user32.EnumWindows(callback_type(visit), 0)
    except Exception as error:
        debug_print(f"[QuickLook] taskbar button failed: {error}")
    return len(visible), len(updated)


class QuickLookTaskbarWatcher(QObject):
    """按空格预览后持续监视 QuickLook 预览窗口并补上任务栏按钮。

    QuickLook 要等内容加载完才显示窗口（大文件可能十几秒），且每次关闭后再预览都会新建窗口，
    所以不能只在固定时间点检查：开始监视后最长 WATCH_SECONDS 秒内等待窗口出现，窗口存在期间持续保持。"""

    WATCH_SECONDS = 60
    INTERVAL_MS = 200

    def __init__(self, parent=None):
        super().__init__(parent)
        self._deadline = 0.0
        self._timer = QTimer(self)
        self._timer.setInterval(self.INTERVAL_MS)
        self._timer.timeout.connect(self._tick)

    def watch(self):
        self._deadline = time.monotonic() + self.WATCH_SECONDS
        if not self._timer.isActive():
            self._timer.start()
        self._tick()

    def _tick(self):
        visible, _updated = show_preview_in_taskbar()
        if not visible and time.monotonic() >= self._deadline:
            self._timer.stop()


def owned_process_id():
    """本程序启动且仍在运行的 QuickLook 进程号；没有则返回 None。"""
    process = _owned_process
    if process is not None and process.poll() is None:
        return process.pid
    return None


def stop_owned_process():
    """退出本程序启动的 QuickLook，避免托盘里残留；不影响用户自己启动的实例。"""
    global _owned_process
    process, _owned_process = _owned_process, None
    if process is not None and process.poll() is None:
        try:
            process.terminate()
        except OSError:
            pass


def stop_owned_process_if_idle(timeout_s=QUICKLOOK_IDLE_TIMEOUT_S):
    """本程序启动的 QuickLook 空闲超过 timeout_s 秒则结束并释放其内存；返回是否已结束。"""
    if owned_process_id() is None or _last_preview_at is None:
        return False
    if time.monotonic() - _last_preview_at < timeout_s:
        return False
    debug_print("[QuickLook] stopping idle QuickLook to release memory")
    stop_owned_process()
    return True


def _file_sha256(path):
    digest = hashlib.sha256()
    with open(path, 'rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _download_archive(archive, url, size, progress, cancelled, opener):
    existing = os.path.getsize(archive) if os.path.isfile(archive) else 0
    if existing >= size:
        os.remove(archive)
        existing = 0
    headers = {'User-Agent': 'TabExplorer'}
    if existing:
        headers['Range'] = f'bytes={existing}-'
    with opener(urllib.request.Request(url, headers=headers), timeout=_NETWORK_TIMEOUT_S) as response:
        resumed = bool(existing) and getattr(response, 'status', 200) == 206
        done = existing if resumed else 0
        with open(archive, 'ab' if resumed else 'wb') as stream:
            while True:
                if cancelled():
                    raise QuickLookCancelled()
                chunk = response.read(_CHUNK_SIZE)
                if not chunk:
                    break
                stream.write(chunk)
                done += len(chunk)
                progress(done, size)


def install_quicklook(progress=None, cancelled=None, url=QUICKLOOK_URL, sha256=QUICKLOOK_SHA256,
                      size=QUICKLOOK_SIZE, opener=urllib.request.urlopen):
    """下载、校验并解压 QuickLook 便携版，返回 QuickLook.exe 路径。

    中断的下载保留 .part 文件，下次从断点继续；校验失败的文件会被删除。"""
    progress = progress or (lambda done, total: None)
    cancelled = cancelled or (lambda: False)
    seven_zip = get_bundled_path(os.path.join('tools', '7zr.exe'))
    if not seven_zip:
        raise QuickLookInstallError(tr("缺少解压组件 7zr.exe"))
    target = quicklook_dir()
    archive = target + '.7z.part'
    staging = target + '.extracting'
    os.makedirs(os.path.dirname(target), exist_ok=True)

    _download_archive(archive, url, size, progress, cancelled, opener)
    if _file_sha256(archive) != sha256.lower():
        os.remove(archive)
        raise QuickLookInstallError(tr("下载文件校验失败，已丢弃，请重试"))
    if cancelled():
        raise QuickLookCancelled()

    shutil.rmtree(staging, ignore_errors=True)
    try:
        result = subprocess.run(
            [seven_zip, 'x', archive, f'-o{staging}', '-y', '-bso0', '-bsp0'],
            capture_output=True, timeout=600, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
        )
        if result.returncode != 0 or not os.path.isfile(os.path.join(staging, 'QuickLook.exe')):
            raise QuickLookInstallError(tr("解压失败：{}").format(result.stderr.decode('utf-8', errors='replace').strip()))
        shutil.rmtree(target, ignore_errors=True)
        os.replace(staging, target)
    finally:
        shutil.rmtree(staging, ignore_errors=True)
    os.remove(archive)
    return os.path.join(target, 'QuickLook.exe')


class QuickLookInstallWorker(QThread):
    progressed = pyqtSignal(int, int)
    completed = pyqtSignal(bool, str)  # 是否成功, 失败原因（取消时为空）

    def run(self):
        try:
            install_quicklook(self.progressed.emit, self.isInterruptionRequested)
        except QuickLookCancelled:
            self.completed.emit(False, '')
        except Exception as error:
            debug_print(f"[QuickLook] install failed: {error}")
            self.completed.emit(False, str(error) or error.__class__.__name__)
        else:
            self.completed.emit(True, '')

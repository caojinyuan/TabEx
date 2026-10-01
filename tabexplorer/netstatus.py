"""网络位置断线提示：失败原因说明、重新连接（WNet）与标签内提示条。"""

import ctypes
import ctypes.wintypes
import os

from PyQt5.QtCore import pyqtSignal, Qt, QThread
from PyQt5.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QSizePolicy, QStyle, QToolButton, QVBoxLayout

from . import theme as _theme
from .i18n import tr
from .debuglog import debug_print
from .widgets import _set_tool_icon


ERROR_ACCESS_DENIED = 5
ERROR_BAD_NETPATH = 53
ERROR_ALREADY_ASSIGNED = 85
ERROR_CONNECTION_UNAVAIL = 1201
ERROR_DEVICE_ALREADY_REMEMBERED = 1202
ERROR_SESSION_CREDENTIAL_CONFLICT = 1219
ERROR_CANCELLED = 1223
ERROR_NOT_CONNECTED = 2250
NAV_TIMEOUT = -2  # 本程序的导航超时，不是系统错误码
# 重新连接返回这些值时连接已可用，直接重新打开位置
RECONNECTED_CODES = (0, ERROR_ALREADY_ASSIGNED, ERROR_DEVICE_ALREADY_REMEMBERED, ERROR_SESSION_CREDENTIAL_CONFLICT)

_UNREACHABLE_CODES = {51, 53, 59, 64, 121, 1203, 1222, 1231, 1232, 1236, 1311, ERROR_NOT_CONNECTED}
_AUTH_CODES = {ERROR_ACCESS_DENIED, 86, ERROR_SESSION_CREDENTIAL_CONFLICT, 1244, 1326, 1327, 1328, 1330,
               1331, 1907, 1909}
_NOT_FOUND_CODES = {2, 3, 67, 161}


def win32_error(hr):
    """FACILITY_WIN32 的 HRESULT 还原为 Win32 错误码，其余按无符号值返回。"""
    value = int(hr) & 0xFFFFFFFF
    return value & 0xFFFF if value & 0xFFFF0000 == 0x80070000 else value


def network_root(path):
    """UNC 路径返回 \\\\server\\share，盘符路径返回 'Z:'；无法识别时返回空字符串。"""
    text = str(path or '').replace('/', '\\')
    if text.startswith('\\\\'):
        parts = [part for part in text[2:].split('\\') if part]
        if len(parts) >= 2 and parts[0] not in ('?', '.'):
            return '\\\\' + parts[0] + '\\' + parts[1]
        return ''
    drive = os.path.splitdrive(text)[0]
    return drive.upper() if len(drive) == 2 and drive[1] == ':' else ''


def remembered_remote_name(drive):
    """映射盘（含已记住但当前未连接）对应的远程共享；不是网络驱动器时返回空字符串。只查本机记录。"""
    buffer = ctypes.create_unicode_buffer(1024)
    length = ctypes.wintypes.DWORD(len(buffer))
    try:
        mpr = ctypes.WinDLL('mpr')
        get_connection = mpr.WNetGetConnectionW
        get_connection.argtypes = [ctypes.wintypes.LPCWSTR, ctypes.wintypes.LPWSTR,
                                   ctypes.POINTER(ctypes.wintypes.DWORD)]
        get_connection.restype = ctypes.wintypes.DWORD
        code = get_connection(drive, buffer, ctypes.byref(length))
    except (AttributeError, OSError):
        return ''
    return buffer.value if code in (0, ERROR_CONNECTION_UNAVAIL) else ''


def is_network_path(path):
    root = network_root(path)
    if not root:
        return False
    if root.startswith('\\\\'):
        return True
    try:
        if ctypes.windll.kernel32.GetDriveTypeW(root + '\\') == 4:  # DRIVE_REMOTE
            return True
    except (AttributeError, OSError):
        return False
    return bool(remembered_remote_name(root))


def describe_failure(code, network=True):
    """返回 (原因, 建议)；原因优先使用系统本地化的错误说明。"""
    if code == NAV_TIMEOUT:
        return tr("连接超时：网络位置长时间无响应。"), tr("请检查网络或 VPN 连接，以及服务器是否在线。")
    if code:
        try:
            reason = ctypes.FormatError(code if code < 0x80000000 else code - 0x100000000).strip()
        except (OSError, OverflowError, ValueError):
            reason = ''
        if not reason or reason.startswith('<'):
            reason = tr("错误代码 {}").format(code if code <= 0xFFFF else f'0x{code:08X}')
    else:
        reason = tr("无法打开该位置。")
    if code in _UNREACHABLE_CODES:
        hint = tr("请检查网络或 VPN 连接，以及服务器是否在线。")
    elif code in _AUTH_CODES:
        hint = tr("可点击“重新连接”输入凭据。")
    elif code == ERROR_CONNECTION_UNAVAIL:
        hint = tr("映射的网络驱动器未连接，可点击“重新连接”恢复。")
    elif code in _NOT_FOUND_CODES:
        hint = tr("请确认路径是否正确。")
    else:
        hint = tr("网络连接可能已断开。") if network else ''
    return reason, hint


class _NETRESOURCEW(ctypes.Structure):
    _fields_ = [('dwScope', ctypes.wintypes.DWORD), ('dwType', ctypes.wintypes.DWORD),
                ('dwDisplayType', ctypes.wintypes.DWORD), ('dwUsage', ctypes.wintypes.DWORD),
                ('lpLocalName', ctypes.wintypes.LPWSTR), ('lpRemoteName', ctypes.wintypes.LPWSTR),
                ('lpComment', ctypes.wintypes.LPWSTR), ('lpProvider', ctypes.wintypes.LPWSTR)]


def reconnect_network_location(path, hwnd):
    """重新建立网络连接，需要时由系统弹出凭据框；返回 Win32 错误码（0 为成功）。会阻塞，须在后台线程调用。"""
    root = network_root(path)
    if not root:
        return ERROR_BAD_NETPATH
    mpr = ctypes.WinDLL('mpr')
    if root.startswith('\\\\'):
        resource = _NETRESOURCEW()
        resource.dwType = 1  # RESOURCETYPE_DISK
        resource.lpRemoteName = root
        add_connection = mpr.WNetAddConnection3W
        add_connection.argtypes = [ctypes.wintypes.HWND, ctypes.POINTER(_NETRESOURCEW), ctypes.wintypes.LPCWSTR,
                                   ctypes.wintypes.LPCWSTR, ctypes.wintypes.DWORD]
        add_connection.restype = ctypes.wintypes.DWORD
        return add_connection(hwnd, ctypes.byref(resource), None, None, 0x8)  # CONNECT_INTERACTIVE
    restore = mpr.WNetRestoreSingleConnectionW
    restore.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.LPCWSTR, ctypes.wintypes.BOOL]
    restore.restype = ctypes.wintypes.DWORD
    return restore(hwnd, root, True)


class ReconnectWorker(QThread):
    """后台重新连接网络位置；completed(path, win32_error)。"""
    completed = pyqtSignal(str, int)

    def __init__(self, path, hwnd, parent=None):
        super().__init__(parent)
        self.path = path
        self.hwnd = hwnd

    def run(self):
        code = ERROR_BAD_NETPATH
        co_init = False
        try:
            co_init = ctypes.windll.ole32.CoInitializeEx(None, 0x2) in (0, 1)  # COINIT_APARTMENTTHREADED
            code = reconnect_network_location(self.path, self.hwnd)
        except (AttributeError, OSError) as error:
            debug_print(f"[NetReconnect] {self.path}: {error}")
        finally:
            if co_init:
                ctypes.windll.ole32.CoUninitialize()
        debug_print(f"[NetReconnect] {self.path}: result={code}")
        self.completed.emit(self.path, int(code))


class _ElidedLabel(QLabel):
    """单行标签，宽度不足时中间省略，完整文字放在提示中。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._full_text = ''
        self.setTextFormat(Qt.PlainText)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)

    def set_full_text(self, text):
        self._full_text = text
        self.setToolTip(text)
        self._elide()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._elide()

    def _elide(self):
        self.setText(self.fontMetrics().elidedText(self._full_text, Qt.ElideMiddle, max(40, self.width())))


class NetworkErrorBanner(QFrame):
    """标签内的位置不可用提示条：说明原因，提供重试、重新连接与关闭。"""
    retryRequested = pyqtSignal(str)
    reconnectRequested = pyqtSignal(str)

    _CSS = (
        "QFrame#networkBanner { background: #fff4e5; border-bottom: 1px solid #f0b35a; }"
        "QLabel { background: transparent; color: #5c3b00; }"
        "QLabel#networkBannerTitle { font-weight: bold; }"
        "QPushButton { background: #ffffff; color: #3d2a00; border: 1px solid #d8b27a;"
        " border-radius: 3px; padding: 3px 12px; }"
        "QPushButton:hover { background: #fff0d6; }"
        "QPushButton:disabled { color: #a38f70; }"
        "QToolButton { background: transparent; border: none; border-radius: 3px; }"
        "QToolButton:hover { background: #f5dcb3; }"
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('networkBanner')
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        self.path = ''
        self.code = 0
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 6, 6, 6)
        layout.setSpacing(8)
        self.icon_label = QLabel(self)
        self.icon_label.setFixedSize(20, 20)
        layout.addWidget(self.icon_label, 0, Qt.AlignTop)
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        self.title_label = _ElidedLabel(self)
        self.title_label.setObjectName('networkBannerTitle')
        self.detail_label = QLabel(self)
        self.detail_label.setTextFormat(Qt.PlainText)
        self.detail_label.setWordWrap(True)
        text_layout.addWidget(self.title_label)
        text_layout.addWidget(self.detail_label)
        layout.addLayout(text_layout, 1)
        self.retry_button = QPushButton(tr("重试"), self)
        self.retry_button.setToolTip(tr("重新打开该位置"))
        self.retry_button.clicked.connect(lambda: self.retryRequested.emit(self.path))
        self.reconnect_button = QPushButton(tr("重新连接…"), self)
        self.reconnect_button.setToolTip(tr("重新建立网络连接，需要时由系统提示输入凭据"))
        self.reconnect_button.clicked.connect(lambda: self.reconnectRequested.emit(self.path))
        self.close_button = QToolButton(self)
        self.close_button.setToolTip(tr("关闭提示"))
        self.close_button.setFixedSize(24, 24)
        _set_tool_icon(self.close_button, 'window-close', QStyle.SP_TitleBarCloseButton, 14)
        self.close_button.clicked.connect(self.hide)
        layout.addWidget(self.retry_button, 0, Qt.AlignVCenter)
        layout.addWidget(self.reconnect_button, 0, Qt.AlignVCenter)
        layout.addWidget(self.close_button, 0, Qt.AlignTop)
        _theme.bind_style(self, self._CSS)
        self.hide()

    def show_failure(self, path, code, network=True, note=''):
        self.path = path
        self.code = code
        reason, hint = describe_failure(code, network)
        title = tr("无法访问网络位置：{}") if network else tr("无法访问此位置：{}")
        self.title_label.set_full_text(title.format(path))
        detail = ''
        for part in (note, reason, hint):
            if part:
                # 中文句末已有全角标点，不再补空格
                detail += part if not detail or detail.endswith(('。', '！', '？', '：')) else ' ' + part
        self.detail_label.setText(detail)
        self.reconnect_button.setVisible(network)
        self.retry_button.setEnabled(True)
        self.reconnect_button.setEnabled(True)
        self.icon_label.setPixmap(self.style().standardIcon(QStyle.SP_MessageBoxWarning).pixmap(20, 20))
        self.show()

    def set_busy(self, text):
        self.detail_label.setText(text)
        self.retry_button.setEnabled(False)
        self.reconnect_button.setEnabled(False)

    def is_busy(self):
        return self.isVisible() and not self.retry_button.isEnabled()

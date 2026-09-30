"""嵌入 Windows 资源管理器视图（IExplorerBrowser）及 Shell 路径工具。"""

import ctypes.wintypes
import os

from PyQt5.QtCore import (
    pyqtSignal, QAbstractNativeEventFilter, QDir, QObject, QRunnable, Qt, QThread, QThreadPool, QTimer,
)
from PyQt5.QtWidgets import QWidget

from .i18n import tr
from .debuglog import dbg_exc, debug_print
from .system import HAS_PYWIN, _retain_thread_until_finished
from .hotkeys import _hotkey_reserved_combos


class _PidlResolver(QThread):
    """后台解析 shell 绝对 PIDL。

    SHParseDisplayName 对网络/UNC/映射盘是同步阻塞调用，若在 UI 线程执行会冻结
    整个 Qt 事件循环——表现为“一个慢标签把所有标签/窗口都卡住”。本线程在后台完成
    解析，把解析出的绝对 PIDL（进程内有效、非 COM 接口指针，可跨线程传递并用
    CoTaskMemFree 释放）通过信号回传 UI 线程，由 UI 线程调用 BrowseToIDList 导航。
    """
    resolved = pyqtSignal(str, object, int, int)  # path, pidl(int|None), hr, generation

    def __init__(self, path, generation, parent=None):
        super().__init__(parent)
        self._path = path
        self._generation = generation

    def run(self):
        import ctypes
        pidl_val = None
        hr = -1
        co_init = False
        try:
            try:
                # 后台线程需自备 COM 环境；COINIT_APARTMENTTHREADED = 0x2
                ctypes.windll.ole32.CoInitializeEx(None, 0x2)
                co_init = True
            except Exception:
                pass
            _spdn = ctypes.windll.shell32.SHParseDisplayName
            _spdn.restype = ctypes.c_long
            _spdn.argtypes = [
                ctypes.c_wchar_p, ctypes.c_void_p,
                ctypes.POINTER(ctypes.c_void_p),
                ctypes.c_ulong, ctypes.POINTER(ctypes.c_ulong),
            ]
            pidl = ctypes.c_void_p(0)
            sfgao = ctypes.c_ulong(0)
            hr = _spdn(self._path, None, ctypes.byref(pidl), 0, ctypes.byref(sfgao))
            if hr == 0 and pidl.value:
                pidl_val = pidl.value  # 绝对 PIDL：跨线程有效
        except Exception as e:
            debug_print(f"[PidlResolver] error for '{self._path}': {e}")
            hr = -1
        finally:
            if co_init:
                try:
                    ctypes.windll.ole32.CoUninitialize()
                except Exception:
                    pass
        self.resolved.emit(self._path, pidl_val, int(hr), self._generation)


# ─────────────────────────────────────────────────────────────────────────────
# IExplorerBrowser-based file view
# Hosts the real Windows Explorer shell component (IExplorerBrowser COM) which
# supports TortoiseGit overlay icons – unlike Shell.Explorer (IE/WebBrowser
# ActiveX) which does not load shell icon-overlay extensions.
# ─────────────────────────────────────────────────────────────────────────────
_COMTYPES_AVAILABLE = False
try:
    import comtypes
    import comtypes.client
    from comtypes import GUID, HRESULT, IUnknown, COMMETHOD
    _COMTYPES_AVAILABLE = True
except ImportError:
    pass

if _COMTYPES_AVAILABLE:
    class _FOLDERSETTINGS(ctypes.Structure):
        _fields_ = [("ViewMode", ctypes.c_uint), ("fFlags", ctypes.c_uint)]

    class _IExplorerBrowserEvents(IUnknown):
        _iid_ = GUID("{361BBDC7-E6EE-4E13-BE58-58E2240C810F}")
        _methods_ = [
            COMMETHOD([], HRESULT, 'OnNavigationPending',
                      (['in'], ctypes.c_void_p, 'pidlFolder')),
            COMMETHOD([], HRESULT, 'OnViewCreated',
                      (['in'], ctypes.POINTER(IUnknown), 'psv')),
            COMMETHOD([], HRESULT, 'OnNavigationComplete',
                      (['in'], ctypes.c_void_p, 'pidlFolder')),
            COMMETHOD([], HRESULT, 'OnNavigationFailed',
                      (['in'], ctypes.c_void_p, 'pidlFolder')),
        ]

    class _IExplorerBrowser(IUnknown):
        _iid_ = GUID("{DFD3B6B5-C10C-4BE9-85F6-A66969F402F6}")
        _methods_ = [
            COMMETHOD([], HRESULT, 'Initialize',
                      (['in'], ctypes.wintypes.HWND, 'hwndParent'),
                      (['in'], ctypes.POINTER(ctypes.wintypes.RECT), 'prc'),
                      (['in'], ctypes.POINTER(_FOLDERSETTINGS), 'pfs')),
            COMMETHOD([], HRESULT, 'Destroy'),
            COMMETHOD([], HRESULT, 'SetRect',
                      (['in'], ctypes.c_void_p, 'phdwp'),
                      (['in'], ctypes.wintypes.RECT, 'rcBrowser')),
            COMMETHOD([], HRESULT, 'SetPropertyBag',
                      (['in'], ctypes.c_wchar_p, 'pszPropertyBag')),
            COMMETHOD([], HRESULT, 'SetEmptyText',
                      (['in'], ctypes.c_wchar_p, 'pszEmptyText')),
            COMMETHOD([], HRESULT, 'SetFolderSettings',
                      (['in'], ctypes.POINTER(_FOLDERSETTINGS), 'pfs')),
            COMMETHOD([], HRESULT, 'Advise',
                      (['in'], ctypes.POINTER(_IExplorerBrowserEvents), 'psbe'),
                      (['out'], ctypes.POINTER(ctypes.c_ulong), 'pdwCookie')),
            COMMETHOD([], HRESULT, 'Unadvise',
                      (['in'], ctypes.c_ulong, 'dwCookie')),
            COMMETHOD([], HRESULT, 'SetOptions',
                      (['in'], ctypes.c_uint, 'dwFlag')),
            COMMETHOD([], HRESULT, 'GetOptions',
                      (['out'], ctypes.POINTER(ctypes.c_uint), 'pdwFlag')),
            COMMETHOD([], HRESULT, 'BrowseToIDList',
                      (['in'], ctypes.c_void_p, 'pidl'),
                      (['in'], ctypes.c_uint, 'uFlags')),
            COMMETHOD([], HRESULT, 'BrowseToObject',
                      (['in'], ctypes.POINTER(IUnknown), 'punk'),
                      (['in'], ctypes.c_uint, 'uFlags')),
            COMMETHOD([], HRESULT, 'FillFromObject',
                      (['in'], ctypes.POINTER(IUnknown), 'punk'),
                      (['in'], ctypes.c_uint, 'dwFlags')),
            COMMETHOD([], HRESULT, 'RemoveAll'),
            COMMETHOD([], HRESULT, 'GetCurrentView',
                      (['in'], ctypes.POINTER(GUID), 'riid'),
                      (['out'], ctypes.POINTER(ctypes.c_void_p), 'ppv')),
        ]

    _CLSID_ExplorerBrowser = GUID("{71F96385-DDD6-48D3-A0C1-AE06E8B055FB}")

    class _NavEventSink(comtypes.COMObject):
        """IExplorerBrowserEvents sink – receives navigation-complete callbacks."""
        _com_interfaces_ = [_IExplorerBrowserEvents]

        def __init__(self, on_complete):
            super().__init__()
            self._on_complete = on_complete

        def OnNavigationPending(self, pidlFolder):
            return 0  # S_OK

        def OnViewCreated(self, psv):
            return 0  # S_OK

        def OnNavigationComplete(self, pidlFolder):
            # DIAGNOSTIC: this is the raw COM navigation event. If this line
            # stops printing after a minimize/restore while in-shell navigation
            # still visibly happens, the COM event sink has gone deaf (which
            # would leave current_path / _location_url permanently stale).
            debug_print("[IEB NavSink] OnNavigationComplete fired (COM event alive)")
            try:
                if pidlFolder and self._on_complete:
                    _fn = ctypes.windll.shell32.SHGetPathFromIDListW
                    _fn.restype  = ctypes.c_bool
                    _fn.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p]
                    buf = ctypes.create_unicode_buffer(32768)
                    if _fn(pidlFolder, buf):
                        path = buf.value
                        if path:
                            debug_print(f"[IEB NavSink] resolved path: {path}")
                            self._on_complete(path)
            except Exception:
                pass
            return 0  # S_OK

        def OnNavigationFailed(self, pidlFolder):
            return 0  # S_OK


# ── TortoiseGit overlay fix ───────────────────────────────────────────────────
# TortoiseOverlays.dll (registered in HKLM ShellIconOverlayIdentifiers) provides
# GetOverlayInfo correctly (icon paths) but its IsMemberOf returns S_FALSE in
# non-Explorer processes (process name check).
#
# Fix: Patch TortoiseOverlays' IsMemberOf vtable[3] in our process to delegate
# to TortoiseGitStub's IsMemberOf (which works in any process).
# TortoiseOverlays type stored at this+0x08 (0=Normal..8=Unversioned).
# TortoiseGitStub type stored at this+0x28 (1=Normal..9=Unversioned).
# ─────────────────────────────────────────────────────────────────────────────

# TortoiseOverlays CLSIDs (from HKLM, shared vtable in TortoiseOverlays64.dll)
_TORTOISE_OVERLAYS_CLSIDS = [
    '{C5994560-53D9-4125-87C9-F193FC689CB2}',  # Normal (type=0)
    '{C5994561-53D9-4125-87C9-F193FC689CB2}',  # Modified (type=1)
    '{C5994562-53D9-4125-87C9-F193FC689CB2}',  # Conflict (type=2)
    '{C5994563-53D9-4125-87C9-F193FC689CB2}',  # Locked (type=3)
    '{C5994564-53D9-4125-87C9-F193FC689CB2}',  # ReadOnly (type=4)
    '{C5994565-53D9-4125-87C9-F193FC689CB2}',  # Deleted (type=5)
    '{C5994566-53D9-4125-87C9-F193FC689CB2}',  # Added (type=6)
    '{C5994567-53D9-4125-87C9-F193FC689CB2}',  # Ignored (type=7)
    '{C5994568-53D9-4125-87C9-F193FC689CB2}',  # Unversioned (type=8)
]

# TortoiseGitStub CLSIDs (IsMemberOf works in any process)
_TORTOISE_OVERLAY_PATCHED = False


def _patch_tortoise_overlays():
    """Pre-load TortoiseGit overlay icons into the process system image list.

    In non-explorer.exe processes, TortoiseOverlays' IsMemberOf returns S_FALSE,
    but the shell's per-user overlay CACHE (populated by explorer.exe) still
    provides correct overlay indices via IShellIconOverlay::GetOverlayIndex.

    The issue is that the overlay ICONS are not loaded into our process's
    system image list until something triggers them. We call SHGetFileInfo
    with SHGFI_OVERLAYINDEX which forces the shell to lazily load the
    overlay handler's icon (via GetOverlayInfo, which works in any process).

    Once loaded, the shell view (DefView/ItemsView) can render overlays.
    """
    global _TORTOISE_OVERLAY_PATCHED
    if _TORTOISE_OVERLAY_PATCHED:
        return

    try:
        import ctypes
        import ctypes.wintypes
        import winreg

        # Check TortoiseOverlays is installed
        try:
            winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                           rf'SOFTWARE\Classes\CLSID\{_TORTOISE_OVERLAYS_CLSIDS[0]}',
                           0, winreg.KEY_READ).Close()
        except OSError:
            return  # TortoiseOverlays not installed

        _TORTOISE_OVERLAY_PATCHED = True
        print("[TortoiseGit] Overlay icon pre-load enabled "
              "(will trigger on first navigation)")

    except Exception as e:
        print(f"[TortoiseGit] Overlay init error: {e}")


# Cache of directories already preloaded (overlay icons persist in system image list)
_OVERLAY_PRELOADED_DIRS = set()


def _file_url_to_path(url):
    """file: URL 转 Windows 路径，兼容 file:///C:/x、file://server/share、file:\\\\server\\share 等写法。"""
    from urllib.parse import unquote
    stripped = unquote(str(url)[5:].lstrip('/\\')).replace('/', '\\')
    if not stripped:
        return ''
    if len(stripped) >= 2 and stripped[1] == ':':
        return stripped
    return '\\\\' + stripped


_EXPLORER_CLSID_PATHS = {
    '{20D04FE0-3AEA-1069-A2D8-08002B30309D}': 'shell:MyComputerFolder',  # 此电脑
    '{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}': 'shell:NetworkPlacesFolder',  # 网络
    '{031E4825-7B94-4DC3-B131-E946B44C8DD5}': 'shell:Libraries',  # 库
}


def _explorer_location_to_path(location, location_name=None):
    """把 Explorer 窗口的 LocationURL / LocationName 转成可嵌入的路径；无法识别时返回主目录。"""
    if location_name and location_name in (tr('控制面板'), 'Control Panel'):
        return 'shell:ControlPanelFolder'
    if location:
        if location.lower().startswith('file:'):
            return _file_url_to_path(location) or QDir.homePath()
        if '::' in location:
            upper = location.upper()
            for clsid, shell_path in _EXPLORER_CLSID_PATHS.items():
                if clsid in upper:
                    return shell_path
            debug_print("[Explorer Monitor] Unknown CLSID, using default home path")
            return QDir.homePath()
        return location
    if location_name in (tr('此电脑'), 'This PC', 'My Computer'):
        return 'shell:MyComputerFolder'
    if location_name in (tr('网络'), 'Network'):
        return 'shell:NetworkPlacesFolder'
    if location_name in (tr('回收站'), 'Recycle Bin'):
        return 'shell:RecycleBinFolder'
    if location_name in (tr('用户帐户'), 'User Accounts', tr('程序和功能'), 'Programs and Features',
                         tr('系统'), 'System', tr('设备管理器'), 'Device Manager',
                         tr('网络和共享中心'), 'Network and Sharing Center'):
        return 'shell:ControlPanelFolder'
    return QDir.homePath()


def _path_is_slow_for_shell(path):
    """模块级慢盘判断：UNC / OneDrive / 映射网络盘上同步 Shell 调用可能阻塞 UI 线程。
    供 _preload_overlay_icons 早退使用（与 FileExplorerTab._is_slow_path 逻辑一致）。"""
    if not path:
        return False
    if path.startswith('\\\\') or path.startswith('//'):
        return True
    if 'onedrive' in path.replace('\\', '/').lower():
        return True
    try:
        import ctypes
        drive = os.path.splitdrive(path)[0]
        if drive and ctypes.windll.kernel32.GetDriveTypeW(drive + '\\') == 4:  # DRIVE_REMOTE
            return True
    except Exception:
        pass
    return False


def _preload_overlay_icons(directory_path):
    """Force-load overlay icons into the system image list by calling
    SHGetFileInfo(SHGFI_OVERLAYINDEX) on files in the given directory.

    This triggers the shell to lazily load overlay icons from handlers
    whose GetOverlayInfo works (TortoiseOverlays provides correct icons).
    The overlay INDEX comes from the cross-process shell cache (populated
    by explorer.exe), so TortoiseOverlays' IsMemberOf doesn't need to work.
    """
    # 慢盘（网络/OneDrive）上 SHGetFileInfo+scandir 是同步阻塞调用，导航前预加载会卡 UI；
    # 跳过即可，overlay 仍由导航完成后 600ms 的 IShellView::Refresh() 兜底渲染。
    if _path_is_slow_for_shell(directory_path):
        return
    try:
        import ctypes
        import ctypes.wintypes

        class _SHFILEINFOW(ctypes.Structure):
            _fields_ = [
                ('hIcon', ctypes.wintypes.HICON),
                ('iIcon', ctypes.c_int),
                ('dwAttributes', ctypes.wintypes.DWORD),
                ('szDisplayName', ctypes.c_wchar * 260),
                ('szTypeName', ctypes.c_wchar * 80),
            ]

        _SHGetFileInfoW = ctypes.windll.shell32.SHGetFileInfoW
        _SHGetFileInfoW.argtypes = [ctypes.c_wchar_p, ctypes.wintypes.DWORD,
                                    ctypes.POINTER(_SHFILEINFOW), ctypes.c_uint,
                                    ctypes.c_uint]
        _SHGetFileInfoW.restype = ctypes.c_void_p

        _DestroyIcon = ctypes.windll.user32.DestroyIcon

        # SHGFI_ICON=0x100, SHGFI_SMALLICON=0x1, SHGFI_OVERLAYINDEX=0x40
        FLAGS = 0x100 | 0x1 | 0x40

        loaded_overlays = set()

        # Query the directory itself first
        sfi = _SHFILEINFOW()
        _SHGetFileInfoW(directory_path, 0, ctypes.byref(sfi),
                        ctypes.sizeof(sfi), FLAGS)
        if sfi.hIcon:
            _DestroyIcon(sfi.hIcon)
        ovl = (sfi.iIcon >> 24) & 0xFF
        if ovl:
            loaded_overlays.add(ovl)

        # Query a few files to trigger overlay icon loading.
        # Once a TortoiseGit overlay (index > 4) is found, we can stop -
        # the overlay slot is process-global and applies to all files.
        MAX_FILES = 3
        count = 0
        found_git_overlay = any(o > 4 for o in loaded_overlays)
        try:
            for entry in os.scandir(directory_path):
                if found_git_overlay or count >= MAX_FILES:
                    break
                # Skip .exe/.msi/.dll - SHGetFileInfo can be very slow for these
                if entry.name.lower().endswith(('.exe', '.msi', '.dll')):
                    continue
                sfi2 = _SHFILEINFOW()
                _SHGetFileInfoW(entry.path, 0, ctypes.byref(sfi2),
                                ctypes.sizeof(sfi2), FLAGS)
                if sfi2.hIcon:
                    _DestroyIcon(sfi2.hIcon)
                ovl2 = (sfi2.iIcon >> 24) & 0xFF
                if ovl2:
                    loaded_overlays.add(ovl2)
                    if ovl2 > 4:
                        found_git_overlay = True
                count += 1
        except OSError:
            pass

        if loaded_overlays:
            debug_print(f"[TortoiseGit] Pre-loaded overlay icons: "
                        f"{sorted(loaded_overlays)} ({count} files scanned)")
            # Mark directory as preloaded (overlay icons are process-global)
            _OVERLAY_PRELOADED_DIRS.add(directory_path)

    except Exception:
        pass


class _OverlayPreloadSignals(QObject):
    """后台 overlay 图标预加载完成信号：done(path)。"""
    done = pyqtSignal(str)


class _OverlayPreloadRunnable(QRunnable):
    """在 QThreadPool 后台线程执行 _preload_overlay_icons（含 SHGetFileInfo 扫描），
    避免在 UI 线程做同步 Shell 调用造成切标签卡顿。完成后经信号回 UI 线程调用
    IShellView::Refresh()（COM 必须在 UI/STA 线程）。工作线程自行 CoInitialize。"""
    def __init__(self, path, signals):
        super().__init__()
        self._path = path
        self._signals = signals

    def run(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            _preload_overlay_icons(self._path)
        except Exception as e:
            debug_print(f"[TortoiseGit] async overlay preload error: {e}")
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            try:
                self._signals.done.emit(self._path)
            except RuntimeError:
                pass  # 信号对象已随标签销毁


# ── Early overlay initialization ─────────────────────────────────────────────
# Must run BEFORE IExplorerBrowser creates its shell view (which triggers SIOM).
if _COMTYPES_AVAILABLE and HAS_PYWIN:
    try:
        _patch_tortoise_overlays()
    except Exception as _e:
        print(f"[TortoiseGit] Early init error: {_e}")


# ── IExplorerBrowser 键盘消息过滤器 ──────────────────────────────────────────
# Qt 的事件循环会把所有 WM_KEYDOWN/WM_KEYUP/WM_CHAR 消息转换成 QKeyEvent。
# 但 IExplorerBrowser 的子窗口（SHELLDLL_DefView / SysListView32）是纯 Win32 控件，
# 它们依赖直接收到 WM_KEYDOWN 来实现 Delete/Ctrl+C/V/X/F2 等功能。
# QAbstractNativeEventFilter 安装在 QApplication 层，能拦截线程中所有 HWND 的消息。
# 我们在此检测键盘消息是否发往 IExplorerBrowser 子窗口，如果是则手动 Dispatch，
# 返回 True 告诉 Qt 跳过后续处理。

class _IEBKeyboardFilter(QAbstractNativeEventFilter):
    """Application-level native event filter that forwards keyboard messages
    to IExplorerBrowser child windows so that Delete, Ctrl+C/V/X, F2, etc. work."""

    # TabEx 自有快捷键 (ctrl, alt, vk)；随设置更新
    _tabex_hotkeys = _hotkey_reserved_combos({})

    def __init__(self):
        super().__init__()
        self._main_window = None
        self._ieb_hwnds = set()  # 缓存 IExplorerBrowserWidget 的顶层 HWND
        self._ieb_widgets = {}   # HWND → IExplorerBrowserWidget 实例映射
        self._debug_first_key = True  # 首次键盘消息诊断
        self._debug_first_call = True  # 首次调用诊断
        # IShellView COM IID
        self._IID_IShellView = None  # 延迟初始化（需要 comtypes）

    def set_main_window(self, mw):
        self._main_window = mw

    def set_tabex_hotkeys(self, combos):
        self._tabex_hotkeys = frozenset(combos)

    def register_ieb_hwnd(self, hwnd, widget=None):
        """注册 IExplorerBrowserWidget 的 winId，用于快速判断子窗口归属"""
        if hwnd:
            self._ieb_hwnds.add(hwnd)
            if widget is not None:
                self._ieb_widgets[hwnd] = widget
            debug_print(f"[IEB KeyFilter] Registered HWND: 0x{hwnd:X}, total: {len(self._ieb_hwnds)}")

    def unregister_ieb_hwnd(self, hwnd):
        self._ieb_hwnds.discard(hwnd)
        self._ieb_widgets.pop(hwnd, None)

    def _find_owner_widget(self, hwnd):
        """从目标 HWND 向上遍历找到拥有它的 IExplorerBrowserWidget"""
        import ctypes
        _GetParent = ctypes.windll.user32.GetParent
        _GetParent.restype = ctypes.wintypes.HWND
        h = hwnd
        for _ in range(30):
            if h in self._ieb_widgets:
                return self._ieb_widgets[h]
            h = _GetParent(h)
            if not h:
                break
        return None

    def _is_ieb_descendant(self, hwnd):
        """判断 hwnd 是否是已注册的某个 IExplorerBrowserWidget 的后代窗口"""
        import ctypes
        _GetParent = ctypes.windll.user32.GetParent
        _GetParent.restype = ctypes.wintypes.HWND  # 确保 64 位 HWND 不被截断
        h = hwnd
        for _ in range(30):
            if h in self._ieb_hwnds:
                return True
            h = _GetParent(h)
            if not h:
                break
        return False

    # 自定义 MSG 结构体，避免依赖 wintypes.MSG 的字段名
    class _MSG(ctypes.Structure):
        _fields_ = [
            ("hwnd", ctypes.c_void_p),
            ("message", ctypes.c_uint),
            ("wParam", ctypes.c_size_t),   # WPARAM = UINT_PTR
            ("lParam", ctypes.c_ssize_t),  # LPARAM = LONG_PTR
            ("time", ctypes.c_uint),
            ("pt_x", ctypes.c_long),
            ("pt_y", ctypes.c_long),
        ]

    def nativeEventFilter(self, eventType, message):
        try:
            # 首次调用诊断 — 确认过滤器被 Qt 调用
            if self._debug_first_call:
                self._debug_first_call = False
                debug_print(f"[IEB KeyFilter] Filter active! eventType={eventType} "
                            f"ieb_hwnds={[hex(h) for h in self._ieb_hwnds]}")

            if eventType not in (b"windows_generic_MSG", b"windows_dispatcher_MSG"):
                return False, 0
            if not self._ieb_hwnds:
                return False, 0

            import ctypes
            from ctypes import cast, POINTER
            msg = cast(int(message), POINTER(self._MSG)).contents

            # 只处理键盘类消息
            if msg.message not in (0x0100, 0x0101, 0x0102, 0x0104, 0x0105, 0x0106):
                return False, 0

            # 首次键盘消息诊断
            if self._debug_first_key and msg.message == 0x0100:
                self._debug_first_key = False
                debug_print(f"[IEB KeyFilter] First WM_KEYDOWN: VK=0x{msg.wParam & 0xFF:02X} "
                            f"target_hwnd=0x{msg.hwnd or 0:X} "
                            f"registered={[hex(h) for h in self._ieb_hwnds]} "
                            f"is_descendant={self._is_ieb_descendant(msg.hwnd) if msg.hwnd else False}")

            # msg.hwnd 是消息的目标窗口；检查它是否属于 IExplorerBrowser
            target_hwnd = msg.hwnd
            if not target_hwnd or not self._is_ieb_descendant(target_hwnd):
                return False, 0

            # 对 WM_KEYDOWN / WM_SYSKEYDOWN，排除 TabEx 自有快捷键
            if msg.message in (0x0100, 0x0104):
                vk = msg.wParam & 0xFF
                ctrl = (ctypes.windll.user32.GetKeyState(0x11) & 0x8000) != 0
                alt  = (ctypes.windll.user32.GetKeyState(0x12) & 0x8000) != 0
                if (ctrl, alt, vk) in self._tabex_hotkeys:
                    return False, 0  # 留给 TabEx 快捷键

            # 通过 IShellView::TranslateAccelerator 转发键盘消息
            # 这是 Shell 控件处理 Ctrl+C/V/X, Delete, F2 等的正确 COM 方式
            widget = self._find_owner_widget(target_hwnd)
            if widget and getattr(widget, '_browser', None):
                try:
                    if self._IID_IShellView is None:
                        from comtypes import GUID as _GUID
                        self._IID_IShellView = _GUID("{000214E3-0000-0000-C000-000000000046}")

                    # comtypes 的 ['out'] 参数自动返回，不需要手动传
                    ppv_result = widget._browser.GetCurrentView(
                        ctypes.byref(self._IID_IShellView))
                    # ppv_result 是 c_void_p 或整数
                    iface_ptr = int(ppv_result) if ppv_result else 0
                    if iface_ptr:
                        try:
                            # COM 对象布局: 对象地址 → vtable 指针 → 函数指针数组
                            # IShellView vtable: [QI(0), AddRef(1), Release(2),
                            #   GetWindow(3), ContextSensitiveHelp(4), TranslateAccelerator(5)]
                            _vp_size = ctypes.sizeof(ctypes.c_void_p)
                            vtable_ptr = ctypes.c_void_p.from_address(iface_ptr).value
                            fn_addr = ctypes.c_void_p.from_address(vtable_ptr + 5 * _vp_size).value
                            # TranslateAccelerator(IShellView* this, MSG* pmsg) → HRESULT
                            _TA = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_void_p)
                            translate_accel = _TA(fn_addr)
                            # 构造 wintypes.MSG（保证内存布局与 Windows 一致）
                            native_msg = ctypes.wintypes.MSG()
                            native_msg.hWnd = target_hwnd
                            native_msg.message = msg.message
                            native_msg.wParam = msg.wParam
                            native_msg.lParam = msg.lParam
                            native_msg.time = msg.time
                            native_msg.pt.x = msg.pt_x
                            native_msg.pt.y = msg.pt_y
                            ta_hr = translate_accel(iface_ptr, ctypes.byref(native_msg))
                            if ta_hr == 0:  # S_OK — Shell 已处理
                                return True, 0
                        finally:
                            # Release IShellView
                            release_addr = ctypes.c_void_p.from_address(
                                vtable_ptr + 2 * _vp_size).value
                            _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                            _REL(release_addr)(iface_ptr)
                except Exception as e:
                    debug_print(f"[IEB KeyFilter] TranslateAccelerator failed: {e}")

            # 回退：直接 Dispatch（处理 TranslateAccelerator 不支持的按键）
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))
            return True, 0

        except Exception as e:
            debug_print(f"[IEB KeyFilter] Exception: {e}")
            return False, 0


# 全局单例（在 main() 中安装）
_ieb_keyboard_filter = _IEBKeyboardFilter()


class IExplorerBrowserWidget(QWidget):
    """
    Qt widget that hosts IExplorerBrowser – the real Windows Explorer shell
    component. Supports TortoiseGit overlay icons, unlike the legacy
    Shell.Explorer (IE/WebBrowser ActiveX) control.

    Provides a drop-in subset of the QAxWidget API used by FileExplorerTab:
      dynamicCall('Navigate[2](...)', url)  – navigate to a URL or path
      dynamicCall('Refresh()')              – refresh current view
      property('LocationURL')               – current location as file:/// URL
      NavigateComplete2 signal              – emitted on navigation complete
      querySubObject(name)                  – returns None (graceful degradation)
      clear()                               – destroy COM resources
    """
    # (pDisp, url) – matches QAxWidget's NavigateComplete2 signature
    NavigateComplete2 = pyqtSignal(object, object)

    # 异步导航（慢盘）状态信号：宿主标签据此显示 loading / 解除导航锁
    navigationStarted  = pyqtSignal(str)        # path — 后台 PIDL 解析已开始
    navigationFinished = pyqtSignal(str, bool)  # path, ok — 解析失败/无法访问时触发

    # IExplorerBrowser option flags  (SDK shobjidl_core.h values)
    _EBO_SHOWFRAMES  = 0x00000002  # 显示左侧导航窗格
    _EBO_NOTRAVELLOG = 0x00000008  # 禁用前进/后退历史（避免与 TabEx 自身历史冲突）
    _EBO_NOBORDER    = 0x00000040  # 隐藏导航工具栏（地址栏）

    def __init__(self, parent=None):
        super().__init__(parent)
        self._browser      = None
        self._cookie       = ctypes.c_ulong(0)
        self._nav_sink     = None
        self._location_url = ''
        self._pending_path = None
        self._init_ok      = False
        # 异步导航（慢盘）：导航代号用于作废过期的后台解析结果；线程引用防 GC
        self._nav_generation = 0
        self._nav_resolver   = None
        # 后台 overlay 图标预加载：懒创建信号对象 + in-flight 去重标志
        self._overlay_signals = None
        self._overlay_preload_inflight = False
        # Force native HWND creation; IExplorerBrowser::Initialize needs a real HWND
        self.setAttribute(Qt.WA_NativeWindow, True)

    # ── QAxWidget compatibility ───────────────────────────────────────────────

    def property(self, name):  # noqa: A003
        """Override QObject.property to expose LocationURL."""
        if name == 'LocationURL':
            return self._location_url
        return super().property(name)

    def dynamicCall(self, sig, *args):
        """QAxWidget-compatible dispatch for Navigate / Refresh calls."""
        s = sig.lower()
        # Match only actual Navigate/Navigate2 calls, skip event signatures
        # (NavigateComplete2, BeforeNavigate2, etc.)
        if (s.startswith('navigate') or s.startswith('navigate2')) and \
           'complete' not in s and 'before' not in s:
            url = str(args[0]) if args else ''
            if url and url != 'None':
                self._navigate_url(url)
        elif 'refresh()' in s:
            self._do_refresh()
        # All other calls (Visible, ToolBar, Silent, NavigateComplete2 events…) are silently ignored

    def querySubObject(self, _name, *_args):
        """Returns None; selection info is not yet available via IExplorerBrowser."""
        return None

    def clear(self):
        """QAxWidget compat: release COM resources."""
        self.cleanup()

    # ── Navigation helpers ────────────────────────────────────────────────────

    def _navigate_url(self, url):
        path = self._url_to_path(url)
        if path:
            self._navigate(path)

    @staticmethod
    def _url_to_path(url):
        if not url:
            return ''
        url = str(url)
        if url.lower().startswith('file:'):
            return _file_url_to_path(url)
        return url  # local path or shell: path – pass through

    def _navigate(self, path):
        if not _COMTYPES_AVAILABLE:
            return
        if not self._ensure_browser():
            self._pending_path = path
            return
        self._pending_path = None
        # 慢盘（网络/UNC/OneDrive/映射远程盘）：把阻塞的 PIDL 解析放到后台线程，避免冻结
        # UI 事件循环——一个慢标签不再拖垮其它标签/窗口。本地快盘保持同步（无线程开销）。
        if _path_is_slow_for_shell(path):
            self._navigate_async(path)
        else:
            self._navigate_sync(path)

    def _navigate_sync(self, path):
        """同步导航（本地快盘）：UI 线程内解析 PIDL 并浏览。"""
        try:
            # Pre-load overlay icons BEFORE navigation so DefView has them
            # when it first renders (skip if already cached for this dir)
            if _TORTOISE_OVERLAY_PATCHED and path not in _OVERLAY_PRELOADED_DIRS:
                _preload_overlay_icons(path)

            _spdn = ctypes.windll.shell32.SHParseDisplayName
            _spdn.restype  = ctypes.c_long   # HRESULT
            _spdn.argtypes = [
                ctypes.c_wchar_p,
                ctypes.c_void_p,
                ctypes.POINTER(ctypes.c_void_p),
                ctypes.c_ulong,
                ctypes.POINTER(ctypes.c_ulong),
            ]
            pidl  = ctypes.c_void_p(0)
            sfgao = ctypes.c_ulong(0)
            hr    = _spdn(path, None, ctypes.byref(pidl), 0, ctypes.byref(sfgao))
            if hr == 0 and pidl.value:
                try:
                    self._browser.BrowseToIDList(pidl, 0)  # 0 = SBSP_ABSOLUTE
                finally:
                    ctypes.windll.ole32.CoTaskMemFree(pidl)
            else:
                debug_print(f"[IExplorerBrowser] SHParseDisplayName failed "
                            f"hr=0x{hr & 0xFFFFFFFF:08x} for '{path}'")
        except Exception as e:
            debug_print(f"[IExplorerBrowser] _navigate error: {e}")

    def _navigate_async(self, path):
        """异步导航（慢盘）：后台线程解析 PIDL，解析完再由 UI 线程浏览。"""
        # 递增导航代号：解析期间若再次发起导航，旧线程的结果会被作废，避免错误落地
        self._nav_generation += 1
        gen = self._nav_generation
        try:
            self.navigationStarted.emit(path)
        except Exception:
            pass
        resolver = _PidlResolver(path, gen, self)
        resolver.resolved.connect(self._on_pidl_resolved)
        self._nav_resolver = resolver
        resolver.start()
        _retain_thread_until_finished(resolver)

    def _on_pidl_resolved(self, path, pidl_val, hr, generation):
        """后台 PIDL 解析完成（UI 线程）：作废过期结果，否则浏览到该 PIDL。"""
        # 期间又发起了新导航 → 丢弃过期结果并释放其 PIDL
        if generation != self._nav_generation:
            if pidl_val:
                try:
                    ctypes.windll.ole32.CoTaskMemFree(ctypes.c_void_p(pidl_val))
                except Exception:
                    pass
            return
        try:
            if pidl_val and self._browser is not None:
                pidl = ctypes.c_void_p(pidl_val)
                try:
                    self._browser.BrowseToIDList(pidl, 0)  # 0 = SBSP_ABSOLUTE
                finally:
                    ctypes.windll.ole32.CoTaskMemFree(pidl)
                # 浏览已发起；内容加载完成由 NavigateComplete2 通知宿主标签隐藏 loading
            else:
                debug_print(f"[IExplorerBrowser] async SHParseDisplayName failed "
                            f"hr=0x{hr & 0xFFFFFFFF:08x} for '{path}'")
                # 解析失败（网络不可达/路径无效）：NavigateComplete2 不会触发，主动通知结束
                if pidl_val:
                    try:
                        ctypes.windll.ole32.CoTaskMemFree(ctypes.c_void_p(pidl_val))
                    except Exception:
                        pass
                try:
                    self.navigationFinished.emit(path, False)
                except Exception:
                    pass
        except Exception as e:
            debug_print(f"[IExplorerBrowser] async _navigate error: {e}")
            try:
                self.navigationFinished.emit(path, False)
            except Exception:
                pass

    def _do_refresh(self):
        if self._location_url:
            self._navigate_url(self._location_url)

    # ── IExplorerBrowser lifecycle ────────────────────────────────────────────

    def _ensure_browser(self):
        """Create/initialize IExplorerBrowser on first use."""
        if self._init_ok:
            return True
        if not _COMTYPES_AVAILABLE:
            return False
        try:
            hwnd = int(self.winId())
            if not hwnd:
                return False
            # TortoiseGit overlay icons are pre-loaded on navigation via
            # _preload_overlay_icons(). No patching needed here.
            browser = comtypes.client.CreateObject(
                _CLSID_ExplorerBrowser,
                interface=_IExplorerBrowser,
            )
            w = max(self.width(),  100)
            h = max(self.height(), 100)
            rc = ctypes.wintypes.RECT(0, 0, w, h)
            fs = _FOLDERSETTINGS(4, 0)  # ViewMode=FVM_DETAILS, fFlags=0
            browser.Initialize(hwnd, ctypes.byref(rc), ctypes.byref(fs))
            # EBO_SHOWFRAMES: 显示左侧导航窗格（目录树）
            browser.SetOptions(self._EBO_SHOWFRAMES | self._EBO_NOTRAVELLOG)
            sink   = _NavEventSink(self._on_nav_complete)
            cookie = browser.Advise(sink)  # out-param returned by comtypes
            self._browser  = browser
            self._nav_sink = sink
            self._cookie   = ctypes.c_ulong(cookie)
            self._init_ok  = True
            _ieb_keyboard_filter.register_ieb_hwnd(hwnd, widget=self)
            debug_print("[IExplorerBrowser] Initialized successfully")
            return True
        except Exception as e:
            debug_print(f"[IExplorerBrowser] Init failed: {e}")
            return False

    def _on_nav_complete(self, path):
        """Called by _NavEventSink when navigation completes."""
        norm = os.path.normpath(path)
        url  = 'file:///' + norm.replace('\\', '/')
        self._location_url = url
        self.NavigateComplete2.emit(None, url)
        # Force shell to re-evaluate icon overlays for this directory
        # Skip if this navigation was triggered by our own Refresh()
        if not getattr(self, '_overlay_refreshing', False):
            self._notify_overlay_refresh(norm)

    def _notify_overlay_refresh(self, path):
        """Refresh overlay rendering after navigation if needed."""
        if not _TORTOISE_OVERLAY_PATCHED:
            return
        # Pre-nav preload already loaded icons; just refresh the view once
        if not getattr(self, '_overlay_refresh_done', False):
            self._overlay_refresh_done = True
            QTimer.singleShot(600, self._do_first_overlay_refresh)

    def _do_first_overlay_refresh(self):
        """One-time refresh to ensure overlays render after first navigation."""
        # Only refresh if this widget is currently visible on screen.
        # Background tabs get refreshed in showEvent when they become active.
        if not self.isVisible():
            self._overlay_refresh_done = False  # allow showEvent to trigger it
            return
        if not getattr(self, '_overlay_refreshing', False):
            self._overlay_refreshing = True
            self._overlay_visible_refreshed = True
            try:
                self._refresh_shell_view()
            finally:
                QTimer.singleShot(500, self._clear_overlay_refreshing)

    def _do_overlay_notify(self, path):
        """后台预加载 overlay 图标，完成后回到 UI 线程刷新视图渲染图标。"""
        if path in _OVERLAY_PRELOADED_DIRS:
            # 已预加载过：无需再扫描，直接在 UI 线程刷新一次即可
            self._overlay_refreshing = True
            self._overlay_visible_refreshed = True
            try:
                self._refresh_shell_view()
            finally:
                QTimer.singleShot(500, self._clear_overlay_refreshing)
            return
        # 预加载（SHGetFileInfo 扫描）放到后台线程，避免在 UI 线程同步 Shell 调用卡顿；
        # in-flight 去重，防止同一控件重复提交。
        if getattr(self, '_overlay_preload_inflight', False):
            return
        try:
            if not getattr(self, '_overlay_signals', None):
                self._overlay_signals = _OverlayPreloadSignals(self)
                self._overlay_signals.done.connect(self._on_overlay_preload_done)
            self._overlay_preload_inflight = True
            QThreadPool.globalInstance().start(_OverlayPreloadRunnable(path, self._overlay_signals))
        except Exception as e:
            self._overlay_preload_inflight = False
            debug_print(f"[TortoiseGit] overlay notify error: {e}")

    def _on_overlay_preload_done(self, _path):
        """后台预加载完成（回到 UI 线程）：调用 IShellView::Refresh() 渲染 overlay。"""
        self._overlay_preload_inflight = False
        try:
            self._overlay_refreshing = True
            self._overlay_visible_refreshed = True
            try:
                self._refresh_shell_view()
            finally:
                QTimer.singleShot(500, self._clear_overlay_refreshing)
        except Exception as e:
            debug_print(f"[TortoiseGit] overlay refresh error: {e}")

    def _clear_overlay_refreshing(self):
        self._overlay_refreshing = False

    def _refresh_shell_view(self):
        """Call IShellView::Refresh() to force the DefView to re-render items."""
        if not self._browser:
            return
        try:
            from comtypes import GUID as _GUID
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_result = self._browser.GetCurrentView(ctypes.byref(iid_sv))
            iface_ptr = int(ppv_result) if ppv_result else 0
            if not iface_ptr:
                return
            try:
                # IShellView vtable layout (inherits IOleWindow):
                # [QI(0), AddRef(1), Release(2), GetWindow(3),
                #  ContextSensitiveHelp(4), TranslateAccelerator(5),
                #  EnableModeless(6), UIActivate(7), Refresh(8), ...]
                _vp_size = ctypes.sizeof(ctypes.c_void_p)
                vtable_ptr = ctypes.c_void_p.from_address(iface_ptr).value
                # Refresh is at vtable index 8
                fn_addr = ctypes.c_void_p.from_address(
                    vtable_ptr + 8 * _vp_size).value
                # IShellView::Refresh(this) → HRESULT
                _REFRESH = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p)
                refresh_fn = _REFRESH(fn_addr)
                hr = refresh_fn(iface_ptr)
                debug_print(f"[TortoiseGit] IShellView::Refresh() → hr=0x{hr & 0xFFFFFFFF:08X}")
            finally:
                # Release the IShellView
                _vp_size = ctypes.sizeof(ctypes.c_void_p)
                vtable_ptr = ctypes.c_void_p.from_address(iface_ptr).value
                release_addr = ctypes.c_void_p.from_address(
                    vtable_ptr + 2 * _vp_size).value
                _RELEASE = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _RELEASE(release_addr)(iface_ptr)
        except Exception as e:
            debug_print(f"[TortoiseGit] _refresh_shell_view error: {e}")

    # ── Qt event overrides ────────────────────────────────────────────────────

    def showEvent(self, event):
        super().showEvent(event)
        if not self._init_ok:
            if self._ensure_browser() and self._pending_path:
                path, self._pending_path = self._pending_path, None
                self._navigate(path)
        elif self._browser and _TORTOISE_OVERLAY_PATCHED:
            # Tab became visible: ensure overlays are rendered
            # Refresh is needed because IShellView::Refresh on hidden views is a no-op
            url = getattr(self, '_location_url', '')
            if url and url.startswith('file:'):
                norm = self._url_to_path(url)
                if norm:
                    if norm not in _OVERLAY_PRELOADED_DIRS:
                        # Not yet preloaded - do full preload + refresh
                        QTimer.singleShot(300, lambda p=norm: self._do_overlay_notify(p))
                    elif not getattr(self, '_overlay_visible_refreshed', False):
                        # Already preloaded but never refreshed while visible
                        self._overlay_visible_refreshed = True
                        QTimer.singleShot(200, self._refresh_shell_view)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._browser:
            try:
                rc = ctypes.wintypes.RECT(0, 0, self.width(), self.height())
                self._browser.SetRect(None, rc)
            except Exception as e:
                debug_print(f"[IExplorerBrowser] SetRect failed: {e}")

    # ── Cleanup ───────────────────────────────────────────────────────────────

    def cleanup(self):
        # 作废任何在途的后台 PIDL 解析，并断开信号，避免结果回调已销毁的 COM 浏览器
        self._nav_generation += 1
        resolver = getattr(self, '_nav_resolver', None)
        if resolver is not None:
            try:
                resolver.resolved.disconnect()
            except Exception:
                pass
            try:
                if resolver.isRunning():
                    resolver.wait(200)  # SHParseDisplayName 无法中断，短等即可，结果会被作废
            except Exception:
                pass
            self._nav_resolver = None
        # 断开后台 overlay 预加载信号，避免延迟到达的 QThreadPool 结果回调已销毁的控件
        sig = getattr(self, '_overlay_signals', None)
        if sig is not None:
            try:
                sig.done.disconnect()
            except Exception:
                pass
            try:
                sig.deleteLater()
            except Exception:
                pass
            self._overlay_signals = None
        self._overlay_preload_inflight = False
        if self._browser is None:
            return
        try:
            if self._cookie.value:
                self._browser.Unadvise(self._cookie)
        except Exception:
            dbg_exc("IExplorerBrowser.Unadvise")
        try:
            self._browser.Destroy()
        except Exception:
            dbg_exc("IExplorerBrowser.Destroy")
        try:
            hwnd = int(self.winId())
            _ieb_keyboard_filter.unregister_ieb_hwnd(hwnd)
        except Exception:
            dbg_exc("IExplorerBrowser.unregister_hwnd")
        self._browser  = None
        self._nav_sink = None
        self._init_ok  = False

    def hibernate(self, path):
        """销毁 Shell 视图以释放资源；下次显示时在 showEvent 中按 path 重建。"""
        if not self._init_ok or not path:
            return False
        self.cleanup()
        self._pending_path = path
        self._overlay_refresh_done = False
        self._overlay_visible_refreshed = False
        return True

    def closeEvent(self, event):
        self.cleanup()
        super().closeEvent(event)

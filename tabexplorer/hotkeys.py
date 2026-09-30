"""快捷键定义、解析、校验与低级键盘钩子。"""

import ctypes.wintypes
import os
import threading

from PyQt5.QtCore import pyqtSignal, QObject

from .i18n import tr
from .debuglog import debug_print


SHORTCUT_POLL_ACTIVE_MS = 60  # 主窗口激活时快捷键轮询频率（需足够快以捕捉"同时按住"的短促组合键）
SHORTCUT_POLL_INACTIVE_MS = 500  # 主窗口非激活时快捷键轮询频率

# 可自定义快捷键：(命令, 启用开关键, 默认按键, 说明)；启用开关为 None 的命令始终启用
HOTKEY_COMMANDS = (
    ('new_tab', 'new_tab', 'Ctrl+T', '新建标签页'),
    ('close_tab', 'close_tab', 'Ctrl+W', '关闭当前标签页'),
    ('reopen_tab', 'reopen_tab', 'Ctrl+Shift+T', '恢复关闭的标签页'),
    ('next_tab', 'switch_tab', 'Ctrl+Tab', '下一个标签'),
    ('prev_tab', 'switch_tab', 'Ctrl+Shift+Tab', '上一个标签'),
    ('search', 'search', 'Ctrl+F', '打开搜索对话框'),
    ('quick_find_current_dir', 'quick_find_current_dir', 'Ctrl+G', '检索当前目录文件夹/文件名'),
    ('go_back', 'navigate', 'Alt+Left', '后退'),
    ('go_forward', 'navigate', 'Alt+Right', '前进'),
    ('go_up', 'go_up', 'Alt+Up', '返回上级目录'),
    ('refresh', 'refresh', 'F5', '刷新当前路径'),
    ('add_bookmark', 'add_bookmark', 'Ctrl+D', '添加当前路径到书签'),
    ('quick_copy', 'quick_copy', 'Alt+C', '快速复制选中项（后台）'),
    ('quick_paste', 'quick_paste', 'Alt+V', '快速粘贴到当前目录（后台）'),
    ('quick_delete', 'quick_delete', 'Alt+Del', '快速删除选中项'),
    ('cancel_file_op', 'cancel_file_op', 'Alt+Q', '取消后台复制/删除'),
    ('copy_filename', 'copy_filename', 'Alt+Z', '复制选中文件名（含后缀）'),
    ('copy_filepath', 'copy_filepath', 'Alt+X', '复制文件路径\\文件名'),
    ('split_view', 'split_view', 'F3', '左右分屏对比'),
    ('insert_group_bookmark', 'insert_group_bookmark', 'F4', '插入标签分组'),
    ('focus_path_bar', None, 'Ctrl+L', '编辑路径栏'),
    ('toggle_ai_panel', None, 'Ctrl+Shift+A', '显示/隐藏 AI 助手面板'),
)
_HOTKEY_ENABLE_KEYS = {command: enable_key for command, enable_key, _default, _label in HOTKEY_COMMANDS}
_HOTKEY_KEY_CODES = {
    **{chr(code): code for code in range(0x41, 0x5B)},
    **{str(digit): 0x30 + digit for digit in range(10)},
    **{f'F{number}': 0x6F + number for number in range(1, 25)},
    'Tab': 0x09, 'Backspace': 0x08, 'Enter': 0x0D, 'Esc': 0x1B, 'Space': 0x20,
    'PgUp': 0x21, 'PgDown': 0x22, 'End': 0x23, 'Home': 0x24,
    'Left': 0x25, 'Up': 0x26, 'Right': 0x27, 'Down': 0x28, 'Ins': 0x2D, 'Del': 0x2E,
    ';': 0xBA, '=': 0xBB, ',': 0xBC, '-': 0xBD, '.': 0xBE, '/': 0xBF, '`': 0xC0,
    '[': 0xDB, '\\': 0xDC, ']': 0xDD, "'": 0xDE,
}
_HOTKEY_KEY_NAMES = {code: name for name, code in _HOTKEY_KEY_CODES.items()}
# Qt 的 PortableText 名称与手写配置的常见别名；Backtab 即 Shift+Tab
_HOTKEY_KEY_LOOKUP = {
    **{name.lower(): name for name in _HOTKEY_KEY_CODES},
    'delete': 'Del', 'insert': 'Ins', 'return': 'Enter', 'escape': 'Esc',
    'pageup': 'PgUp', 'pagedown': 'PgDown', 'backtab': 'Tab',
}
# 钩子不拦截按键，绑定这些组合会与 Explorer 的文件操作同时生效
_EXPLORER_RESERVED_HOTKEYS = frozenset(('Ctrl+C', 'Ctrl+X', 'Ctrl+V', 'Ctrl+A', 'Ctrl+Z', 'Ctrl+Y',
                                        'Ctrl+Shift+N', 'F2', 'Alt+Enter', 'Alt+F4'))


def _parse_hotkey(text):
    """'Ctrl+Shift+T' -> (ctrl, shift, alt, vk)；不支持的写法返回 None。"""
    parts = [part.strip() for part in str(text or '').split('+')]
    name = _HOTKEY_KEY_LOOKUP.get(parts[-1].lower())
    modifiers = {part.lower() for part in parts[:-1]}
    if name is None or len(modifiers) != len(parts) - 1 or not modifiers <= {'ctrl', 'shift', 'alt'}:
        return None
    shift = 'shift' in modifiers or parts[-1].lower() == 'backtab'
    return ('ctrl' in modifiers, shift, 'alt' in modifiers, _HOTKEY_KEY_CODES[name])


def _format_hotkey(ctrl, shift, alt, vk):
    modifiers = [name for flag, name in ((ctrl, 'Ctrl'), (alt, 'Alt'), (shift, 'Shift')) if flag]
    return '+'.join(modifiers + [_HOTKEY_KEY_NAMES[vk]])


def _hotkey_bindings(config):
    """合并用户设置与默认值，返回 {命令: 按键文本}；空字符串表示未绑定。"""
    stored = config.get('hotkey_bindings') if isinstance(config, dict) else None
    stored = stored if isinstance(stored, dict) else {}
    bindings = {}
    for command, _enable_key, default, _label in HOTKEY_COMMANDS:
        value = stored.get(command, default)
        bindings[command] = value if isinstance(value, str) else default
    return bindings


def _hotkey_table(config):
    """{(ctrl, shift, alt, vk): 命令}；无法解析的绑定被忽略。"""
    table = {}
    for command, text in _hotkey_bindings(config).items():
        parsed = _parse_hotkey(text)
        if parsed is not None:
            table.setdefault(parsed, command)
    return table


def _hotkey_text(config, command):
    parsed = _parse_hotkey(_hotkey_bindings(config).get(command, ''))
    return _format_hotkey(*parsed) if parsed else ''


def _hotkey_hint(config, label, command):
    key = _hotkey_text(config, command)
    return f"{label} ({key})" if key else label


def _hotkey_reserved_combos(config):
    """不交给 Shell 快捷键翻译的 (ctrl, alt, vk)，与旧版按 Ctrl/Alt+键码判断的方式一致。"""
    combos = {(ctrl, alt, vk) for ctrl, _shift, alt, vk in _hotkey_table(config)}
    hotkeys = config.get('hotkeys', {}) if isinstance(config, dict) else {}
    if hotkeys.get('switch_tab_number', True):
        combos.update((True, False, vk) for vk in range(0x31, 0x3A))
    return frozenset(combos)


def _validate_hotkey_bindings(bindings, number_switch_enabled=True):
    """返回错误说明列表：无法识别、缺少修饰键、与 Explorer 或 Ctrl+数字冲突、重复分配。"""
    labels = {command: tr(label) for command, _enable_key, _default, label in HOTKEY_COMMANDS}
    errors, owners = [], {}
    for command, text in bindings.items():
        if not text:
            continue
        parsed = _parse_hotkey(text)
        if parsed is None:
            errors.append(tr("{}：不支持的按键 {}").format(labels[command], text))
            continue
        ctrl, shift, alt, vk = parsed
        key = _format_hotkey(*parsed)
        if not (ctrl or alt or 0x70 <= vk <= 0x87):
            errors.append(tr("{}：{} 需要包含 Ctrl 或 Alt（F1–F24 可单独使用）").format(labels[command], key))
        elif key in _EXPLORER_RESERVED_HOTKEYS:
            errors.append(tr("{}：{} 是 Explorer 文件操作快捷键").format(labels[command], key))
        elif number_switch_enabled and ctrl and not shift and not alt and 0x31 <= vk <= 0x39:
            errors.append(tr("{}：{} 已用于切换到第 N 个标签").format(labels[command], key))
        owners.setdefault(key, []).append(labels[command])
    for key, names in owners.items():
        if len(names) > 1:
            errors.append(tr("{} 同时分配给：{}").format(key, ' / '.join(names)))
    return errors


# 标题栏按钮提示中显示的快捷键：按钮属性名 -> (说明, 命令)
_TOOLBAR_HOTKEY_TOOLTIPS = {
    'back_button': ('后退', 'go_back'),
    'forward_button': ('前进', 'go_forward'),
    'add_tab_button': ('新建标签页', 'new_tab'),
    'reopen_tab_button': ('恢复关闭的标签页', 'reopen_tab'),
    'search_button': ('搜索当前文件夹', 'search'),
    'insert_group_btn': ('插入分组', 'insert_group_bookmark'),
    'split_view_btn': ('分屏对比', 'split_view'),
    'ai_chat_btn': ('AI 助手面板', 'toggle_ai_panel'),
}


def _toolbar_hotkey_tooltip(owner, name):
    label, command = _TOOLBAR_HOTKEY_TOOLTIPS[name]
    return _hotkey_hint(getattr(owner, 'config', {}), tr(label), command)


def _foreground_pid():
    """前台窗口所属进程 PID；没有前台窗口时返回 0。"""
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    pid = ctypes.c_ulong(0)
    if hwnd:
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


class _KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ('vkCode', ctypes.wintypes.DWORD),
        ('scanCode', ctypes.wintypes.DWORD),
        ('flags', ctypes.wintypes.DWORD),
        ('time', ctypes.wintypes.DWORD),
        ('dwExtraInfo', ctypes.c_size_t),
    ]


class _ShortcutKeyHook(QObject):
    """WH_KEYBOARD_LL 快捷键监听：钩子运行在独立线程，UI 线程繁忙时不会触发系统超时摘除钩子。

    只上报本进程窗口在前台时的首次按下（忽略自动重复），不拦截按键。"""
    keyPressed = pyqtSignal(int, bool, bool, bool)  # vk, ctrl, shift, alt

    _MODIFIER_VKS = frozenset((0x10, 0x11, 0x12, 0x5B, 0x5C, 0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5))

    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread = None
        self._thread_id = 0
        self._installed = False
        self._ready = threading.Event()

    def start(self):
        if os.name != 'nt':
            return False
        self._ready.clear()
        self._thread = threading.Thread(target=self._run, name='ShortcutKeyHook', daemon=True)
        self._thread.start()
        self._ready.wait(2.0)
        return self._installed

    def stop(self):
        thread, self._thread = self._thread, None
        if thread is None:
            return
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(self._thread_id, 0x0012, 0, 0)  # WM_QUIT
        thread.join(1.0)

    def _run(self):
        wt = ctypes.wintypes
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        hook_proc_type = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_int, wt.WPARAM, wt.LPARAM)
        user32.SetWindowsHookExW.argtypes = (ctypes.c_int, hook_proc_type, wt.HINSTANCE, wt.DWORD)
        user32.SetWindowsHookExW.restype = ctypes.c_void_p
        user32.CallNextHookEx.argtypes = (ctypes.c_void_p, ctypes.c_int, wt.WPARAM, wt.LPARAM)
        user32.CallNextHookEx.restype = ctypes.c_ssize_t
        user32.UnhookWindowsHookEx.argtypes = (ctypes.c_void_p,)
        user32.GetForegroundWindow.restype = wt.HWND
        user32.GetWindowThreadProcessId.argtypes = (wt.HWND, ctypes.POINTER(wt.DWORD))
        user32.GetAsyncKeyState.restype = ctypes.c_short
        user32.GetMessageW.argtypes = (ctypes.POINTER(wt.MSG), wt.HWND, ctypes.c_uint, ctypes.c_uint)
        user32.PeekMessageW.argtypes = (ctypes.POINTER(wt.MSG), wt.HWND, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint)
        kernel32.GetModuleHandleW.restype = wt.HMODULE
        pid = os.getpid()
        modifiers = self._MODIFIER_VKS

        def is_down(vk):
            return bool(user32.GetAsyncKeyState(vk) & 0x8000)

        def hook_proc(n_code, w_param, l_param):
            try:
                if n_code == 0 and w_param in (0x0100, 0x0104):  # WM_KEYDOWN / WM_SYSKEYDOWN
                    vk = int(ctypes.cast(l_param, ctypes.POINTER(_KBDLLHOOKSTRUCT)).contents.vkCode)
                    # 钩子先于异步键状态更新：此时已按下说明是自动重复
                    if vk not in modifiers and not is_down(vk):
                        hwnd = user32.GetForegroundWindow()
                        owner = wt.DWORD(0)
                        if hwnd:
                            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
                        if owner.value == pid:
                            self.keyPressed.emit(vk, is_down(0x11), is_down(0x10), is_down(0x12))
            except Exception:
                pass
            return user32.CallNextHookEx(None, n_code, w_param, l_param)

        callback = hook_proc_type(hook_proc)
        msg = wt.MSG()
        user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)  # 创建消息队列，保证 WM_QUIT 可投递
        self._thread_id = kernel32.GetCurrentThreadId()
        handle = user32.SetWindowsHookExW(13, callback, kernel32.GetModuleHandleW(None), 0)
        self._installed = bool(handle)
        self._ready.set()
        if not handle:
            debug_print(f"[ShortcutHook] SetWindowsHookExW(WH_KEYBOARD_LL) failed err={ctypes.get_last_error()}")
            return
        try:
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        finally:
            user32.UnhookWindowsHookEx(handle)
            self._thread_id = 0

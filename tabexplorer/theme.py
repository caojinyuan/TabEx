"""浅色/深色主题：Qt 调色板、样式表换色、Windows 原生深色（标题栏、资源管理器视图、原生菜单）。

样式表与自绘颜色一律按浅色主题书写；深色主题下由 qss()/fg()/bg()/line() 按颜色用途换算。
bind_style() 记录浅色原文，切换主题时由 set_dark() 统一重新换算。"""

import ctypes
import ctypes.wintypes
import re
import sys
from functools import lru_cache

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QPalette
from PyQt5.QtWidgets import QApplication, QToolTip

from .debuglog import debug_print


THEME_MODES = ('system', 'light', 'dark')
_STYLE_PROPERTY = 'tabexThemeCss'
_FRAME_PROPERTY = 'tabexDarkFrame'
_ICON_COLOR_DARK = '#d4d4d4'

_dark = False
_original_style = None
_original_palette = None
_widget_hooks = []


def normalize_mode(mode):
    return mode if mode in THEME_MODES else 'system'


def system_prefers_dark():
    """读取 Windows“应用模式”：AppsUseLightTheme 为 0 表示深色。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize') as key:
            return winreg.QueryValueEx(key, 'AppsUseLightTheme')[0] == 0
    except OSError:
        return False


def resolve_dark(mode):
    mode = normalize_mode(mode)
    return mode == 'dark' or (mode == 'system' and system_prefers_dark())


def is_dark():
    return _dark


def icon_color():
    """深色主题下线条图标的颜色；浅色主题返回 None 表示使用图标原色。"""
    return _ICON_COLOR_DARK if _dark else None


@lru_cache(maxsize=1024)
def _to_dark(role, value):
    color = QColor(value)
    if not color.isValid():
        return value
    hue, sat, light, alpha = color.getHslF()
    gray = hue < 0 or sat < 0.2
    if role == 'fg':
        # 浅色文字本就用于深色/彩色底（如强调按钮上的白字），保持不变
        if light >= (0.8 if gray else 0.75):
            return value
        light = 0.95 - light * 0.8 if gray else max(0.62, 0.95 - light * 0.8)
    elif role == 'bg':
        # 深色或强调色背景（L<0.55）保持不变，浅色背景按亮度反转到深色区间
        if light < 0.55:
            return value
        light = min(0.45, 0.118 + (1.0 - light) * 1.05)
        sat = 0.0 if gray else sat * 0.45
    else:
        if light < 0.55 or (not gray and light < 0.75):
            return value
        light = min(0.5, max(0.24, 0.118 + (1.0 - light) * 1.25))
        sat = 0.0 if gray else sat * 0.5
    result = QColor.fromHslF(max(hue, 0.0), sat, light, alpha)
    if result.alpha() == 255:
        return result.name()
    return f'rgba({result.red()}, {result.green()}, {result.blue()}, {result.alpha()})'


def fg(value):
    """文字/线条前景色。"""
    return _to_dark('fg', value) if _dark else value


def bg(value):
    """背景填充色。"""
    return _to_dark('bg', value) if _dark else value


def line(value):
    """边框/分隔线颜色。"""
    return _to_dark('line', value) if _dark else value


_DECLARATION = re.compile(r'([A-Za-z-]+)(\s*:\s*)([^;{}]*)')
_COLOR = re.compile(r'#[0-9A-Fa-f]{3,8}\b|\b(?:white|black)\b')


def _role_for(prop):
    prop = prop.lower()
    if prop in ('color', 'selection-color'):
        return 'fg'
    if prop.startswith('background') or prop == 'alternate-background-color':
        return 'bg'
    if prop.startswith('border') or prop in ('gridline-color', 'outline', 'outline-color'):
        return 'line'
    return None


def qss(css):
    """把按浅色主题书写的样式表换算为当前主题；selection-background-color 等强调色保持不变。"""
    if not _dark or not css:
        return css

    def convert(match):
        role = _role_for(match.group(1))
        if role is None:
            return match.group(0)
        value = _COLOR.sub(lambda color: _to_dark(role, color.group(0)), match.group(3))
        return match.group(1) + match.group(2) + value

    return _DECLARATION.sub(convert, css)


def bind_style(widget, css):
    """设置随主题切换自动重新换算的样式表。"""
    widget.setProperty(_STYLE_PROPERTY, css)
    widget.setStyleSheet(qss(css))


def add_widget_hook(hook):
    """登记主题切换后对每个控件执行的刷新函数（如重新着色图标）。"""
    if hook not in _widget_hooks:
        _widget_hooks.append(hook)


def restyle_widgets():
    for widget in QApplication.allWidgets():
        css = widget.property(_STYLE_PROPERTY)
        if css is not None:
            widget.setStyleSheet(qss(css))
        for hook in _widget_hooks:
            hook(widget)


def _windows_build():
    try:
        return sys.getwindowsversion().build
    except AttributeError:
        return 0


def _set_native_app_mode(dark):
    """进程级 Win32 深色（资源管理器视图、原生右键菜单）；已创建的 Shell 视图需重建才完整换色。"""
    if _windows_build() < 18362:
        return
    try:
        uxtheme = ctypes.WinDLL('uxtheme')
        # 未公开导出序号：135 SetPreferredAppMode（2=ForceDark，0=Default）、104 刷新配色策略、136 刷新菜单主题
        ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int)((135, uxtheme))(2 if dark else 0)
        ctypes.WINFUNCTYPE(None)((104, uxtheme))()
        ctypes.WINFUNCTYPE(None)((136, uxtheme))()
    except (AttributeError, OSError) as error:
        debug_print(f"[Theme] native dark mode unavailable: {error}")


def _set_dwm_dark(hwnd, dark):
    build = _windows_build()
    if build < 17763:
        return False
    value = ctypes.c_int(1 if dark else 0)
    attribute = 20 if build >= 18985 else 19  # DWMWA_USE_IMMERSIVE_DARK_MODE
    try:
        return ctypes.windll.dwmapi.DwmSetWindowAttribute(
            ctypes.wintypes.HWND(hwnd), attribute, ctypes.byref(value), ctypes.sizeof(value)) == 0
    except (AttributeError, OSError):
        return False


def apply_window_frame(widget, force=False):
    """按当前主题设置顶层窗口系统标题栏的深浅色。"""
    try:
        if not widget.isWindow() or widget.windowFlags() & Qt.FramelessWindowHint:
            return
        if widget.windowType() in (Qt.Popup, Qt.ToolTip, Qt.SplashScreen, Qt.Desktop):
            return
        if bool(widget.property(_FRAME_PROPERTY)) == _dark and not force:
            return
        hwnd = int(widget.effectiveWinId() or 0)
        if not hwnd or not _set_dwm_dark(hwnd, _dark):
            return
        widget.setProperty(_FRAME_PROPERTY, _dark)
        if widget.isVisible():
            # RDW_INVALIDATE | RDW_FRAME：已显示的窗口立即重绘标题栏
            ctypes.windll.user32.RedrawWindow(ctypes.wintypes.HWND(hwnd), None, None, 0x0001 | 0x0400)
    except RuntimeError:
        pass


def _dark_palette():
    palette = QPalette()
    colors = {
        QPalette.Window: '#202020', QPalette.WindowText: '#e8e8e8', QPalette.Base: '#1b1b1b',
        QPalette.AlternateBase: '#262626', QPalette.ToolTipBase: '#2b2b2b', QPalette.ToolTipText: '#e8e8e8',
        QPalette.PlaceholderText: '#8c8c8c', QPalette.Text: '#e8e8e8', QPalette.Button: '#2d2d2d',
        QPalette.ButtonText: '#e8e8e8', QPalette.BrightText: '#ff6b6b', QPalette.Light: '#3c3c3c',
        QPalette.Midlight: '#333333', QPalette.Mid: '#2a2a2a', QPalette.Dark: '#151515',
        QPalette.Shadow: '#000000', QPalette.Highlight: '#2f6fdb', QPalette.HighlightedText: '#ffffff',
        QPalette.Link: '#6ea8ff', QPalette.LinkVisited: '#b39ddb',
    }
    for role, value in colors.items():
        palette.setColor(role, QColor(value))
    for role, value in ((QPalette.WindowText, '#6e6e6e'), (QPalette.Text, '#6e6e6e'),
                        (QPalette.ButtonText, '#6e6e6e'), (QPalette.Highlight, '#3a3a3a'),
                        (QPalette.HighlightedText, '#9a9a9a')):
        palette.setColor(QPalette.Disabled, role, QColor(value))
    return palette


def set_dark(dark):
    """切换深浅色；返回是否发生变化。已存在的 Shell 视图需由调用方重建。"""
    global _dark, _original_style, _original_palette
    dark = bool(dark)
    if dark == _dark:
        return False
    app = QApplication.instance()
    if _original_style is None:
        _original_style = app.style().objectName() or 'windowsvista'
        _original_palette = QPalette(app.palette())
    _dark = dark
    _set_native_app_mode(dark)
    app.setStyle('Fusion' if dark else _original_style)
    palette = _dark_palette() if dark else QPalette(_original_palette)
    app.setPalette(palette)
    QToolTip.setPalette(palette)
    restyle_widgets()
    for widget in app.topLevelWidgets():
        if widget.isVisible():
            apply_window_frame(widget)
    debug_print(f"[Theme] switched to {'dark' if dark else 'light'}")
    return True

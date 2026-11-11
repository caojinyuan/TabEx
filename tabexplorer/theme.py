"""颜色主题：参照 VS Code 的主题注册表、样式表换色、Qt 调色板与 Windows 原生深色（标题栏、资源管理器视图、原生菜单）。

样式表与自绘颜色一律按浅色主题书写：
- `@令牌`（如 `background: @statusbar-bg`）取当前主题的界面配色，用于标题栏、标签、状态栏等主框架；
- 其余字面颜色由 qss()/fg()/bg()/line() 按用途换算到当前主题（Light Modern、Light+ 保持原样）；
  带强调色板的深色主题（如 Monokai Black）把彩色文字换成色板里色相最近的颜色，红绿蓝等含义不变。
bind_style() 记录原文，切换主题时由 set_theme() 统一重新换算。"""

import ctypes
import ctypes.wintypes
import re
import sys
from functools import lru_cache

from PyQt5.QtCore import QPointF, QRectF, Qt
from PyQt5.QtGui import QColor, QPainter, QPalette
from PyQt5.QtWidgets import (QApplication, QPlainTextEdit, QProxyStyle, QStyle, QStyleOptionButton, QTextEdit,
                             QToolTip)

from .debuglog import debug_print
from .i18n import tr


_STYLE_PROPERTY = 'tabexThemeCss'
_FRAME_PROPERTY = 'tabexDarkFrame'
_CAPTION_PROPERTY = 'tabexCaptionTheme'
# 原生资源管理器视图只有两种固定底色，主题色调靠叠加半透明层实现
_NATIVE_VIEW_DARK = (25, 25, 25)
_NATIVE_VIEW_LIGHT = (255, 255, 255)


def _rgba(value):
    """'#RGB'、'#RRGGBB' 或 VS Code 写法的 '#RRGGBBAA' -> (r, g, b, a)。"""
    text = value.lstrip('#')
    if len(text) == 3:
        text = ''.join(char * 2 for char in text)
    red, green, blue = (int(text[index:index + 2], 16) for index in (0, 2, 4))
    return red, green, blue, int(text[6:8], 16) if len(text) == 8 else 255


def _hex(red, green, blue, alpha=255):
    red, green, blue, alpha = (max(0, min(255, int(round(value)))) for value in (red, green, blue, alpha))
    return f'#{red:02X}{green:02X}{blue:02X}' + ('' if alpha == 255 else f'{alpha:02X}')


def _mix(first, second, amount):
    """按 amount（0 为 first，1 为 second）混合两种颜色，结果不透明。"""
    start, end = _rgba(first), _rgba(second)
    return _hex(*(start[index] + (end[index] - start[index]) * amount for index in range(3)))


def _alpha(value, alpha):
    return _hex(*_rgba(value)[:3], alpha)


def _lightness(value):
    red, green, blue, _alpha_value = _rgba(value)
    return (max(red, green, blue) + min(red, green, blue)) / 510.0


def _luminance(value):
    linear = [channel / 12.92 if channel <= 0.03928 else ((channel + 0.055) / 1.055) ** 2.4
              for channel in (component / 255.0 for component in _rgba(value)[:3])]
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]


def _contrast(first, second):
    """WCAG 对比度（1–21）。"""
    high, low = sorted((_luminance(first), _luminance(second)), reverse=True)
    return (high + 0.05) / (low + 0.05)


def _separator(base, toward, target=1.9):
    """从 base 向 toward 逐步混合，直到与 base 的对比度达到 target：各主题的分隔线同样清楚。"""
    for step in range(1, 21):
        color = _mix(base, toward, step / 20.0)
        if _contrast(color, base) >= target:
            return color
    return toward


def _neutral_offset(editor):
    """主题底色与同亮度纯灰的 RGB 差值：中性灰加上它即带上主题色调。"""
    red, green, blue, _alpha_value = _rgba(editor)
    gray = round(_lightness(editor) * 255)
    return red - gray, green - gray, blue - gray


def _dark_remap(editor, **params):
    """深色主题的字面颜色换算参数：浅色背景按亮度反转到以 editor 为底的深色区间。

    accents 为强调色板时，彩色文字不再按亮度换算，而是取色板中色相最近的颜色。"""
    base = _lightness(editor)
    remap = dict(mode='dark', base=base, offset=_neutral_offset(editor), bg_gain=1.05, bg_cap=0.45,
                 line_gain=1.25, line_floor=base + 0.12, line_cap=0.5, fg_top=0.95, fg_gain=0.8,
                 fg_tint=0.35, contrast_line=None, accents=())
    remap.update(params)
    return remap


def _light_remap(editor, **params):
    """带色调或高对比度的浅色主题：白底换成 editor，compress<1 时浅灰更接近白，fg_scale<1 时文字更深。"""
    remap = dict(mode='light', scale=_lightness(editor), offset=_neutral_offset(editor), compress=1.0,
                 fg_scale=1.0, fg_tint=0.5, contrast_line=None)
    remap.update(params)
    return remap


def _complete_colors(dark, prominent, colors):
    """按 VS Code 的默认规则补齐未单独指定的界面令牌。"""
    tokens = dict(colors)
    put = tokens.setdefault
    put('chrome-fg', tokens['fg'])
    put('fg-disabled', _mix(tokens['fg-muted'], tokens['editor-bg'], 0.45))
    put('link-fg', tokens['focus-border'])
    put('icon-fg', tokens['chrome-fg'])
    put('titlebar-bg', tokens['chrome-bg'])
    put('titlebar-fg', tokens['chrome-fg'])
    put('titlebar-border', tokens['chrome-border'])
    put('window-border', None)  # Windows 11 系统窗口边框色；None 为系统默认
    put('toolbar-hover-bg', '#5A5D5E50' if dark else '#B8B8B850')
    put('toolbar-active-bg', '#63666780' if dark else '#A6A6A680')
    put('toolbar-checked-bg', _alpha(tokens['focus-border'], 0x40))
    put('toolbar-checked-border', tokens['focus-border'])
    put('cc-bg', '#FFFFFF0D' if dark else '#0000000A')
    put('cc-hover-bg', '#FFFFFF1A' if dark else '#00000014')
    put('cc-fg', tokens['titlebar-fg'])
    put('cc-border', _alpha(tokens['titlebar-fg'], 0x33))
    put('tab-bar-bg', tokens['chrome-bg'])
    put('tab-active-bg', tokens['editor-bg'])
    put('tab-active-fg', tokens['fg'])
    put('tab-inactive-bg', tokens['tab-bar-bg'])
    put('tab-inactive-fg', tokens['fg-muted'])
    put('tab-hover-bg', tokens['tab-active-bg'])
    put('tab-border', tokens['chrome-border'])
    put('tab-separator', tokens['tab-border'] if _contrast(tokens['tab-border'], tokens['tab-inactive-bg']) >= 1.9
        else _separator(tokens['tab-inactive-bg'], tokens['fg']))
    put('tab-active-border-top', tokens['focus-border'])
    put('tab-unfocused-border-top', _mix(tokens['fg-muted'], tokens['tab-active-bg'], 0.4))
    put('tab-unfocused-active-fg', _mix(tokens['tab-active-fg'], tokens['tab-active-bg'], 0.3))
    put('crumb-fg', tokens['fg-muted'])
    put('crumb-separator-fg', tokens['crumb-fg'])
    put('crumb-active-fg', tokens['fg'])
    put('crumb-hover-bg', tokens['toolbar-hover-bg'])
    put('pane-inactive-border', _mix(tokens['fg-muted'], tokens['editor-bg'], 0.45))
    put('statusbar-bg', tokens['chrome-bg'])
    put('statusbar-fg', tokens['fg'])
    put('statusbar-border', tokens['chrome-border'])
    put('statusbar-emphasis-fg', tokens['statusbar-fg'] if prominent else tokens['link-fg'])
    # 状态栏按钮底下是标签页背景，悬停色必须不透明
    put('statusbar-hover-bg', _mix(tokens['statusbar-bg'], '#FFFFFF' if dark or prominent else '#000000', 0.12))
    put('input-bg', _mix(tokens['editor-bg'], '#FFFFFF', 0.08) if dark else '#FFFFFF')
    put('input-fg', tokens['fg'])
    put('input-border', tokens['chrome-border'])
    # 输入框、按钮、勾选框的外框（Fusion 绘制的控件）；浅色主题沿用 Fusion 自己的外框（None）
    put('control-border', tokens['input-border'] if dark else None)
    put('input-placeholder', tokens['fg-muted'])
    put('button-bg', tokens['focus-border'])
    put('button-fg', '#FFFFFF')
    put('button-hover-bg', tokens['button-bg'])
    put('menu-bg', tokens['editor-bg'])
    put('menu-fg', tokens['fg'])
    put('menu-border', tokens['chrome-border'])
    put('menu-selection-bg', tokens['focus-border'])
    put('menu-selection-fg', '#FFFFFF')
    put('menu-selection-border', tokens['menu-selection-bg'])
    put('menu-separator', tokens['menu-border'])
    put('list-hover-bg', '#FFFFFF12' if dark else '#0000000A')
    put('list-active-bg', tokens['menu-selection-bg'])
    put('list-active-fg', tokens['menu-selection-fg'])
    put('list-active-border', tokens['list-active-bg'])
    put('list-highlight-fg', tokens['link-fg'])
    put('list-active-highlight-fg', tokens['list-highlight-fg'])
    put('widget-bg', tokens['menu-bg'])
    put('widget-fg', tokens['fg'])
    put('widget-border', tokens['menu-border'])
    put('widget-shadow', '#0000005C' if dark else '#00000029')
    put('picker-group-fg', tokens['list-highlight-fg'])
    put('picker-group-border', tokens['menu-separator'])
    put('keybinding-bg', '#8080802B' if dark else '#DDDDDD66')
    put('keybinding-fg', tokens['fg'])
    put('keybinding-border', '#33333399' if dark else '#CCCCCC66')
    put('keybinding-bottom-border', '#44444499' if dark else '#BBBBBB66')
    put('scrollbar-bg', '#79797966' if dark else '#64646466')
    put('scrollbar-hover-bg', '#646464B3')
    put('scrollbar-active-bg', '#BFBFBF66' if dark else '#00000099')
    put('tooltip-bg', tokens['widget-bg'])
    put('tooltip-fg', tokens['widget-fg'])
    put('tooltip-border', tokens['widget-border'])
    put('sash-hover', tokens['focus-border'])
    put('selection-bg', '#264F78' if dark else '#ADD6FF')
    put('selection-fg', '#FFFFFF' if dark else '#000000')
    put('palette-window', tokens['editor-bg'])
    put('palette-base', tokens['chrome-bg'] if dark else '#FFFFFF')
    put('palette-alt-base', _mix(tokens['palette-base'], tokens['fg'], 0.05))
    put('palette-button', tokens['input-bg'])
    return tokens


class ColorTheme:
    """一个颜色主题：界面令牌、字面颜色换算参数，以及原生视图上的叠色。"""

    def __init__(self, theme_id, label, kind, colors, remap=None, native_style=False, high_contrast=False,
                 prominent_statusbar=False, tint_view=False):
        self.id = theme_id
        self.label = label
        self.kind = kind
        self.dark = kind == 'dark'
        self.high_contrast = high_contrast
        self.native_style = native_style  # 沿用系统控件风格和调色板（Light Modern / Light+）
        self.prominent_statusbar = prominent_statusbar  # 状态栏是强调色底，彩色文字统一改用状态栏前景色
        self.remap = remap
        self.colors = _complete_colors(self.dark, prominent_statusbar, colors)
        self.view_tint = self._view_tint() if tint_view else None

    @property
    def group(self):
        return 'hc' if self.high_contrast else self.kind

    def _view_tint(self, max_alpha=0.32):
        """把原生视图的固定底色调成主题底色的叠加色 (r, g, b, a)；透明度不超过 max_alpha，避免文字发灰。"""
        native = _NATIVE_VIEW_DARK if self.dark else _NATIVE_VIEW_LIGHT
        target = _rgba(self.colors['editor-bg'])[:3]
        need = 0.0
        for wanted, base in zip(target, native):
            if wanted > base:
                need = max(need, (wanted - base) / (255.0 - base))
            elif wanted < base:
                need = max(need, (base - wanted) / float(base))
        alpha = min(max_alpha, need)
        if alpha < 0.02:
            return None
        channels = [(wanted - (1.0 - alpha) * base) / alpha for wanted, base in zip(target, native)]
        return tuple(max(0, min(255, int(round(value)))) for value in channels) + (int(round(alpha * 255)),)


# Monokai 原色：粉（关键字）、橙（参数）、黄（字符串）、绿（函数名）、青（类型）、紫（常量）
_MONOKAI_ACCENTS = ('#F92672', '#FD971F', '#E6DB74', '#A6E22E', '#66D9EF', '#AE81FF')
# Monokai 注释色：Monokai Black 的输入框、按钮、勾选框及样式表里的浅灰边框统一用它，黑底和深灰底上对比度都 ≥3:1
_MONOKAI_COMMENT = '#75715E'

_THEMES = {theme.id: theme for theme in (
    ColorTheme('light', 'Light Modern', 'light', {
        'editor-bg': '#FFFFFF', 'fg': '#3B3B3B', 'fg-muted': '#6F6F6F', 'focus-border': '#005FB8',
        'chrome-bg': '#F8F8F8', 'chrome-border': '#E5E5E5', 'titlebar-fg': '#1E1E1E', 'icon-fg': None,
        'tab-inactive-fg': '#868686', 'input-border': '#CECECE', 'input-placeholder': '#767676',
        'button-hover-bg': '#0258A8', 'menu-bg': '#FFFFFF', 'menu-border': '#CECECE', 'menu-separator': '#E5E5E5',
        'list-hover-bg': '#F2F2F2', 'list-active-bg': '#E8E8E8', 'list-active-fg': '#000000',
        'list-active-border': '#005FB8', 'list-highlight-fg': '#0066BF', 'widget-bg': '#F8F8F8',
        'widget-border': '#E5E5E5', 'tooltip-border': '#C8C8C8',
    }, native_style=True),
    ColorTheme('light_plus', 'Light+', 'light', {
        'editor-bg': '#FFFFFF', 'fg': '#616161', 'fg-muted': '#717171', 'focus-border': '#0090F1',
        'chrome-bg': '#F3F3F3', 'chrome-border': '#E7E7E7', 'icon-fg': None, 'titlebar-bg': '#DDDDDD',
        'titlebar-fg': '#333333', 'titlebar-border': '#DDDDDD', 'tab-inactive-bg': '#ECECEC',
        'tab-active-fg': '#333333', 'tab-inactive-fg': '#6F6F6F', 'tab-border': '#F3F3F3',
        'tab-active-border-top': '#007ACC', 'tab-hover-bg': '#F6F6F6', 'statusbar-bg': '#007ACC',
        'statusbar-fg': '#FFFFFF', 'statusbar-border': '#007ACC', 'link-fg': '#006AB1',
        'input-border': '#CECECE', 'input-placeholder': '#767676', 'button-bg': '#007ACC',
        'button-hover-bg': '#0062A3', 'menu-bg': '#FFFFFF', 'menu-fg': '#333333', 'menu-border': '#D4D4D4',
        'menu-selection-bg': '#0060C0', 'menu-separator': '#D4D4D4', 'list-hover-bg': '#E8E8E8',
        'list-active-bg': '#0060C0', 'list-active-fg': '#FFFFFF', 'list-highlight-fg': '#0066BF',
        'list-active-highlight-fg': '#BBE7FF', 'widget-bg': '#F3F3F3', 'widget-fg': '#616161',
        'widget-border': '#C8C8C8', 'picker-group-border': '#CCCEDB',
    }, native_style=True, prominent_statusbar=True),
    ColorTheme('solarized_light', 'Solarized Light', 'light', {
        'editor-bg': '#FDF6E3', 'fg': '#53686F', 'fg-muted': '#6C8389', 'focus-border': '#B49471',
        'chrome-bg': '#EEE8D5', 'chrome-border': '#DDD6C1', 'tab-bar-bg': '#D9D2C2',
        'tab-inactive-bg': '#D3CBB7', 'tab-active-fg': '#073642', 'tab-inactive-fg': '#586E75',
        'tab-border': '#DDD6C1', 'tab-active-border-top': '#B58900', 'tab-hover-bg': '#E6DFCB',
        'link-fg': '#268BD2', 'input-bg': '#FDF6E3', 'input-border': '#D3CBB7', 'input-placeholder': '#839496',
        'button-bg': '#AC9D57', 'button-hover-bg': '#9B8D4E', 'menu-bg': '#EEE8D5', 'menu-border': '#D3CBB7',
        'menu-selection-bg': '#DFCA88', 'menu-selection-fg': '#4F4F4F', 'menu-separator': '#D3CBB7',
        'list-hover-bg': '#DFCA8844', 'list-active-bg': '#DFCA88', 'list-active-fg': '#4F4F4F',
        'list-highlight-fg': '#B58900', 'list-active-highlight-fg': '#8A6900', 'widget-bg': '#EEE8D5',
        'widget-border': '#D3CBB7', 'picker-group-fg': '#2AA198', 'selection-bg': '#DFCA88',
        'selection-fg': '#4F4F4F', 'palette-window': '#EEE8D5', 'palette-base': '#FDF6E3',
        'palette-button': '#E6DFC8', 'statusbar-emphasis-fg': '#1A6BA0',
    }, remap=_light_remap('#FDF6E3'), tint_view=True),
    ColorTheme('dark', 'Dark Modern', 'dark', {
        'editor-bg': '#1F1F1F', 'fg': '#CCCCCC', 'fg-muted': '#9D9D9D', 'focus-border': '#0078D4',
        'chrome-bg': '#181818', 'chrome-border': '#2B2B2B', 'tab-active-fg': '#FFFFFF', 'link-fg': '#4DAAFC',
        'input-bg': '#313131', 'input-border': '#3C3C3C', 'input-placeholder': '#989898',
        'button-hover-bg': '#026EC1', 'menu-bg': '#1F1F1F', 'menu-border': '#454545', 'menu-separator': '#454545',
        'list-hover-bg': '#2A2D2E', 'list-active-bg': '#04395E', 'list-active-fg': '#FFFFFF',
        'list-active-border': '#0078D4', 'list-highlight-fg': '#2AAAFF', 'widget-bg': '#222222',
        'widget-border': '#313131', 'tooltip-bg': '#202020', 'tooltip-border': '#454545',
        'picker-group-fg': '#3794FF', 'picker-group-border': '#3C3C3C', 'palette-window': '#1F1F1F',
        'palette-base': '#181818', 'palette-alt-base': '#222222', 'palette-button': '#313131',
        'selection-bg': '#0078D4',
    }, remap=_dark_remap('#1F1F1F')),
    ColorTheme('dark_plus', 'Dark+', 'dark', {
        'editor-bg': '#1E1E1E', 'fg': '#CCCCCC', 'fg-muted': '#9D9D9D', 'focus-border': '#007FD4',
        'chrome-bg': '#252526', 'chrome-border': '#333333', 'titlebar-bg': '#3C3C3C',
        'titlebar-border': '#3C3C3C', 'tab-inactive-bg': '#2D2D2D', 'tab-active-fg': '#FFFFFF',
        'tab-inactive-fg': '#969696', 'tab-border': '#252526', 'tab-active-border-top': '#007ACC',
        'tab-hover-bg': '#323233', 'statusbar-bg': '#007ACC', 'statusbar-fg': '#FFFFFF',
        'statusbar-border': '#007ACC', 'link-fg': '#3794FF', 'input-bg': '#3C3C3C', 'input-border': '#3C3C3C',
        'input-placeholder': '#A6A6A6', 'button-bg': '#0E639C', 'button-hover-bg': '#1177BB',
        'menu-bg': '#252526', 'menu-border': '#454545', 'menu-selection-bg': '#04395E', 'menu-separator': '#606060',
        'list-hover-bg': '#2A2D2E', 'list-active-bg': '#04395E', 'list-active-fg': '#FFFFFF',
        'list-highlight-fg': '#2AAAFF', 'widget-bg': '#252526', 'widget-border': '#454545',
        'picker-group-fg': '#3794FF', 'picker-group-border': '#3F3F46', 'palette-window': '#252526',
        'palette-base': '#1E1E1E', 'palette-alt-base': '#2A2A2A', 'palette-button': '#3A3D41',
    }, remap=_dark_remap('#1E1E1E'), prominent_statusbar=True),
    ColorTheme('gray', 'Gray', 'dark', {
        'editor-bg': '#404040', 'fg': '#E8E8E8', 'fg-muted': '#B4B4B4', 'focus-border': '#4A9EFF',
        'chrome-bg': '#383838', 'chrome-border': '#4F4F4F', 'tab-active-fg': '#FFFFFF',
        'input-bg': '#4A4A4A', 'input-border': '#5A5A5A', 'input-placeholder': '#9A9A9A',
        'button-bg': '#2F6FDB', 'button-hover-bg': '#3A7BE6', 'menu-bg': '#3A3A3A', 'menu-border': '#555555',
        'menu-selection-bg': '#2F6FDB', 'menu-separator': '#555555', 'list-hover-bg': '#FFFFFF14',
        'list-active-bg': '#2F5A8F', 'list-active-fg': '#FFFFFF', 'list-highlight-fg': '#7CC0FF',
        'widget-bg': '#3A3A3A', 'widget-border': '#555555', 'widget-shadow': '#00000066',
        'tooltip-bg': '#4A4A4A', 'tooltip-border': '#606060', 'scrollbar-bg': '#9E9E9E66',
        'scrollbar-hover-bg': '#B0B0B0A0', 'scrollbar-active-bg': '#D0D0D080', 'palette-window': '#404040',
        'palette-base': '#383838', 'palette-alt-base': '#444444', 'palette-button': '#4C4C4C',
        'selection-bg': '#2F6FDB', 'statusbar-emphasis-fg': '#7CC0FF',
    }, remap=_dark_remap('#404040', base=0.25, bg_gain=0.95, bg_cap=0.5, line_floor=0.38, line_cap=0.58),
        tint_view=True),
    ColorTheme('monokai', 'Monokai', 'dark', {
        'editor-bg': '#272822', 'fg': '#F8F8F2', 'fg-muted': '#A59F85', 'focus-border': '#99947C',
        'chrome-bg': '#1E1F1C', 'chrome-border': '#34352F', 'link-fg': '#66D9EF', 'tab-inactive-bg': '#34352F',
        'tab-inactive-fg': '#CCCCC7', 'tab-border': '#1E1F1C', 'tab-active-border-top': '#F92672',
        'tab-hover-bg': '#3E3D32', 'statusbar-bg': '#414339', 'statusbar-border': '#414339',
        'input-bg': '#414339', 'input-border': '#414339', 'input-placeholder': '#90908A',
        'button-bg': '#75715E', 'button-hover-bg': '#8A8570', 'menu-bg': '#1E1F1C', 'menu-fg': '#CCCCCC',
        'menu-border': '#414339', 'menu-selection-bg': '#75715E', 'menu-selection-fg': '#F8F8F2',
        'menu-separator': '#414339', 'list-hover-bg': '#3E3D32', 'list-active-bg': '#75715E',
        'list-active-fg': '#F8F8F2', 'list-highlight-fg': '#E6DB74', 'widget-bg': '#1E1F1C',
        'widget-border': '#414339', 'picker-group-fg': '#A6E22E', 'tooltip-bg': '#3E3D32',
        'tooltip-border': '#75715E', 'palette-window': '#272822', 'palette-base': '#1E1F1C',
        'palette-alt-base': '#2D2E27', 'palette-button': '#414339', 'selection-bg': '#75715E',
        'selection-fg': '#F8F8F2',
    }, remap=_dark_remap('#272822'), tint_view=True),
    # 纯黑底配 Monokai 原色，色彩按代码里的角色使用：粉色是结构标记（活动标签顶线、路径分隔符），
    # 绿色是当前名称（活动标签、当前文件夹），青色是焦点、链接与分组，橙色是匹配高亮，深蓝是选区；
    # 黑底上窗口边框（Windows 11）和标签之间的分隔线用橙色，浅色或深色背景前都看得清；
    # 输入框、按钮、勾选框和样式表里的浅灰边框统一为 Monokai 注释色，看得见又不抢眼。
    # 大面积文字保持中性白/灰，彩色只落在小元素和细线上。原生视图保持系统深色底，不叠色。
    ColorTheme('monokai_black', 'Monokai Black', 'dark', {
        'editor-bg': '#000000', 'fg': '#F8F8F2', 'fg-muted': '#A6A6A0', 'focus-border': '#66D9EF',
        'chrome-bg': '#000000', 'chrome-border': '#242524', 'window-border': '#FD971F', 'link-fg': '#66D9EF',
        'tab-active-fg': '#A6E22E', 'tab-active-border-top': '#F92672', 'tab-separator': '#FD971F',
        'tab-hover-bg': '#161616', 'crumb-active-fg': '#A6E22E', 'crumb-separator-fg': '#F92672',
        'statusbar-fg': '#CFCFC2', 'statusbar-emphasis-fg': '#66D9EF',
        'input-bg': '#161616', 'input-border': _MONOKAI_COMMENT, 'input-placeholder': '#8F8F89',
        'menu-bg': '#121212', 'menu-border': '#333333', 'menu-selection-bg': '#203561',
        'menu-selection-fg': '#FFFFFF', 'menu-separator': '#2E2E2E', 'list-hover-bg': '#FFFFFF14',
        'list-active-bg': '#203561', 'list-active-fg': '#FFFFFF', 'list-highlight-fg': '#FD971F',
        'list-active-highlight-fg': '#FD971F', 'widget-bg': '#121212', 'widget-border': '#333333',
        'picker-group-fg': '#66D9EF', 'tooltip-bg': '#121212', 'tooltip-border': '#3A3A3A',
        'palette-window': '#121212', 'palette-base': '#000000', 'palette-alt-base': '#0D0D0D',
        'palette-button': '#1E1E1E', 'selection-bg': '#203561', 'selection-fg': '#FFFFFF',
    }, remap=_dark_remap('#000000', accents=_MONOKAI_ACCENTS, contrast_line=_MONOKAI_COMMENT)),
    ColorTheme('solarized_dark', 'Solarized Dark', 'dark', {
        'editor-bg': '#002B36', 'fg': '#93A1A1', 'fg-muted': '#6F8A90', 'focus-border': '#2AA198',
        'chrome-bg': '#00212B', 'chrome-border': '#0A3F4C', 'titlebar-bg': '#002C39', 'link-fg': '#268BD2',
        'tab-bar-bg': '#004052', 'tab-inactive-bg': '#004052', 'tab-active-fg': '#EEE8D5',
        'tab-inactive-fg': '#93A1A1', 'tab-border': '#003847', 'tab-hover-bg': '#003847',
        'statusbar-bg': '#00212B', 'statusbar-border': '#0A3F4C', 'input-bg': '#003847',
        'input-border': '#0A4B5A', 'input-placeholder': '#6F8A90', 'button-bg': '#197271',
        'button-hover-bg': '#1F8A87', 'menu-bg': '#00212B', 'menu-border': '#0A3F4C',
        'menu-selection-bg': '#005A6F', 'menu-selection-fg': '#EEE8D5', 'menu-separator': '#0A3F4C',
        'list-hover-bg': '#004454', 'list-active-bg': '#005A6F', 'list-active-fg': '#EEE8D5',
        'list-highlight-fg': '#1EBCC5', 'widget-bg': '#00212B', 'widget-border': '#0A3F4C',
        'picker-group-fg': '#2AA198', 'palette-window': '#002B36', 'palette-base': '#00212B',
        'palette-alt-base': '#073642', 'palette-button': '#003847', 'selection-bg': '#005A6F',
        'selection-fg': '#EEE8D5',
    }, remap=_dark_remap('#002B36'), tint_view=True),
    ColorTheme('one_dark_pro', 'One Dark Pro', 'dark', {
        'editor-bg': '#282C34', 'fg': '#ABB2BF', 'fg-muted': '#7F848E', 'focus-border': '#528BFF',
        'chrome-bg': '#21252B', 'chrome-border': '#181A1F', 'titlebar-fg': '#9DA5B4', 'link-fg': '#61AFEF',
        'tab-active-fg': '#DCDFE4', 'tab-border': '#181A1F', 'tab-hover-bg': '#323842',
        'statusbar-fg': '#9DA5B4', 'input-bg': '#1D1F23', 'input-border': '#3E4452',
        'input-placeholder': '#5C6370', 'button-bg': '#404754', 'button-fg': '#ECECEC',
        'button-hover-bg': '#4D5565', 'menu-bg': '#21252B', 'menu-border': '#3E4452',
        'menu-selection-bg': '#2C313A', 'menu-selection-fg': '#D7DAE0', 'menu-separator': '#3E4452',
        'list-hover-bg': '#2C313A', 'list-active-bg': '#2C313A', 'list-active-fg': '#D7DAE0',
        'list-active-border': '#528BFF', 'list-highlight-fg': '#61AFEF', 'widget-bg': '#21252B',
        'widget-border': '#3E4452', 'widget-shadow': '#00000066', 'picker-group-fg': '#C678DD',
        'scrollbar-bg': '#4E566660', 'scrollbar-hover-bg': '#5A637580', 'scrollbar-active-bg': '#747D9180',
        'palette-window': '#282C34', 'palette-base': '#21252B', 'palette-alt-base': '#2C313A',
        'palette-button': '#333842', 'selection-bg': '#3E4452',
    }, remap=_dark_remap('#282C34'), tint_view=True),
    ColorTheme('github_dark', 'GitHub Dark', 'dark', {
        'editor-bg': '#0D1117', 'fg': '#E6EDF3', 'fg-muted': '#7D8590', 'focus-border': '#1F6FEB',
        'chrome-bg': '#010409', 'chrome-border': '#30363D', 'link-fg': '#2F81F7',
        'tab-active-border-top': '#F78166', 'statusbar-bg': '#0D1117', 'statusbar-fg': '#7D8590',
        'input-bg': '#0D1117', 'input-border': '#30363D', 'input-placeholder': '#6E7681',
        'button-bg': '#238636', 'button-hover-bg': '#2EA043', 'menu-bg': '#161B22', 'menu-border': '#30363D',
        'menu-selection-bg': '#1F6FEB', 'menu-separator': '#30363D', 'list-hover-bg': '#6E76811A',
        'list-active-bg': '#6E768166', 'list-active-fg': '#E6EDF3', 'list-highlight-fg': '#2F81F7',
        'widget-bg': '#161B22', 'widget-border': '#30363D', 'widget-shadow': '#010409',
        'scrollbar-bg': '#6E768133', 'scrollbar-hover-bg': '#6E768145', 'scrollbar-active-bg': '#6E768187',
        'palette-window': '#0D1117', 'palette-base': '#010409', 'palette-alt-base': '#161B22',
        'palette-button': '#21262D', 'selection-bg': '#1F6FEB',
    }, remap=_dark_remap('#0D1117')),
    ColorTheme('hc_dark', 'Dark High Contrast', 'dark', {
        'editor-bg': '#000000', 'fg': '#FFFFFF', 'fg-muted': '#D0D0D0', 'focus-border': '#F38518',
        'chrome-bg': '#000000', 'chrome-border': '#6FC3DF', 'link-fg': '#21A6FF', 'tab-inactive-fg': '#FFFFFF',
        'tab-border': '#6FC3DF', 'tab-unfocused-border-top': '#6FC3DF', 'tab-hover-bg': '#000000',
        'toolbar-hover-bg': '#FFFFFF1F', 'toolbar-active-bg': '#FFFFFF33', 'toolbar-checked-bg': '#F3851833',
        'cc-bg': '#000000', 'cc-border': '#6FC3DF', 'cc-hover-bg': '#FFFFFF1F', 'input-bg': '#000000',
        'input-border': '#6FC3DF', 'input-placeholder': '#B3B3B3', 'button-bg': '#000000',
        'button-hover-bg': '#000000', 'menu-bg': '#000000', 'menu-border': '#6FC3DF',
        'menu-selection-bg': '#000000', 'menu-selection-border': '#F38518', 'menu-separator': '#6FC3DF',
        'list-hover-bg': '#FFFFFF1A', 'list-active-bg': '#000000', 'list-active-fg': '#FFFFFF',
        'list-active-border': '#F38518', 'list-highlight-fg': '#F38518', 'widget-bg': '#0C141F',
        'widget-border': '#6FC3DF', 'widget-shadow': '#00000000', 'keybinding-bg': '#000000',
        'keybinding-border': '#6FC3DF', 'keybinding-bottom-border': '#6FC3DF', 'scrollbar-bg': '#6FC3DF99',
        'scrollbar-hover-bg': '#6FC3DFCC', 'scrollbar-active-bg': '#6FC3DF', 'palette-window': '#000000',
        'palette-base': '#000000', 'palette-alt-base': '#0C141F', 'palette-button': '#000000',
        'selection-bg': '#F38518', 'selection-fg': '#000000',
    }, remap=_dark_remap('#000000', bg_gain=0.6, bg_cap=0.18, fg_top=1.0, fg_gain=0.35, contrast_line='#6FC3DF'),
        high_contrast=True),
    ColorTheme('hc_light', 'Light High Contrast', 'light', {
        'editor-bg': '#FFFFFF', 'fg': '#292929', 'fg-muted': '#3D3D3D', 'focus-border': '#006BBD',
        'chrome-bg': '#FFFFFF', 'chrome-border': '#0F4A85', 'link-fg': '#0F4A85', 'tab-inactive-fg': '#292929',
        'tab-active-border-top': '#B5200D', 'tab-unfocused-border-top': '#0F4A85', 'tab-hover-bg': '#FFFFFF',
        'toolbar-hover-bg': '#0F4A851A', 'toolbar-active-bg': '#0F4A8533', 'toolbar-checked-bg': '#0F4A851A',
        'cc-bg': '#FFFFFF', 'cc-border': '#0F4A85', 'cc-hover-bg': '#0F4A851A', 'input-border': '#0F4A85',
        'input-placeholder': '#6D6D6D', 'button-bg': '#0F4A85', 'button-hover-bg': '#0F4A85',
        'menu-bg': '#FFFFFF', 'menu-border': '#0F4A85', 'menu-selection-bg': '#0F4A85',
        'menu-separator': '#0F4A85', 'list-hover-bg': '#0F4A851A', 'list-active-bg': '#FFFFFF',
        'list-active-fg': '#292929', 'list-active-border': '#006BBD', 'list-highlight-fg': '#0F4A85',
        'widget-bg': '#FFFFFF', 'widget-border': '#0F4A85', 'widget-shadow': '#00000000',
        'keybinding-bg': '#FFFFFF', 'keybinding-border': '#0F4A85', 'keybinding-bottom-border': '#0F4A85',
        'scrollbar-bg': '#0F4A8566', 'scrollbar-hover-bg': '#0F4A8599', 'scrollbar-active-bg': '#0F4A85',
        'selection-bg': '#0F4A85', 'selection-fg': '#FFFFFF', 'palette-alt-base': '#F2F2F2',
        'palette-button': '#FFFFFF',
    }, remap=_light_remap('#FFFFFF', compress=0.35, fg_scale=0.45, contrast_line='#0F4A85'), high_contrast=True),
)}

THEME_MODES = ('system',) + tuple(_THEMES)

_current = _THEMES['light']
_applied = False
_dark = False
_gray = False  # 灰色是较浅的深色变体，_dark 同为 True
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


def resolve_theme(mode):
    """配置值 -> 实际主题 id；跟随系统时按 Windows 应用模式取 Light Modern 或 Dark Modern。"""
    mode = normalize_mode(mode)
    if mode == 'system':
        return 'dark' if system_prefers_dark() else 'light'
    return mode


def resolve_dark(mode):
    return _THEMES[resolve_theme(mode)].dark


def resolve_gray(mode):
    return normalize_mode(mode) == 'gray'


def current_theme_id():
    return _current.id


def theme_label(theme_id):
    if theme_id == 'system':
        return tr('跟随系统')
    if theme_id == 'gray':
        return tr('灰色')
    theme = _THEMES.get(theme_id)
    return theme.label if theme is not None else theme_id


def theme_group_label(group):
    return {'light': tr('浅色主题'), 'dark': tr('深色主题'), 'hc': tr('高对比度主题')}.get(group, '')


def theme_choices():
    """[(配置值, 显示名, 分组)]：跟随系统在前，其余按浅色、深色、高对比度排列。"""
    return [('system', theme_label('system'), 'system')] + [
        (theme.id, theme_label(theme.id), theme.group) for theme in _THEMES.values()]


def is_dark():
    return _dark


def is_gray():
    return _dark and _gray


def is_high_contrast():
    return _current.high_contrast


def icon_color():
    """当前主题线条图标的颜色；Light Modern/Light+ 返回 None 表示使用图标原色。"""
    return _current.colors.get('icon-fg')


def color(token):
    """当前主题令牌的样式表颜色文本；未定义时返回空字符串。"""
    value = _current.colors.get(token)
    if not value:
        return ''
    red, green, blue, alpha = _rgba(value)
    return f'#{red:02x}{green:02x}{blue:02x}' if alpha == 255 else f'rgba({red}, {green}, {blue}, {alpha})'


def qcolor(token):
    """当前主题令牌的 QColor（含透明度）；未定义时返回无效颜色。"""
    value = _current.colors.get(token)
    return QColor(*_rgba(value)) if value else QColor()


def view_tint():
    """叠在原生资源管理器视图上的主题色调；不需要时返回 None。"""
    tint = _current.view_tint
    return QColor(*tint) if tint else None


def _shift(light, offset, amount=1.0):
    return tuple(light * 255 + delta * amount for delta in offset)


def _from_hsl(hue, sat, light):
    result = QColor.fromHslF(max(hue, 0.0), max(0.0, min(1.0, sat)), max(0.0, min(1.0, light)))
    return result.red(), result.green(), result.blue()


def _nearest_accent(hue, accents):
    """强调色板中色相最近的颜色 (r, g, b)：红色提示仍是红，绿色成功仍是绿，只是换成主题的那一种。"""
    def distance(accent):
        gap = abs(hue - QColor(accent).hslHueF()) % 1.0
        return min(gap, 1.0 - gap)
    return _rgba(min(accents, key=distance))[:3]


def _convert_dark(role, hue, sat, light, gray, remap):
    if role == 'fg':
        # 浅色文字本就用于深色/彩色底（如强调按钮上的白字），保持不变
        if light >= (0.8 if gray else 0.75):
            return None
        if gray:
            return _shift(remap['fg_top'] - light * remap['fg_gain'], remap['offset'], remap['fg_tint'])
        if remap['accents']:
            return _nearest_accent(hue, remap['accents'])
        return _from_hsl(hue, sat, max(0.62, remap['fg_top'] - light * remap['fg_gain']))
    if role == 'bg':
        # 深色或强调色背景（L<0.55）保持不变，浅色背景按亮度反转到深色区间
        if light < 0.55:
            return None
        if gray:
            return _shift(min(remap['bg_cap'], remap['base'] + (1.0 - light) * remap['bg_gain']), remap['offset'])
        base = max(remap['base'], 0.118)
        return _from_hsl(hue, sat * 0.45, min(max(remap['bg_cap'], 0.45), base + (1.0 - light) * remap['bg_gain']))
    if light < 0.55 or (not gray and light < 0.75):
        return None
    if gray and remap['contrast_line']:
        return _rgba(remap['contrast_line'])[:3]
    light = min(remap['line_cap'], max(remap['line_floor'], remap['base'] + (1.0 - light) * remap['line_gain']))
    return _shift(light, remap['offset']) if gray else _from_hsl(hue, sat * 0.5, light)


def _convert_light(role, hue, sat, light, gray, remap):
    if role == 'fg':
        if light >= 0.8:
            return None
        light *= remap['fg_scale']
        return _shift(light, remap['offset'], remap['fg_tint']) if gray else _from_hsl(hue, sat, light)
    if light < 0.55:
        return None
    if role == 'line' and gray and remap['contrast_line']:
        return _rgba(remap['contrast_line'])[:3]
    light = (1.0 - (1.0 - light) * remap['compress']) * remap['scale']
    return _shift(light, remap['offset']) if gray else _from_hsl(hue, sat, light)


@lru_cache(maxsize=4096)
def _convert(role, value, theme_id):
    remap = _THEMES[theme_id].remap
    source = QColor(value)
    if remap is None or not source.isValid():
        return value
    hue, sat, light, _alpha_value = source.getHslF()
    gray = hue < 0 or sat < 0.2
    convert = _convert_dark if remap['mode'] == 'dark' else _convert_light
    result = convert(role, hue, sat, light, gray, remap)
    if result is None:
        return value
    red, green, blue = (max(0, min(255, int(round(channel)))) for channel in result)
    if source.alpha() == 255:
        return f'#{red:02x}{green:02x}{blue:02x}'
    return f'rgba({red}, {green}, {blue}, {source.alpha()})'


def fg(value):
    """文字/线条前景色。"""
    return _convert('fg', value, _current.id)


def bg(value):
    """背景填充色。"""
    return _convert('bg', value, _current.id)


def line(value):
    """边框/分隔线颜色。"""
    return _convert('line', value, _current.id)


def status_color(value):
    """状态栏里的彩色文字：强调色底的状态栏（Light+/Dark+）统一用状态栏前景色，否则按主题换算。"""
    return color('statusbar-fg') if _current.prominent_statusbar else fg(value)


def group_tab_fill_alpha(group_color, active):
    """标签分组底色的不透明度（0–255）。

    浅色主题保持 72/52。深色底上同样的透明度只剩一层暗红/暗蓝，看不出是同一组：在不低于原值的前提下逐步加深，
    直到标签文字对比度将低于 4.5（当前标签）/ 3.0（其他标签）或达到上限。"""
    base = 72 if active else 52
    if not _dark:
        return base
    tokens = _current.colors
    text = tokens['tab-active-fg' if active else 'tab-inactive-fg'][:7]
    under = tokens['tab-active-bg' if active else 'tab-inactive-bg'][:7]
    minimum, ceiling = (4.5, 112) if active else (3.0, 84)
    best = base
    for alpha in range(base + 4, ceiling + 1, 4):
        if _contrast(text, _mix(under, group_color, alpha / 255.0)) < minimum:
            break
        best = alpha
    return best


_DECLARATION = re.compile(r'([A-Za-z-]+)(\s*:\s*)([^;{}]*)')
_COLOR = re.compile(r'#[0-9A-Fa-f]{3,8}\b|\b(?:white|black)\b')
_TOKEN = re.compile(r'@([a-z][a-z0-9-]*)')

# 主窗口及其子窗口共用的菜单、提示、滚动条样式（VS Code 风格）
GLOBAL_CSS = """
QMenu { background-color: @menu-bg; color: @menu-fg; border: 1px solid @menu-border; padding: 4px 0px; }
QMenu::item { background: transparent; color: @menu-fg; padding: 5px 24px 5px 12px; margin: 0px 4px;
              border: 1px solid transparent; border-radius: 4px; }
QMenu::item:selected { background: @menu-selection-bg; color: @menu-selection-fg; border-color: @menu-selection-border; }
QMenu::item:disabled { color: @fg-disabled; background: transparent; border-color: transparent; }
QMenu::separator { height: 1px; background: @menu-separator; margin: 4px 10px; }
QToolTip { background-color: @tooltip-bg; color: @tooltip-fg; border: 1px solid @tooltip-border; padding: 4px 6px; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 0px; border: none; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 0px; border: none; }
QScrollBar::handle:vertical { background: @scrollbar-bg; min-height: 24px; border: none; }
QScrollBar::handle:horizontal { background: @scrollbar-bg; min-width: 24px; border: none; }
QScrollBar::handle:hover { background: @scrollbar-hover-bg; }
QScrollBar::handle:pressed { background: @scrollbar-active-bg; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0px; height: 0px; border: none; background: none; }
QScrollBar::add-page, QScrollBar::sub-page { background: none; }
"""


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
    """把按浅色主题书写的样式表换算为当前主题：先换算字面颜色，再填入 @令牌。

    selection-background-color 等强调色保持不变。"""
    if not css:
        return css
    theme_id = _current.id
    if _current.remap is not None:
        def convert(match):
            role = _role_for(match.group(1))
            if role is None:
                return match.group(0)
            value = _COLOR.sub(lambda found: _convert(role, found.group(0), theme_id), match.group(3))
            return match.group(1) + match.group(2) + value

        css = _DECLARATION.sub(convert, css)
    if '@' in css:
        css = _TOKEN.sub(lambda found: color(found.group(1)) or found.group(0), css)
    return css


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
            converted = qss(css)
            if widget.styleSheet() != converted:
                widget.setStyleSheet(converted)
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


def _colorref(value):
    red, green, blue, _alpha_value = _rgba(value)
    return red | (green << 8) | (blue << 16)


def _dwm_frame_colors():
    """Windows 11 系统窗口框架的颜色 [(DWM 属性, COLORREF)]：标题栏底色、标题文字色和窗口边框色。

    主题未指定 window-border 时边框恢复系统默认（DWMWA_COLOR_DEFAULT），从带边框色的主题切走后不残留。"""
    tokens = _current.colors
    border = tokens.get('window-border')
    return [(35, _colorref(tokens['titlebar-bg'])),  # DWMWA_CAPTION_COLOR
            (36, _colorref(tokens['titlebar-fg'])),  # DWMWA_TEXT_COLOR
            (34, _colorref(border) if border else 0xFFFFFFFF)]  # DWMWA_BORDER_COLOR


def _set_dwm_caption(hwnd, colors):
    """Windows 11 起可设置标题栏底色、文字色和边框色，让系统标题栏与下方工具栏连成一体。"""
    if _windows_build() < 22000:
        return False
    for attribute, value in colors:
        colorref = ctypes.c_uint32(value)
        try:
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    ctypes.wintypes.HWND(hwnd), attribute, ctypes.byref(colorref), ctypes.sizeof(colorref)) != 0:
                return False
        except (AttributeError, OSError):
            return False
    return True


def apply_window_frame(widget, force=False):
    """按当前主题设置顶层窗口系统标题栏：深浅色，Windows 11 上同时设置标题栏底色、文字色与窗口边框色。"""
    try:
        if not widget.isWindow() or widget.windowFlags() & Qt.FramelessWindowHint:
            return
        if widget.windowType() in (Qt.Popup, Qt.ToolTip, Qt.SplashScreen, Qt.Desktop):
            return
        dark_stale = force or bool(widget.property(_FRAME_PROPERTY)) != _dark
        caption_stale = force or widget.property(_CAPTION_PROPERTY) != _current.id
        if not (dark_stale or caption_stale):
            return
        hwnd = int(widget.effectiveWinId() or 0)
        if not hwnd:
            return
        changed = False
        if dark_stale and _set_dwm_dark(hwnd, _dark):
            widget.setProperty(_FRAME_PROPERTY, _dark)
            changed = True
        if caption_stale and _set_dwm_caption(hwnd, _dwm_frame_colors()):
            widget.setProperty(_CAPTION_PROPERTY, _current.id)
            changed = True
        if changed and widget.isVisible():
            # RDW_INVALIDATE | RDW_FRAME：已显示的窗口立即重绘标题栏
            ctypes.windll.user32.RedrawWindow(ctypes.wintypes.HWND(hwnd), None, None, 0x0001 | 0x0400)
    except RuntimeError:
        pass


def _build_palette(theme):
    """非系统风格主题的 Qt 调色板（Fusion 风格的按钮、列表、输入框等标准控件使用）。"""
    tokens = theme.colors
    button = tokens['palette-button']
    if theme.dark:
        bevels = (_mix(button, '#FFFFFF', 0.1), _mix(button, '#FFFFFF', 0.05),
                  _mix(button, '#000000', 0.1), _mix(button, '#000000', 0.45))
    else:
        bevels = ('#FFFFFF', _mix(button, '#FFFFFF', 0.5), _mix(button, '#000000', 0.25), _mix(button, '#000000', 0.4))
    roles = {
        QPalette.Window: tokens['palette-window'], QPalette.WindowText: tokens['fg'],
        QPalette.Base: tokens['palette-base'], QPalette.AlternateBase: tokens['palette-alt-base'],
        QPalette.ToolTipBase: tokens['tooltip-bg'], QPalette.ToolTipText: tokens['tooltip-fg'],
        QPalette.PlaceholderText: tokens['input-placeholder'], QPalette.Text: tokens['fg'],
        QPalette.Button: button, QPalette.ButtonText: tokens['fg'], QPalette.BrightText: '#FF6B6B',
        QPalette.Light: bevels[0], QPalette.Midlight: bevels[1], QPalette.Mid: bevels[2], QPalette.Dark: bevels[3],
        QPalette.Shadow: '#000000', QPalette.Highlight: tokens['selection-bg'],
        QPalette.HighlightedText: tokens['selection-fg'], QPalette.Link: tokens['link-fg'],
        QPalette.LinkVisited: '#B39DDB' if theme.dark else '#7B1FA2',
    }
    palette = QPalette()
    for role, value in roles.items():
        palette.setColor(role, QColor(*_rgba(value)))
    for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
        palette.setColor(QPalette.Disabled, role, QColor(*_rgba(tokens['fg-disabled'])))
    palette.setColor(QPalette.Disabled, QPalette.Highlight,
                     QColor(*_rgba(_mix(tokens['selection-bg'], tokens['palette-base'], 0.6))))
    palette.setColor(QPalette.Disabled, QPalette.HighlightedText,
                     QColor(*_rgba(_mix(tokens['selection-fg'], tokens['palette-base'], 0.3))))
    return palette


def _control_outline(option, focused):
    """输入框、按钮、勾选框的外框色：获得焦点（或默认按钮）用 focus-border，禁用时 control-border 变淡。"""
    enabled = bool(option.state & QStyle.State_Enabled)
    if focused and enabled:
        return qcolor('focus-border')
    color = qcolor('control-border')
    if not enabled:
        color.setAlphaF(color.alphaF() * 0.45)
    return color


def _stroke(painter, color, draw, dx=0.5, dy=0.5):
    """按 Fusion 的几何（半像素平移、抗锯齿、1px）在它画的外框上重描，正好盖住原来的颜色。"""
    painter.save()
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.translate(dx, dy)
    painter.setPen(color)
    painter.setBrush(Qt.NoBrush)
    draw(painter)
    painter.restore()


class _ControlBorderStyle(QProxyStyle):
    """非系统风格主题使用的 Fusion。

    Fusion 的输入框、按钮、勾选框外框由窗口底色加深得到，深色底上几乎看不见。主题定义了 control-border 时，
    在 Fusion 绘制后按相同几何重描外框：常态用 control-border，获得焦点的输入框、键盘焦点所在的按钮/勾选框
    以及默认按钮用 focus-border，禁用时变淡。样式表自己写了边框的控件不经过这里。"""

    def __init__(self):
        super().__init__('Fusion')
        self.setObjectName('fusion')

    def drawPrimitive(self, element, option, painter, widget=None):
        super().drawPrimitive(element, option, painter, widget)
        if not _current.colors.get('control-border'):
            return
        state = option.state
        has_focus = bool(state & QStyle.State_HasFocus)
        keyboard_focus = has_focus and bool(state & QStyle.State_KeyboardFocusChange)
        rect = QRectF(option.rect)
        if element == QStyle.PE_FrameLineEdit:
            _stroke(painter, _control_outline(option, has_focus),
                    lambda p: p.drawRoundedRect(rect.adjusted(0, 0, -1, -1), 2, 2))
        elif element == QStyle.PE_PanelButtonCommand:
            features = int(option.features) if isinstance(option, QStyleOptionButton) else 0
            if features & int(QStyleOptionButton.Flat) and not state & (QStyle.State_Sunken | QStyle.State_On):
                return  # Fusion 不给平面按钮画底框
            default = bool(features & int(QStyleOptionButton.DefaultButton))
            _stroke(painter, _control_outline(option, keyboard_focus or default),
                    lambda p: p.drawRoundedRect(rect.adjusted(0, 1, -1, 0), 2, 2), dy=-0.5)
        elif element == QStyle.PE_IndicatorCheckBox:
            _stroke(painter, _control_outline(option, keyboard_focus),
                    lambda p: p.drawRect(rect.adjusted(0, 0, -1, -1)))
        elif element == QStyle.PE_IndicatorRadioButton:
            width = option.rect.width()
            center = QPointF(option.rect.center()) + QPointF(1, 1)
            radius = (width + (width + 1) % 2) / 2.0 - 1
            _stroke(painter, _control_outline(option, keyboard_focus),
                    lambda p: p.drawEllipse(center, radius, radius), dx=0, dy=0)
        elif element == QStyle.PE_Frame and isinstance(widget, (QTextEdit, QPlainTextEdit)):
            _stroke(painter, _control_outline(option, widget.hasFocus()),
                    lambda p: p.drawRect(rect.adjusted(0, 0, -1, -1)))

    def drawComplexControl(self, control, option, painter, widget=None):
        super().drawComplexControl(control, option, painter, widget)
        # 下拉框由 Fusion 内部调用上面的 PE_FrameLineEdit / PE_PanelButtonCommand，这里只需补数字框
        if control == QStyle.CC_SpinBox and getattr(option, 'frame', True) and _current.colors.get('control-border'):
            rect = QRectF(option.rect)
            _stroke(painter, _control_outline(option, bool(option.state & QStyle.State_HasFocus)),
                    lambda p: p.drawRoundedRect(rect.adjusted(0, 0, -1, -1), 2, 2))


def set_theme(theme_id):
    """切换到指定主题（可传配置值，system 按 Windows 设置解析）；返回是否发生变化。

    已存在的 Shell 视图只在深浅色变化时需要由调用方重建。"""
    global _current, _applied, _dark, _gray, _original_style, _original_palette
    theme = _THEMES.get(resolve_theme(theme_id), _THEMES['light'])
    if theme is _current and _applied:
        return False
    app = QApplication.instance()
    if _original_style is None:
        _original_style = app.style().objectName() or 'windowsvista'
        _original_palette = QPalette(app.palette())
    native_changed = theme.dark != _dark or not _applied
    _current, _applied, _dark, _gray = theme, True, theme.dark, theme.id == 'gray'
    if native_changed:
        _set_native_app_mode(theme.dark)
    if theme.native_style:
        if app.style().objectName().lower() != _original_style.lower():
            app.setStyle(_original_style)
    elif not isinstance(app.style(), _ControlBorderStyle):
        app.setStyle(_ControlBorderStyle())
    palette = QPalette(_original_palette) if theme.native_style else _build_palette(theme)
    app.setPalette(palette)
    QToolTip.setPalette(palette)
    restyle_widgets()
    for widget in app.topLevelWidgets():
        if widget.isVisible():
            apply_window_frame(widget)
    debug_print(f"[Theme] switched to {theme.id}")
    return True


def set_dark(dark, gray=False):
    """兼容旧接口：深色为 Dark Modern，gray=True 为灰色，浅色为 Light Modern。"""
    return set_theme('gray' if dark and gray else 'dark' if dark else 'light')

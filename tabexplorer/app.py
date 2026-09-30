"""程序启动与单实例转发。"""

import socket
import time

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication

from . import debuglog as _debuglog
from .debuglog import debug_print, _init_debug_log_file, qt_message_handler, _start_diagnostics
from .widgets import _build_te_icon
from .shellview import _ieb_keyboard_filter
from .mainwindow import MainWindow


def try_send_to_existing_instance(path):
    """尝试将路径发送给已运行的实例"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client.settimeout(2.0)  # 增加超时时间
            client.connect(('127.0.0.1', 58923))
            client.send(path.encode('utf-8'))
            client.close()
            debug_print(f"[Client] Successfully sent path to existing instance: {path}")
            return True
        except Exception as e:
            debug_print(f"[Client] Attempt {attempt + 1}/{max_retries} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(0.1)  # 短暂等待后重试
            continue
    debug_print("[Client] No existing instance found, starting new instance")
    return False


def main():
    # 支持命令行参数：打开指定路径
    import sys
    import os

    # TortoiseOverlays IsMemberOf vtable patch is applied at module load time
    # (see early init block near _patch_tortoise_overlays). No registry changes needed.

    path_to_open = None
    if len(sys.argv) > 1:
        path = sys.argv[1]
        # 处理可能的引号
        path = path.strip('"').strip("'")
        if os.path.exists(path):
            # 如果是文件，打开其所在目录
            if os.path.isfile(path):
                path = os.path.dirname(path)
            path_to_open = path
            
            # 尝试发送给已运行的实例
            if try_send_to_existing_instance(path):
                print(f"Sent path to existing instance: {path}")
                sys.exit(0)  # 退出程序，不启动新实例
    
    # 禁用 Qt 的警告输出（在创建 QApplication 之前设置）
    os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.qpa.*=false'
    
    # 启用高DPI支持（在创建QApplication之前）
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    _start_diagnostics()
    _init_debug_log_file()
    from PyQt5.QtCore import qInstallMessageHandler
    qInstallMessageHandler(qt_message_handler)
    
    # 启动新实例
    app = QApplication(sys.argv)
    app.setApplicationName("TabExplorer")
    
    # 获取屏幕DPI和缩放因子
    screen = app.primaryScreen()
    dpi = screen.logicalDotsPerInch()
    scale_factor = dpi / 96.0  # 96是标准DPI
    debug_print(f"[DPI] Screen DPI: {dpi}, Scale Factor: {scale_factor:.2f}")
    
    # 根据DPI动态调整全局样式
    base_font_size = 9  # 基础字体大小 (pt)
    scaled_font_size = int(base_font_size * scale_factor)
    
    # 设置全局字体
    from PyQt5.QtGui import QFont
    app_font = QFont("Microsoft YaHei UI", scaled_font_size)
    app.setFont(app_font)
    debug_print(f"[DPI] Global font size: {scaled_font_size}pt")
    
    # 安装 IExplorerBrowser 键盘消息过滤器
    app.installNativeEventFilter(_ieb_keyboard_filter)
    
    # 图标将在 MainWindow.__init__ 中生成并设置，确保 Qt 完全初始化后执行
    
    # 应用级（任务栏）窗口图标：必须在创建主窗口及其书签菜单之前设置。
    # 此时尚无任何弹出型 QMenu 顶层部件，图标变更事件只广播到极少数顶层部件，
    # 从而避免英文界面下书签菜单栏溢出时，图标变更事件在弹出菜单之间无限递归
    # 传播（导致栈溢出崩溃）。
    try:
        app.setWindowIcon(_build_te_icon())
    except Exception as _icon_e:
        debug_print(f"[Icon] app icon early-set failed: {_icon_e}")

    # 创建窗口（图标在 MainWindow.__init__ 内部生成）
    window = MainWindow()
    window._diagnostic_timer = QTimer(window)
    window._diagnostic_timer.setInterval(10000)
    window._diagnostic_timer.timeout.connect(window._capture_diagnostic_snapshot)
    window._diagnostic_timer.start()
    window._capture_diagnostic_snapshot()
    
    
    # 如果有路径参数，在新窗口中打开
    # 注意：固定标签页在延迟初始化阶段加载，这里要避免和固定标签重复
    if path_to_open:
        pinned_norm = {
            window._normalize_path_for_compare(p)
            for p in window.config.get("pinned_tabs", []) if p
        }
        if window._normalize_path_for_compare(path_to_open) in pinned_norm:
            debug_print(f"[App] Skip argv path (already pinned): {path_to_open}")
        else:
            window.add_new_tab(path_to_open)
    
    # 启动时最大化显示
    window.showMaximized()
    

    # No HKCU cleanup needed – overlay fix is vtable-only (process-local).

    exit_code = app.exec_()
    if _debuglog._diagnostics is not None:
        _debuglog._diagnostics.close()
    sys.exit(exit_code)

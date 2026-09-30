"""程序启动与单实例转发。"""

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication

from . import debuglog as _debuglog
from .debuglog import debug_print, _init_debug_log_file, qt_message_handler, _start_diagnostics
from .paths import get_app_base_dir, get_app_data_path
from .instance import InstanceCoordinator


def _configure_standard_streams():
    import sys
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, 'reconfigure'):
            try:
                stream.reconfigure(encoding='utf-8', errors='backslashreplace')
            except (OSError, ValueError):
                pass


def try_send_to_existing_instance(path, coordinator=None):
    coordinator = coordinator or InstanceCoordinator(get_app_data_path())
    return coordinator.send(path or '')


def main():
    # 支持命令行参数：打开指定路径
    import sys
    import os
    import time
    started_at = time.time()

    # TortoiseOverlays IsMemberOf vtable patch is applied at module load time
    # (see early init block near _patch_tortoise_overlays). No registry changes needed.

    validation = len(sys.argv) > 1 and sys.argv[1] == '--self-test'
    if validation and (not os.environ.get('TABEX_DATA_DIR')
                       or not os.path.isfile(get_app_data_path('.tabex-validation'))):
        raise RuntimeError('Self-test requires an isolated validation data directory')
    path_to_open = sys.argv[1].strip('"\'') if len(sys.argv) > 1 and not validation else ''
    coordinator = InstanceCoordinator(get_app_data_path())
    if not coordinator.acquire():
        if try_send_to_existing_instance(path_to_open, coordinator):
            sys.exit(0)
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, 'TabExplorer is already starting or not responding. Please retry.',
                                        'TabExplorer', 0x10)
        sys.exit(1)
    import atexit
    atexit.register(coordinator.close)
    from .widgets import _build_te_icon
    from .shellview import _ieb_keyboard_filter
    from .mainwindow import MainWindow
    _configure_standard_streams()
    
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
    window._instance_coordinator = coordinator
    coordinator.set_receiver(window.open_path_signal.emit)
    window._diagnostic_timer = QTimer(window)
    window._diagnostic_timer.setInterval(10000)
    window._diagnostic_timer.timeout.connect(window._capture_diagnostic_snapshot)
    window._diagnostic_timer.start()
    window._capture_diagnostic_snapshot()
    
    
    # 如果有路径参数，在新窗口中打开
    # 注意：固定标签页在延迟初始化阶段加载，这里要避免和固定标签重复
    if path_to_open:
        pinned_norm = {window._normalize_path_for_compare(path) for path in window._get_pinned_paths_from_config()}
        if window._normalize_path_for_compare(path_to_open) in pinned_norm:
            debug_print(f"[App] Skip argv path (already pinned): {path_to_open}")
        else:
            window.handle_open_path_from_instance(path_to_open)
    
    # 启动时最大化显示
    window.showMaximized()
    if validation:
        from .runtime_validation import RuntimeValidation
        window._runtime_validation = RuntimeValidation(window, started_at)
    

    # No HKCU cleanup needed – overlay fix is vtable-only (process-local).

    exit_code = app.exec_()
    coordinator.close()
    app.removeNativeEventFilter(_ieb_keyboard_filter)
    from PyQt5.QtCore import QCoreApplication, QEvent
    window.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
    if _debuglog._diagnostics is not None:
        _debuglog._diagnostics.close()
    sys.exit(exit_code)

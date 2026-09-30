"""调试日志、崩溃诊断接入与 Qt 消息处理。"""

import io
import os
import threading
import time

from .paths import get_app_data_path
from .constants import APP_VERSION


# 全局调试开关
_DEBUG_MODE = False  # 生产环境关闭，避免性能损耗
_EXPLORER_MONITOR_DEBUG = False  # Explorer Monitor 单独的日志开关
_DEBUG_LOG_PATH = get_app_data_path("TabEx_debug_latest.log")
_DEBUG_LOG_LOCK = threading.Lock()
_DEBUG_LOG_READY = False
_diagnostics = None


def _start_diagnostics():
    global _diagnostics
    from .diagnostics import Diagnostics
    import tempfile
    for base in (os.environ.get('LOCALAPPDATA'), tempfile.gettempdir()):
        if not base:
            continue
        try:
            _diagnostics = Diagnostics(os.path.join(base, 'TabEx', 'diagnostics'), APP_VERSION)
            _diagnostics.install()
            _diagnostics.capture_debug(_DEBUG_LOG_PATH)
            return
        except OSError:
            continue


def _diagnostic_event(kind, details):
    if _diagnostics is not None:
        _diagnostics.record(kind, details)


def _init_debug_log_file():
    """初始化调试日志文件（覆盖旧内容，仅保留本次运行日志）。"""
    global _DEBUG_LOG_READY
    try:
        with _DEBUG_LOG_LOCK:
            with open(_DEBUG_LOG_PATH, "w", encoding="utf-8", errors="replace") as f:
                f.write(time.strftime("[%Y-%m-%d %H:%M:%S] ") + "[DebugLog] session started\n")
        _DEBUG_LOG_READY = True
    except Exception:
        _DEBUG_LOG_READY = False


def _append_debug_log_line(message):
    """将一行调试信息追加到日志文件。"""
    if not _DEBUG_LOG_READY:
        return
    try:
        ts = time.strftime("[%Y-%m-%d %H:%M:%S]")
        line = f"{ts} {message}\n"
        with _DEBUG_LOG_LOCK:
            with open(_DEBUG_LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
                f.write(line)
    except Exception:
        pass


def debug_print(*args, **kwargs):
    """根据调试开关输出调试信息，并写入本地日志文件。"""
    if _diagnostics is not None and args and isinstance(args[0], str):
        if args[0].startswith(('[ClosedTabs]', '[TabSwitch]', '[AsyncLoad]', '[PidlResolver]')):
            _diagnostic_event('navigation', ' '.join(str(value) for value in args))
    should_emit = False
    if _DEBUG_MODE:
        # 检查是否是 Explorer Monitor 日志
        if args and isinstance(args[0], str) and '[Explorer Monitor]' in args[0]:
            if _EXPLORER_MONITOR_DEBUG:
                should_emit = True
        else:
            should_emit = True

    if not should_emit:
        return

    print(*args, **kwargs)

    try:
        local_kwargs = dict(kwargs)
        local_kwargs.pop('file', None)
        local_kwargs.pop('flush', None)
        buf = io.StringIO()
        print(*args, file=buf, **local_kwargs)
        text = buf.getvalue().rstrip('\r\n')
        if text:
            _append_debug_log_line(text)
    except Exception:
        pass


def dbg_exc(where=""):
    """在调试模式下记录当前正在处理的异常（含类型与消息），生产环境为零开销空操作。
    用于替代关键路径上原本静默吞掉异常的 `except ...: pass`，提升可诊断性而不影响行为。"""
    if not _DEBUG_MODE:
        return
    try:
        import sys as _sys
        exc = _sys.exc_info()[1]
        if exc is not None:
            debug_print(f"[swallowed]{(' ' + where) if where else ''}: {type(exc).__name__}: {exc}")
    except Exception:
        pass


def set_explorer_monitor_debug(enabled):
    """设置 Explorer Monitor 日志开关"""
    global _EXPLORER_MONITOR_DEBUG
    _EXPLORER_MONITOR_DEBUG = enabled
    debug_print(f"[Config] Explorer Monitor debug output: {'enabled' if enabled else 'disabled'}")


def set_debug_mode(enabled):
    """设置全局调试模式"""
    global _DEBUG_MODE
    _DEBUG_MODE = enabled


def qt_message_handler(mode, context, message):
    """自定义 Qt 消息处理器，过滤 QAxBase 等不需要的警告"""
    from PyQt5.QtCore import QtCriticalMsg, QtFatalMsg
    if mode in (QtCriticalMsg, QtFatalMsg) and _diagnostics is not None:
        _diagnostics.record('qt_fatal' if mode == QtFatalMsg else 'qt_critical', message)
        if mode == QtFatalMsg:
            _diagnostics.mark('fatal')
            _diagnostics.dump_threads()
    # 只在调试模式下输出 Qt 警告
    if _DEBUG_MODE:
        # 如果是调试模式，输出所有消息
        debug_print(f"Qt Message: {message}")
    else:
        # 非调试模式下，只输出严重错误（Critical 和 Fatal）
        from PyQt5.QtCore import QtCriticalMsg, QtFatalMsg
        if mode in (QtCriticalMsg, QtFatalMsg):
            debug_print(f"Qt Error: {message}")
        # 其他消息（Debug, Warning, Info）都被过滤

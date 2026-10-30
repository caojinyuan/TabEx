"""调试日志、崩溃诊断接入与 Qt 消息处理。"""

import io
import os
import shutil
import sys
import threading
import time

from .paths import get_app_data_path, get_app_log_path
from .constants import APP_VERSION, DEBUG_LOG_BACKUP_COUNT, DEBUG_LOG_MAX_BYTES


# 全局调试开关
_DEBUG_MODE = False  # 生产环境关闭，避免性能损耗
_EXPLORER_MONITOR_DEBUG = False  # Explorer Monitor 单独的日志开关
_DEBUG_LOG_PATH = get_app_log_path("TabEx_debug_latest.log")
_DEBUG_LOG_LOCK = threading.Lock()
_DEBUG_LOG_READY = False
_DEBUG_LOG_ERROR = ''
_diagnostics = None


def _start_diagnostics():
    global _diagnostics, _DEBUG_LOG_PATH
    from .diagnostics import Diagnostics
    import tempfile
    directories = [get_app_log_path()]
    legacy_directories = []
    if not os.environ.get('TABEX_DATA_DIR'):
        for base in (os.environ.get('LOCALAPPDATA'), tempfile.gettempdir()):
            if base:
                directories.append(os.path.join(base, 'TabEx', 'logs'))
                legacy_directories.append(os.path.join(base, 'TabEx', 'diagnostics'))
                legacy_directories.append(os.path.join(base, 'TabEx', 'logs', 'diagnostics'))
    _diagnostics = None
    _DEBUG_LOG_PATH = get_app_log_path("TabEx_debug_latest.log")
    for directory in dict.fromkeys(directories):
        service = None
        try:
            service = Diagnostics(os.path.join(directory, 'diagnostics'), APP_VERSION)
            service.install()
        except OSError as error:
            if service is not None:
                service.close()
            print(f"[Diagnostics] Cannot use log directory {directory}: {error}", file=sys.stderr)
            continue
        _diagnostics = service
        _DEBUG_LOG_PATH = os.path.join(directory, "TabEx_debug_latest.log")
        previous_debug = get_app_data_path("TabEx_debug_latest.log")
        if os.path.isfile(previous_debug) and not os.path.exists(_DEBUG_LOG_PATH):
            try:
                shutil.move(previous_debug, _DEBUG_LOG_PATH)
            except OSError as error:
                service.record('log_migration_failed', str(error))
                print(f"[Diagnostics] Cannot migrate previous debug log: {error}", file=sys.stderr)
        service.migrate_sessions(legacy_directories)
        service.capture_debug(_DEBUG_LOG_PATH if os.path.isfile(_DEBUG_LOG_PATH) else previous_debug)
        return
    print("[Diagnostics] No writable log directory; diagnostics unavailable.", file=sys.stderr)


def _diagnostic_event(kind, details):
    if _diagnostics is not None:
        _diagnostics.record(kind, details)


def _init_debug_log_file():
    """初始化调试日志文件（覆盖旧内容，仅保留本次运行日志）。"""
    global _DEBUG_LOG_READY, _DEBUG_LOG_ERROR
    try:
        with _DEBUG_LOG_LOCK:
            os.makedirs(os.path.dirname(_DEBUG_LOG_PATH), exist_ok=True)
            with open(_DEBUG_LOG_PATH, "wb") as f:
                f.write((time.strftime("[%Y-%m-%d %H:%M:%S] ") + "[DebugLog] session started\n").encode('utf-8'))
        _DEBUG_LOG_READY = True
        _DEBUG_LOG_ERROR = ''
    except OSError as error:
        _DEBUG_LOG_READY = False
        _report_debug_log_error(f"Cannot initialize log file {_DEBUG_LOG_PATH}: {error}")


def _report_debug_log_error(error):
    global _DEBUG_LOG_ERROR
    if not _DEBUG_LOG_ERROR:
        print(f"[DebugLog] Write failed: {error}", file=sys.stderr)
        _diagnostic_event('debug_log_write_failed', str(error))
    _DEBUG_LOG_ERROR = str(error)


def _rotate_debug_log():
    for index in range(DEBUG_LOG_BACKUP_COUNT - 1, 0, -1):
        source = f'{_DEBUG_LOG_PATH}.{index}'
        if os.path.exists(source):
            os.replace(source, f'{_DEBUG_LOG_PATH}.{index + 1}')
    if os.path.exists(_DEBUG_LOG_PATH):
        os.replace(_DEBUG_LOG_PATH, f'{_DEBUG_LOG_PATH}.1')


def _append_debug_log_line(message):
    """将一行调试信息追加到日志文件。"""
    global _DEBUG_LOG_ERROR
    if not _DEBUG_LOG_READY:
        return False
    try:
        ts = time.strftime("[%Y-%m-%d %H:%M:%S]")
        line = f"{ts} {message}\n".encode('utf-8', errors='replace')
        if len(line) > DEBUG_LOG_MAX_BYTES:
            marker = b'\n[DebugLog] oversized entry truncated\n'
            line = line[:DEBUG_LOG_MAX_BYTES - len(marker)].decode('utf-8', errors='ignore').encode('utf-8') + marker
        with _DEBUG_LOG_LOCK:
            try:
                size = os.path.getsize(_DEBUG_LOG_PATH)
            except FileNotFoundError:
                size = 0
            if size + len(line) > DEBUG_LOG_MAX_BYTES:
                _rotate_debug_log()
            with open(_DEBUG_LOG_PATH, "ab") as f:
                f.write(line)
        _DEBUG_LOG_ERROR = ''
        return True
    except OSError as error:
        _report_debug_log_error(error)
        return False


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

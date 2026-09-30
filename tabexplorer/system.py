"""进程资源信息、外部程序启动与后台线程保留。"""

import os
import subprocess
import threading
import time

from .i18n import tr
from .constants import SUPPORTED_TERMINAL_TOOLS, TITLE_SHORTCUT_EXTENSIONS
from .debuglog import debug_print


def detect_notepad_plus_plus():
    """检测系统中是否安装了 Notepad++。"""
    import shutil

    exe_path = shutil.which('notepad++.exe')
    if exe_path:
        return exe_path

    common_paths = [
        r'C:\Program Files\Notepad++\notepad++.exe',
        r'C:\Program Files (x86)\Notepad++\notepad++.exe',
        os.path.expandvars(r'%PROGRAMFILES%\Notepad++\notepad++.exe'),
        os.path.expandvars(r'%PROGRAMFILES(X86)%\Notepad++\notepad++.exe'),
        os.path.expandvars(r'%LOCALAPPDATA%\Programs\Notepad++\notepad++.exe'),
    ]
    for path in common_paths:
        if os.path.exists(path):
            return path
    return None


def format_file_size(size_bytes):
    """将字节数格式化为带单位的人类可读字符串。"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    if size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


def get_process_memory_usage_mb():
    """Return current process working set in MB on Windows, else None."""
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
            _fields_ = [
                ('cb', wintypes.DWORD),
                ('PageFaultCount', wintypes.DWORD),
                ('PeakWorkingSetSize', ctypes.c_size_t),
                ('WorkingSetSize', ctypes.c_size_t),
                ('QuotaPeakPagedPoolUsage', ctypes.c_size_t),
                ('QuotaPagedPoolUsage', ctypes.c_size_t),
                ('QuotaPeakNonPagedPoolUsage', ctypes.c_size_t),
                ('QuotaNonPagedPoolUsage', ctypes.c_size_t),
                ('PagefileUsage', ctypes.c_size_t),
                ('PeakPagefileUsage', ctypes.c_size_t),
                ('PrivateUsage', ctypes.c_size_t),
            ]

        counters = PROCESS_MEMORY_COUNTERS_EX()
        counters.cb = ctypes.sizeof(counters)
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        psapi = ctypes.WinDLL('psapi', use_last_error=True)
        get_current_process = kernel32.GetCurrentProcess
        get_current_process.restype = wintypes.HANDLE
        get_process_memory_info = psapi.GetProcessMemoryInfo
        get_process_memory_info.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(PROCESS_MEMORY_COUNTERS_EX),
            wintypes.DWORD,
        ]
        get_process_memory_info.restype = wintypes.BOOL

        if get_process_memory_info(get_current_process(), ctypes.byref(counters), counters.cb):
            return round(float(counters.WorkingSetSize) / (1024 * 1024), 2)
    except Exception:
        return None
    return None


# CPU 占用率采样状态（进程 CPU 时间 / 挂钟时间 增量法，无 psutil 依赖）
_cpu_last_proc_time = None
_cpu_last_wall_time = None
_cpu_logical_count = os.cpu_count() or 1


def get_process_cpu_percent():
    """返回本进程自上次调用以来的平均 CPU 占用率（0~100，按逻辑核归一）。
    Windows 用 GetProcessTimes 取内核+用户时间增量除以挂钟增量；非 Windows 返回 None。"""
    global _cpu_last_proc_time, _cpu_last_wall_time
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD), ('dwHighDateTime', wintypes.DWORD)]

        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        kernel32.GetProcessTimes.argtypes = [wintypes.HANDLE, ctypes.POINTER(FILETIME),
                                             ctypes.POINTER(FILETIME), ctypes.POINTER(FILETIME),
                                             ctypes.POINTER(FILETIME)]
        kernel32.GetProcessTimes.restype = wintypes.BOOL
        creation = FILETIME(); exitt = FILETIME(); kern = FILETIME(); user = FILETIME()
        if not kernel32.GetProcessTimes(kernel32.GetCurrentProcess(),
                                        ctypes.byref(creation), ctypes.byref(exitt),
                                        ctypes.byref(kern), ctypes.byref(user)):
            return None
        proc_100ns = ((kern.dwHighDateTime << 32) | kern.dwLowDateTime) + \
                     ((user.dwHighDateTime << 32) | user.dwLowDateTime)
        now = time.monotonic()
        if _cpu_last_proc_time is None:
            _cpu_last_proc_time = proc_100ns
            _cpu_last_wall_time = now
            return None
        wall_delta = now - _cpu_last_wall_time
        proc_delta = (proc_100ns - _cpu_last_proc_time) / 1e7  # 100ns -> seconds
        _cpu_last_proc_time = proc_100ns
        _cpu_last_wall_time = now
        if wall_delta <= 0:
            return None
        pct = (proc_delta / wall_delta) * 100.0 / _cpu_logical_count
        return max(0.0, min(100.0, pct))
    except Exception:
        return None


# 整机 CPU 采样状态（GetSystemTimes：idle/kernel/user 时间增量，1 - idle/total = 占用率）
_sys_cpu_last_idle = None
_sys_cpu_last_total = None


def get_system_cpu_percent():
    """返回整机 CPU 占用率（0~100）。Windows 用 GetSystemTimes 计算 idle/total 增量；
    非 Windows 返回 None。需间隔调用两次取增量。"""
    global _sys_cpu_last_idle, _sys_cpu_last_total
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD), ('dwHighDateTime', wintypes.DWORD)]

        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        idle = FILETIME(); kern = FILETIME(); user = FILETIME()
        if not kernel32.GetSystemTimes(ctypes.byref(idle), ctypes.byref(kern), ctypes.byref(user)):
            return None
        idle_t = (idle.dwHighDateTime << 32) | idle.dwLowDateTime
        kern_t = (kern.dwHighDateTime << 32) | kern.dwLowDateTime  # 含 idle
        user_t = (user.dwHighDateTime << 32) | user.dwLowDateTime
        total = kern_t + user_t
        if _sys_cpu_last_total is None:
            _sys_cpu_last_idle = idle_t
            _sys_cpu_last_total = total
            return None
        idle_d = idle_t - _sys_cpu_last_idle
        total_d = total - _sys_cpu_last_total
        _sys_cpu_last_idle = idle_t
        _sys_cpu_last_total = total
        if total_d <= 0:
            return None
        return max(0.0, min(100.0, (1.0 - idle_d / total_d) * 100.0))
    except Exception:
        return None


def get_system_memory_status():
    """返回整机内存 (used_mb, total_mb, percent_used)；非 Windows 或失败返回 None。"""
    if os.name != 'nt':
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ('dwLength', wintypes.DWORD),
                ('dwMemoryLoad', wintypes.DWORD),
                ('ullTotalPhys', ctypes.c_ulonglong),
                ('ullAvailPhys', ctypes.c_ulonglong),
                ('ullTotalPageFile', ctypes.c_ulonglong),
                ('ullAvailPageFile', ctypes.c_ulonglong),
                ('ullTotalVirtual', ctypes.c_ulonglong),
                ('ullAvailVirtual', ctypes.c_ulonglong),
                ('ullAvailExtendedVirtual', ctypes.c_ulonglong),
            ]

        stat = MEMORYSTATUSEX()
        stat.dwLength = ctypes.sizeof(stat)
        if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat)):
            return None
        total_mb = stat.ullTotalPhys / (1024 * 1024)
        used_mb = (stat.ullTotalPhys - stat.ullAvailPhys) / (1024 * 1024)
        return (used_mb, total_mb, int(stat.dwMemoryLoad))
    except Exception:
        return None


# 启动外部进程时与当前进程解耦，避免主程序退出时连带关闭子进程
_DETACHED_FLAGS = 0
_NEW_PROCESS_GROUP = 0
_BREAKAWAY_FROM_JOB = 0
if os.name == 'nt':
    _DETACHED_FLAGS = getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
    _NEW_PROCESS_GROUP = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
    # break away 确保即使当前进程被放入 Job 对象，子进程也能独立存活
    _BREAKAWAY_FROM_JOB = 0x01000000


def launch_detached(cmd, cwd=None, extra_creationflags=0):
    """以与主进程解耦的方式启动外部程序，确保主进程退出后子进程仍然存活。

    在 VS Code / CI 等受限 Job Object 环境下，CREATE_BREAKAWAY_FROM_JOB 会被拒绝
    （ERROR_ACCESS_DENIED）。此时自动降级到 ShellExecuteW，Shell 进程本身在 Job
    之外创建新进程，完全独立于父进程的生命周期。
    """
    if os.name == 'nt':
        flags = _DETACHED_FLAGS | _NEW_PROCESS_GROUP | _BREAKAWAY_FROM_JOB | extra_creationflags
        try:
            return subprocess.Popen(cmd, cwd=cwd, creationflags=flags, close_fds=True)
        except OSError:
            # Job Object 不允许 breakaway —— 改用 ShellExecuteW（经由 Shell 创建，天然在 Job 外）
            exe = cmd[0] if isinstance(cmd, list) else cmd
            params = subprocess.list2cmdline(cmd[1:]) if isinstance(cmd, list) and len(cmd) > 1 else None
            try:
                import ctypes
                ctypes.windll.shell32.ShellExecuteW(
                    None, 'open', exe, params, cwd, 1  # 1 = SW_SHOWNORMAL
                )
                return None  # ShellExecuteW 不返回 Popen 对象
            except Exception:
                # 最终兜底：cmd /c start "" 也能从 Job Object 解脱
                start_cmd = ['cmd.exe', '/c', 'start', '""'] + (cmd if isinstance(cmd, list) else [cmd])
                return subprocess.Popen(start_cmd, cwd=cwd, close_fds=True)
    # 非 Windows 环境使用新 session，避免收到父进程信号
    return subprocess.Popen(cmd, cwd=cwd, start_new_session=True, close_fds=True)


def launch_detached_async(cmd, cwd=None, extra_creationflags=0):
    """在后台守护线程执行 launch_detached，避免 CreateProcess（及 Job Object 降级链）
    在 UI 线程阻塞。仅用于 fire-and-forget 场景（不使用返回的 Popen）。

    子进程一旦创建即与父进程解耦（DETACHED/BREAKAWAY），后台线程随即结束不影响其存活。
    调用方应先在 UI 线程完成校验（路径/可执行文件存在性）再调用本函数，以便错误提示仍能弹出。"""
    def _worker():
        try:
            launch_detached(cmd, cwd=cwd, extra_creationflags=extra_creationflags)
        except Exception as e:
            debug_print(f"[launch_detached_async] failed for {cmd!r}: {e}")
    try:
        threading.Thread(target=_worker, daemon=True).start()
    except Exception as e:
        # 线程创建失败极罕见：退回同步启动，保证功能可用
        debug_print(f"[launch_detached_async] thread start failed, fallback sync: {e}")
        try:
            launch_detached(cmd, cwd=cwd, extra_creationflags=extra_creationflags)
        except Exception as e2:
            debug_print(f"[launch_detached_async] sync fallback failed: {e2}")


def normalize_external_launch_dir(path):
    if not path:
        return None
    return os.path.normpath(path) if os.name == 'nt' else path


def find_git_install_root():
    git_root_candidates = [
        r"C:\Program Files\Git",
        r"C:\Program Files (x86)\Git",
        os.path.expandvars(r"%PROGRAMFILES%\Git"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Git"),
    ]
    return next((path for path in git_root_candidates if os.path.isdir(path)), None)


def launch_shell_tool(tool_name, cwd=None):
    cwd = normalize_external_launch_dir(cwd)

    if tool_name in ('cmd', 'powershell', 'git-bash'):
        if not cwd or not os.path.isdir(cwd):
            raise FileNotFoundError(tr("当前路径无效，无法启动终端"))

    # 使用 CREATE_NEW_CONSOLE | CREATE_NEW_PROCESS_GROUP | CREATE_BREAKAWAY_FROM_JOB
    # 确保启动的终端/程序在 TabEx 退出后仍然存活
    # 注意：CREATE_NEW_CONSOLE 和 DETACHED_PROCESS 互斥，终端需要 console 所以不用 DETACHED
    _DETACH_CONSOLE = (getattr(subprocess, 'CREATE_NEW_CONSOLE', 0x00000010)
                       | _NEW_PROCESS_GROUP | _BREAKAWAY_FROM_JOB)

    def _spawn_console_async(argv):
        """在后台守护线程创建带新控制台的进程，避免 CreateProcess 阻塞 UI 线程。
        路径/工具校验已在上方同步完成，故此处异常仅记录日志。"""
        def _worker():
            try:
                subprocess.Popen(argv, creationflags=_DETACH_CONSOLE, close_fds=True, cwd=cwd)
            except Exception as e:
                debug_print(f"[launch_shell_tool] async spawn failed for {argv!r}: {e}")
        try:
            threading.Thread(target=_worker, daemon=True).start()
        except Exception:
            _worker()  # 线程创建失败：退回同步

    if tool_name == 'cmd':
        _spawn_console_async(['cmd.exe', '/K', 'cd', '/d', cwd])
        return None

    if tool_name == 'powershell':
        escaped_dir = cwd.replace("'", "''")
        _spawn_console_async(['powershell.exe', '-NoExit', '-Command', f"Set-Location -LiteralPath '{escaped_dir}'"])
        return None

    if tool_name == 'git-bash':
        git_root = find_git_install_root()
        if not git_root:
            raise FileNotFoundError(tr("未找到 Git Bash，请确认已安装 Git for Windows"))
        git_bash_exe = os.path.join(git_root, 'git-bash.exe')
        if not os.path.exists(git_bash_exe):
            raise FileNotFoundError(tr("未找到可用的 Git Bash 可执行文件"))
        launch_detached_async([git_bash_exe, f'--cd={cwd}'], cwd=cwd)
        return None

    if tool_name == 'calculator':
        if os.name != 'nt':
            raise OSError(tr("当前系统不支持打开计算器"))
        launch_detached_async(['calc.exe'])
        return None

    raise ValueError(f"Unsupported tool: {tool_name}")


def is_supported_title_shortcut_path(path):
    return isinstance(path, str) and bool(path) and os.path.isfile(path) and path.lower().endswith(TITLE_SHORTCUT_EXTENSIONS)


def normalize_terminal_tool_name(tool_name, default='cmd'):
    if not isinstance(tool_name, str):
        return default
    value = tool_name.strip().lower()
    alias_map = {
        'cmd': 'cmd',
        'command prompt': 'cmd',
        'powershell': 'powershell',
        'ps': 'powershell',
        'pwsh': 'powershell',
        'git bash': 'git-bash',
        'git-bash': 'git-bash',
        'bash': 'git-bash',
    }
    normalized = alias_map.get(value, value)
    return normalized if normalized in SUPPORTED_TERMINAL_TOOLS else default


# ==================== 异步文件夹大小检查线程 ====================
_retained_background_threads = set()


def _retain_thread_until_finished(worker):
    if worker in _retained_background_threads:
        return
    worker.setParent(None)
    _retained_background_threads.add(worker)

    def release():
        if worker not in _retained_background_threads:
            return
        _retained_background_threads.discard(worker)
        worker.deleteLater()

    worker.finished.connect(release)
    if not worker.isRunning():
        try:
            worker.finished.disconnect(release)
        except (TypeError, RuntimeError):
            pass
        release()


def _thread_name_summary(limit=10):
    """按去掉末尾编号后的线程名计数，用于在诊断记录中发现线程堆积。"""
    import re
    from collections import Counter
    names = Counter(re.sub(r'[-_ ]?\d+$', '', thread.name) for thread in threading.enumerate())
    return dict(names.most_common(limit))


# Optional native hit-test support (Windows)
try:
    import ctypes
    import win32gui
    import win32con
    HAS_PYWIN = True
except Exception:
    HAS_PYWIN = False

# Windows API for monitoring new Explorer windows
if HAS_PYWIN:
    try:
        import ctypes.wintypes as wintypes
        user32 = ctypes.windll.user32
        ole32 = ctypes.windll.ole32
        
        # Constants for SetWinEventHook
        EVENT_OBJECT_CREATE = 0x8000
        EVENT_SYSTEM_FOREGROUND = 0x0003
        WINEVENT_OUTOFCONTEXT = 0x0000
        
        # Define callback type
        WinEventProcType = ctypes.WINFUNCTYPE(
            None,
            wintypes.HANDLE,
            wintypes.DWORD,
            wintypes.HWND,
            wintypes.LONG,
            wintypes.LONG,
            wintypes.DWORD,
            wintypes.DWORD
        )
    except Exception as e:
        debug_print(f"Failed to setup Windows API monitoring: {e}")

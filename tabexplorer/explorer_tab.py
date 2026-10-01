"""单个文件浏览标签页。"""

import ctypes.wintypes
import os
import time

from PyQt5.QtCore import pyqtSignal, QDir, QFileSystemWatcher, QObject, QRunnable, Qt, QThreadPool, QTimer
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QApplication, QHBoxLayout, QLabel, QMenu, QProgressBar, QPushButton, QVBoxLayout, QWidget,
)

from . import theme as _theme
from .paths import get_app_base_dir, translate_common_path
from .i18n import tr
from .constants import (
    APP_INTERNAL_CHANGE_FILENAMES, ASYNC_LOAD_ENABLED, ASYNC_NAV_TIMEOUT_MS, COM_POLL_HARD_DEADLINE_MS,
    COM_POLL_MIN_GAP_MS, COM_POLL_SLOW_MS, COM_POLL_STRESS_BACKOFF_MS, FOLDER_CHECK_TIMEOUT,
    MAX_NAVIGATION_HISTORY, STATUS_SELECTION_METADATA_LIMIT, STATUS_TRACKING_INTERVAL_MS,
    STATUS_TRACKING_WINDOW_MS, STATUS_UPDATE_DEFER_MS,
)
from .debuglog import debug_print
from .system import (
    detect_notepad_plus_plus, find_git_install_root, HAS_PYWIN, launch_detached, launch_detached_async,
    launch_shell_tool, _retain_thread_until_finished,
)
from .hotkeys import _hotkey_hint, _hotkey_text
from .widgets import GestureOverlay, show_toast
from .workers import FolderSizeChecker, GitStatusWorker
from .search import _search_cache
from .pathbar import SimplePathBar
from .netstatus import (
    ERROR_CANCELLED, ERROR_CONNECTION_UNAVAIL, is_network_path, NAV_TIMEOUT, NetworkErrorBanner, RECONNECTED_CODES,
    ReconnectWorker, win32_error,
)
from .shellview import _COMTYPES_AVAILABLE, _file_url_to_path, IExplorerBrowserWidget

if HAS_PYWIN:
    import win32gui


# 兜底轮询快照逐项 stat 的上限：超大目录每 8s 在 UI 线程逐项 stat 会造成周期性卡顿，
# 超过此上限后停止逐项统计，改用“总项目数 + 目录自身 mtime”兜底检测增删/重命名。
DIR_SNAPSHOT_MAX_ENTRIES = 5000


def _compute_dir_snapshot(path, ignore_check=None):
    """计算目录轻量元数据快照 (count, latest_mtime_ns, size_sum, name_hash)，失败返回 None。

    纯 I/O + 计算，不触碰任何 Qt 对象，可安全在后台线程（QRunnable）中执行。
    ignore_check(path, name) 用于忽略应用自身写出的配置/日志文件（可为 None）。"""
    try:
        count = 0
        latest_mtime_ns = 0
        size_sum = 0
        name_hash = 0
        truncated = False
        with os.scandir(path) as entries:
            for entry in entries:
                if ignore_check is not None and ignore_check(path, entry.name):
                    continue
                count += 1
                if count > DIR_SNAPSHOT_MAX_ENTRIES:
                    # 超大目录：停止逐项 stat，仍继续累计总数，配合目录自身 mtime 兜底
                    truncated = True
                    continue
                try:
                    is_dir = entry.is_dir(follow_symlinks=False)
                    stat = entry.stat(follow_symlinks=False)
                    mtime_ns = getattr(stat, 'st_mtime_ns', None)
                    if mtime_ns is None:
                        mtime_ns = int(stat.st_mtime * 1_000_000_000)
                    if mtime_ns > latest_mtime_ns:
                        latest_mtime_ns = mtime_ns
                    if not is_dir:
                        size_sum += getattr(stat, 'st_size', 0)
                    name_hash ^= hash((entry.name, is_dir))
                except Exception:
                    continue
        if truncated:
            # 折叠目录自身 mtime，保证超出上限部分的增删/重命名仍能触发刷新
            try:
                dir_mtime_ns = os.stat(path).st_mtime_ns
                if dir_mtime_ns > latest_mtime_ns:
                    latest_mtime_ns = dir_mtime_ns
            except Exception:
                pass
        return (count, latest_mtime_ns, size_sum, name_hash)
    except Exception:
        return None


class _DirSnapshotSignals(QObject):
    """后台目录快照计算完成信号：done(path, snapshot_or_None)。"""
    done = pyqtSignal(str, object)


class _DirSnapshotRunnable(QRunnable):
    """在 QThreadPool 后台线程计算目录快照，完成后经信号回到 UI 线程比较。

    用于 FileExplorerTab 的 8 秒兜底轮询，避免在 UI 线程逐项 stat 造成周期性卡顿。"""
    def __init__(self, path, ignore_check, signals):
        super().__init__()
        self._path = path
        self._ignore_check = ignore_check
        self._signals = signals

    def run(self):
        snap = _compute_dir_snapshot(self._path, self._ignore_check)
        try:
            self._signals.done.emit(self._path, snap)
        except RuntimeError:
            # 信号对象已随标签销毁：忽略
            pass


class FileExplorerTab(QWidget):
    # Singleton WH_MOUSE_LL hook for double-click detection (shared across all IEB tabs)
    _global_mouse_hook_handle = None
    _global_mouse_hook_cb = None

    def _force_remove_watcher(self, path):
        """强制移除watcher路径，防止事件风暴"""
        try:
            if path in self.file_watcher.directories():
                self.file_watcher.removePath(path)
                debug_print(f"[FileWatcher] Force removed: {path}")
            if path in self.file_watcher.files():
                self.file_watcher.removePath(path)
                debug_print(f"[FileWatcher] Force removed file: {path}")
        except Exception as e:
            debug_print(f"[FileWatcher] Exception in force_remove_watcher: {e}")

    def _build_dir_snapshot(self, path):
        """构建当前目录的轻量元数据快照，用于检测文件修改时间/大小变化。

        纯计算逻辑委托给模块级 _compute_dir_snapshot()，以便定时兜底轮询可在后台线程
        复用同一算法（见 _poll_directory_changes 的 QRunnable 异步路径），避免 UI 线程逐项 stat。"""
        try:
            if not path or self._is_slow_path(path):
                return None
            return _compute_dir_snapshot(path, self._should_ignore_internal_dir_entry)
        except Exception:
            return None

    def _should_ignore_internal_dir_entry(self, dir_path, entry_name):
        try:
            if not dir_path or not entry_name:
                return False
            app_base_dir = os.path.normcase(os.path.normpath(get_app_base_dir()))
            current_dir = os.path.normcase(os.path.normpath(dir_path))
            if current_dir != app_base_dir:
                return False
            entry_name_lower = str(entry_name).lower()
            if entry_name_lower in APP_INTERNAL_CHANGE_FILENAMES:
                return True
            if entry_name_lower.startswith('config.json.'):
                return True
        except Exception:
            return False
        return False

    def _refresh_file_watch_paths(self, path):
        """同步当前目录下的文件 watcher，提升对文件内容修改的检测能力。"""
        # 文件级 Watcher 已关闭（节省系统句柄资源），目录 Watcher + DirPoll 兜底仍生效
        return
        if not hasattr(self, 'file_watcher'):
            return
        try:
            # 清理旧文件监听
            old_files = getattr(self, '_watched_files', set())
            for file_path in list(old_files):
                self._force_remove_watcher(file_path)
            self._watched_files = set()

            if not path or not os.path.isdir(path):
                return
            # OneDrive/网络路径的os.scandir()会阻塞UI线程，跳过文件监视
            if self._is_slow_path(path):
                debug_print(f"[FileWatcher] Skipping file watch for slow path: {path}")
                return

            max_watch_files = 200
            watched = set()
            with os.scandir(path) as entries:
                for entry in entries:
                    if len(watched) >= max_watch_files:
                        break
                    try:
                        if entry.is_file(follow_symlinks=False):
                            file_path = os.path.normpath(entry.path)
                            if self.file_watcher.addPath(file_path):
                                watched.add(file_path)
                    except Exception:
                        continue
            self._watched_files = watched
            debug_print(f"[FileWatcher] Watching {len(watched)} files in current dir")
        except Exception as e:
            debug_print(f"[FileWatcher] Failed to refresh file watch paths: {e}")

    def set_refresh_active(self, active: bool):
        """设置当前标签的刷新活跃态：仅当前可见标签执行高频刷新。"""
        self._refresh_active = bool(active)
        self._last_active_at = time.monotonic()
        current_path = getattr(self, 'current_path', '')

        if self._refresh_active:
            if getattr(self, '_hibernated', False):
                self._hibernated = False
                mw = getattr(self, 'main_window', None)
                if mw is not None and hasattr(mw, '_schedule_tab_labels'):
                    mw._schedule_tab_labels()
            if current_path and not current_path.startswith(('shell:', '::')):
                is_slow = self._is_slow_path(current_path)
                # OneDrive/网络路径：跳过os.scandir操作，避免阻塞UI线程和COM消息泵
                if not is_slow:
                    self._poll_directory_changes()
                if hasattr(self, 'dir_poll_timer') and self.dir_poll_timer and not self.dir_poll_timer.isActive():
                    if not is_slow:
                        self.dir_poll_timer.start()
                        debug_print(f"[DirPoll] Activated polling for visible tab: {current_path}")
                    else:
                        debug_print(f"[DirPoll] Skipped polling for slow path: {current_path}")
                self._consume_pending_refresh(fallback_reason="activate_tab")
            # 标签激活时，若主同步未运行，启动保活轮询以捕获导航变化
            if not (hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive()):
                self._start_keepalive_sync()
        else:
            self._refresh_file_watch_paths(None)
            if hasattr(self, 'dir_poll_timer') and self.dir_poll_timer and self.dir_poll_timer.isActive():
                self.dir_poll_timer.stop()
                debug_print(f"[DirPoll] Suspended polling for background tab: {current_path}")
            # 标签停用时，停止保活轮询
            if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer:
                self._keepalive_sync_timer.stop()
            # 若标签在导航过程中被切到后台，立即停止瞬时路径同步轮询，
            # 避免后台标签继续以 120~360ms 高频跨进程读取 COM LocationURL
            # （否则需等到下一次 tick 才自停，确保“仅活动标签轮询”）。
            if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive():
                self._path_sync_timer.stop()
            if hasattr(self, '_path_sync_stop_timer') and self._path_sync_stop_timer and self._path_sync_stop_timer.isActive():
                self._path_sync_stop_timer.stop()

    def is_auto_refresh_frozen(self):
        return bool(getattr(self, '_manual_refresh_frozen', False))

    def hibernate_shell_view(self):
        """释放长时间未使用的后台标签的 Shell 视图，切回时按当前路径重建。"""
        explorer = getattr(self, 'explorer', None)
        path = getattr(self, 'current_path', '')
        if not path or not isinstance(explorer, IExplorerBrowserWidget) or not explorer._init_ok:
            return False
        if getattr(self, '_refresh_active', False) or self.isVisible() or getattr(self, '_nav_in_progress', False):
            return False
        try:
            worker = getattr(self, '_file_op_worker', None)
            if worker is not None and worker.isRunning():
                return False
        except RuntimeError:
            pass  # worker 已被 deleteLater 销毁
        if not explorer.hibernate(path):
            return False
        self._hibernated = True
        debug_print(f"[Hibernate] Released shell view of idle tab: {path}")
        return True

    def set_auto_refresh_frozen(self, frozen):
        self._manual_refresh_frozen = bool(frozen)
        if self._manual_refresh_frozen and hasattr(self, 'refresh_timer') and self.refresh_timer.isActive():
            self.refresh_timer.stop()
        if not self._manual_refresh_frozen:
            self._consume_pending_refresh(fallback_reason="manual_unfreeze")
        return self._manual_refresh_frozen

    def _arm_selection_guard(self, seconds=6.0):
        """Ctrl+G 选中后，开启一段时间的选中保护窗口，期间抑制自动刷新。"""
        try:
            self._selection_guard_until = time.monotonic() + float(seconds)
        except Exception:
            self._selection_guard_until = 0.0

    def _selection_guard_active(self):
        try:
            return time.monotonic() < float(getattr(self, '_selection_guard_until', 0) or 0)
        except Exception:
            return False

    def _request_refresh(self, reason="manual"):
        """统一记录刷新请求，当前标签可见时立即调度，不可见时延后到激活后消费。"""
        self._refresh_pending = True
        self._refresh_pending_reason = reason

        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[AutoRefresh] Suppressed during navigation (reason={reason})")
            return False
        if self._selection_guard_active():
            debug_print(f"[AutoRefresh] Suppressed during selection guard (reason={reason})")
            return False
        if getattr(self, '_manual_refresh_frozen', False):
            debug_print(f"[AutoRefresh] Manually frozen (reason={reason})")
            return False
        if not getattr(self, '_refresh_active', True):
            debug_print(f"[AutoRefresh] Deferred while tab inactive (reason={reason})")
            return False

        self._schedule_refresh(reason=reason)
        return True

    def _consume_pending_refresh(self, fallback_reason="manual"):
        if not getattr(self, '_refresh_pending', False):
            return False
        if getattr(self, '_manual_refresh_frozen', False):
            return False
        reason = getattr(self, '_refresh_pending_reason', None) or fallback_reason
        self._schedule_refresh(reason=reason)
        return True

    def on_file_changed(self, path):
        """文件内容或元数据变化时，主动刷新当前视图。"""
        try:
            import time
            now_ms = time.time() * 1000
            norm_file = os.path.normcase(os.path.normpath(path))
            # 文件级去抖：同一路径短时间内重复事件只处理一次
            last_ms = self._last_file_event.get(norm_file, 0)
            if now_ms - last_ms < getattr(self, '_file_event_debounce_ms', 400):
                return
            self._last_file_event[norm_file] = now_ms
            if len(self._last_file_event) > 200:
                # 仅保留最近触发的100条，避免字典无限增长
                items = sorted(self._last_file_event.items(), key=lambda x: x[1], reverse=True)
                self._last_file_event = dict(items[:100])

            debug_print(f"[FileWatcher] File changed: {path}")
            current_dir = getattr(self, 'current_path', '')
            if current_dir and os.path.normcase(os.path.dirname(path)) == os.path.normcase(current_dir):
                if not getattr(self, '_refresh_active', True):
                    self._request_refresh(reason="file_changed")
                    debug_print(f"[FileWatcher] Background tab file changed, marked dirty: {path}")
                    return
                # 某些编辑器会以重命名替换文件，变化后需重新添加 watcher
                if os.path.isfile(path):
                    try:
                        if path not in self.file_watcher.files():
                            self.file_watcher.addPath(path)
                    except Exception:
                        pass
                self._request_refresh(reason="file_changed")
        except Exception as e:
            debug_print(f"[FileWatcher] on_file_changed error: {e}")

    def _activate_path_bar(self):
        owner = self.main_window
        if owner is not None:
            for tabs, stack in owner._all_groups():
                if stack.indexOf(self) >= 0:
                    owner.set_active_pane_to_group(tabs)
                    break

    def update_tab_title(self):
        if hasattr(self, 'current_path'):
            # 兜底同步路径栏：有些导航路径变化来自 Explorer 内部事件，
            # 可能只触发标题更新，不走 navigate_to()。
            try:
                if hasattr(self, 'path_bar') and self.path_bar:
                    self.path_bar.set_path(self.current_path)
            except Exception:
                pass

            # shell: 路径中文映射（字典 O(1) 查找）
            _SHELL_PATH_MAP = {
                'shell:RecycleBinFolder': tr('回收站'),
                'shell:MyComputerFolder': tr('此电脑'),
                'shell:Desktop': '桌面',
                'shell:NetworkPlacesFolder': tr('网络'),
            }
            path = self.current_path
            display = _SHELL_PATH_MAP.get(path, None)
            
            if not display and path.startswith('shell:'):
                display = path  # 兜底显示原始shell:路径
            
            if not display:
                # 普通路径：显示最后一层文件夹名称
                # 标准化路径分隔符
                normalized_path = path.replace('/', '\\') if os.name == 'nt' else path
                # 移除末尾的分隔符
                normalized_path = normalized_path.rstrip('\\/')
                
                if normalized_path:
                    # 获取最后一层文件夹名称
                    folder_name = os.path.basename(normalized_path)
                    
                    # 如果是驱动器根目录（如 C:），直接显示
                    if not folder_name and ':' in normalized_path:
                        folder_name = normalized_path
                    # 如果是UNC路径根目录（如 \\server\share），显示share
                    elif not folder_name and normalized_path.startswith('\\\\'):
                        parts = normalized_path.split('\\')
                        folder_name = parts[-1] if parts[-1] else parts[-2] if len(parts) > 2 else normalized_path
                    
                    display = folder_name if folder_name else path
                else:
                    display = path
            
            # 处理固定标签和普通标签的显示
            is_pinned = getattr(self, 'is_pinned', False)
            pin_prefix = "📌 " if is_pinned else ""
            
            # 如果是固定标签，限制display长度以确保📌始终显示
            # 标签宽度120px，📌+空格约占15px，剩余约105px可显示文本
            # 一个中文字符约12px，英文约7px，预估最多显示约15个字符
            if is_pinned:
                max_display_len = 15  # 为📌预留空间
                if len(display) > max_display_len:
                    display = "..." + display[-(max_display_len-3):]
            
            title = pin_prefix + display
            debug_print(f"DEBUG update_tab_title: path={path}, is_pinned={is_pinned}, pin_prefix='{pin_prefix}', title='{title}'")
            mw = self.main_window
            if mw and hasattr(mw, 'tab_widget'):
                # 定位该标签所属的组（左侧主组或右侧分屏组），只更新其所在标签栏
                target_tw = None
                idx = -1
                if hasattr(mw, '_all_groups'):
                    for cand_tw, cand_cs in mw._all_groups():
                        if cand_cs is not None:
                            j = cand_cs.indexOf(self)
                            if j != -1:
                                target_tw, idx = cand_tw, j
                                break
                if target_tw is None:
                    # 回退：默认左侧主组（因 tab_widget 只含占位符，需在 content_stack 中查找索引）
                    target_tw = mw.tab_widget
                    if hasattr(mw, 'content_stack'):
                        idx = mw.content_stack.indexOf(self)
                    else:
                        idx = mw.tab_widget.indexOf(self)

                if idx != -1:
                    if hasattr(mw, '_refresh_tab_labels'):
                        mw._refresh_tab_labels()
                    else:
                        target_tw.setTabText(idx, title)
                    if hasattr(mw, '_apply_tab_group_color'):
                        mw._apply_tab_group_color(target_tw, idx, self)
                    debug_print(f"DEBUG: Set tab {idx} text to '{title}'")
                    if hasattr(mw, '_schedule_session_snapshot'):
                        mw._schedule_session_snapshot()

    def start_path_sync_timer(self, duration_ms=2000):
        """启动路径同步定时器，duration_ms 后自动停止（按需触发，减少持续COM调用）"""
        from PyQt5.QtCore import QTimer
        # 主同步启动时暂停保活轮询，避免双重COM查询
        if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer and self._keepalive_sync_timer.isActive():
            self._keepalive_sync_timer.stop()
        self._path_sync_stable_hits = 0
        self._path_sync_interval_ms = 120
        if not hasattr(self, '_path_sync_timer') or self._path_sync_timer is None:
            self._path_sync_timer = QTimer(self)
            self._path_sync_timer.timeout.connect(self.sync_path_bar_with_explorer)
        # 设置自动停止定时器：导航期间轮询，稳定后停止
        if not hasattr(self, '_path_sync_stop_timer') or self._path_sync_stop_timer is None:
            self._path_sync_stop_timer = QTimer(self)
            self._path_sync_stop_timer.setSingleShot(True)
            self._path_sync_stop_timer.timeout.connect(self._stop_path_sync_timer)
        self._path_sync_stop_timer.start(duration_ms)
        if not self._path_sync_timer.isActive():
            self._path_sync_timer.start(self._path_sync_interval_ms)
        elif self._path_sync_timer.interval() != self._path_sync_interval_ms:
            self._path_sync_timer.setInterval(self._path_sync_interval_ms)

    def _stop_path_sync_timer(self):
        """停止路径同步轮询（稳定后调用，避免持续COM跨进程调用）"""
        if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive():
            self._path_sync_timer.stop()
        self._path_sync_stable_hits = 0
        self._path_sync_interval_ms = 120
        # 主同步结束后，启动低频保活轮询以兜底 NavigateComplete2 遗漏的用户导航
        if getattr(self, '_refresh_active', False):
            self._start_keepalive_sync()

    def _read_location_url_timed(self):
        """读取 LocationURL（同步跨进程 COM），并测量耗时以检测系统高负载。

        返回读取到的 URL；若本次调用耗时超过 COM_POLL_SLOW_MS，则设置高负载退避截止
        时间戳 self._com_stress_until，供轮询函数据此拉长间隔，打断 CPU 高时的死亡螺旋。

        对真正跨进程的旧 QAx Shell.Explorer 控件，改用带硬超时的
        看门狗读取：调用挂死时 UI 线程最多阻塞 COM_POLL_HARD_DEADLINE_MS 即放弃并返回
        上次已知 URL，彻底根治“无超时同步 COM 卡死 UI”。IExplorerBrowser 的 LocationURL
        是进程内缓存值（快），直接在 UI 线程读取。"""
        # IEB 缓存值：进程内、零阻塞，直接读
        if isinstance(self.explorer, IExplorerBrowserWidget):
            return self.explorer.property('LocationURL')
        # 旧 QAx：跨进程 COM，加看门狗硬超时
        start = time.perf_counter()
        url = self._read_qax_location_with_watchdog()
        elapsed_ms = (time.perf_counter() - start) * 1000.0
        if elapsed_ms >= COM_POLL_SLOW_MS:
            self._com_stress_until = time.monotonic() + (COM_POLL_STRESS_BACKOFF_MS / 1000.0)
            debug_print(f"[PathSync] Slow LocationURL read {elapsed_ms:.0f}ms -> backoff "
                        f"{COM_POLL_STRESS_BACKOFF_MS}ms")
        if url:
            self._last_location_url = url
        return url

    def _read_qax_location_with_watchdog(self):
        """带硬超时的 QAx LocationURL 读取：超时则返回上次缓存值，避免阻塞 UI。

        看门狗线程读取跨进程 COM 属性前先 CoInitialize（STA），避免在未初始化
        COM 的工作线程上调用导致失败；同时用 _com_inflight 抑制重叠提交，避免超时后旧任务
        仍占用单 worker 造成 future 堆积。"""
        from concurrent.futures import ThreadPoolExecutor, TimeoutError as _FTimeout
        # 上一次读取尚未返回：不重复提交，直接用缓存值（worker 阻塞时避免任务堆积）
        if getattr(self, '_com_inflight', False):
            return getattr(self, '_last_location_url', '')
        ex = getattr(self, '_com_executor', None)
        if ex is None:
            ex = ThreadPoolExecutor(max_workers=1, thread_name_prefix='com-loc')
            self._com_executor = ex
        self._com_inflight = True
        try:
            fut = ex.submit(self._qax_read_location_worker)
            return fut.result(timeout=COM_POLL_HARD_DEADLINE_MS / 1000.0)
        except _FTimeout:
            self._com_stress_until = time.monotonic() + (COM_POLL_STRESS_BACKOFF_MS / 1000.0)
            debug_print(f"[PathSync] LocationURL watchdog timeout -> using cached, backoff")
            return getattr(self, '_last_location_url', '')
        except Exception as e:
            debug_print(f"[PathSync] watchdog read error: {e}")
            return getattr(self, '_last_location_url', '')

    def _qax_read_location_worker(self):
        """worker 线程：初始化 STA 公寓后读取 LocationURL，结束后清除 in-flight 标志。"""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            return self.explorer.property('LocationURL')
        finally:
            self._com_inflight = False
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass



    def _com_under_stress(self):
        """当前是否处于 COM 高负载退避窗口内。"""
        return time.monotonic() < getattr(self, '_com_stress_until', 0.0)

    def _poll_burst_guard(self):
        """挂钟防抖：吸收 UI 线程卡顿后堆积的爆发式定时器触发。

        若距上次实际轮询的真实时间间隔小于 COM_POLL_MIN_GAP_MS，返回 True 表示应跳过本次，
        避免连续多次慢 COM 调用把线程进一步压死。"""
        now_ms = time.monotonic() * 1000.0
        last_ms = getattr(self, '_last_poll_wall_ms', 0.0)
        if now_ms - last_ms < COM_POLL_MIN_GAP_MS:
            return True
        self._last_poll_wall_ms = now_ms
        return False

    def _start_keepalive_sync(self):
        """启动低频保活轮询（每1500ms检测LocationURL变化，主同步未覆盖时兜底）"""
        # 如果主同步正在运行，不启动保活（避免双重轮询）
        if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.isActive():
            return
        if not hasattr(self, '_keepalive_sync_timer') or self._keepalive_sync_timer is None:
            self._keepalive_sync_timer = QTimer(self)
            self._keepalive_sync_timer.timeout.connect(self._keepalive_sync_check)
        if not self._keepalive_sync_timer.isActive():
            self._keepalive_sync_timer.start(1500)

    def _keepalive_sync_check(self):
        """保活检查：若检测到LocationURL与current_path不一致，重启主同步更新路径栏。"""
        if not getattr(self, '_refresh_active', False):
            if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer:
                self._keepalive_sync_timer.stop()
            return
        # 窗口最小化或失去前台焦点时跳过COM轮询：用户无法在内嵌shell视图中导航，
        # 窗口最小化时跳过COM轮询：用户无法在内嵌shell视图中导航，
        # 持续读取 LocationURL 只会触发后台 shell worker 线程 churn。
        # 注意：不要用 isActiveWindow() 判断——内嵌 IExplorerBrowser/QAx 控件经常抢占焦点，
        # 主窗口虽在前台却报告非激活，会导致路径同步永久休眠、地址栏冻结需手动 resize 才恢复。
        mw = getattr(self, 'main_window', None)
        if mw is not None:
            try:
                if mw.windowState() & Qt.WindowMinimized:
                    return
            except Exception:
                pass
        # 抗高负载：爆发触发跳过 + 退避窗口内跳过 COM 读取（与主同步一致），避免 CPU 高时压死 UI 线程
        if self._poll_burst_guard():
            return
        if self._com_under_stress():
            return
        try:
            url = self._read_location_url_timed()
            if url:
                url_str = str(url)
                local_path = None
                if url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = unquote(url_str[8:])
                    if os.name == 'nt' and local_path.startswith('/'):
                        local_path = local_path[1:]
                    local_path = self._normalize_local_path(local_path)
                elif url_str.startswith('file://') and not url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = self._normalize_local_path('\\\\' + unquote(url_str[7:]).replace('/', '\\'))
                elif '::' in url_str:
                    # 尝试从 CLSID URL 中提取盘符路径（如映射网络驱动器 ::{...}\X:\）
                    import re
                    drive_match = re.search(r'([A-Za-z]:\\[^"]*)', url_str.replace('/', '\\'))
                    if not drive_match:
                        drive_match = re.search(r'([A-Za-z]:/[^"]*)', url_str)
                    if drive_match:
                        local_path = self._normalize_local_path(drive_match.group(1))
                current_path = self._normalize_local_path(getattr(self, 'current_path', ''))
                shell_activation_alive = False
                if local_path and local_path != current_path:
                    if self._is_mycomputer_shell_target(current_path):
                        allow_until = float(getattr(self, '_shell_item_activation_until', 0.0) or 0.0)
                        now = time.monotonic()
                        if allow_until <= now:
                            debug_print(
                                f"[Keepalive] REJECT shell->local '{local_path}' "
                                f"reason=no-user-activation current='{current_path}'"
                            )
                            self._keepalive_candidate_path = None
                            self._keepalive_candidate_count = 0
                            return
                        remain_ms = int((allow_until - now) * 1000.0)
                        debug_print(
                            f"[Keepalive] ALLOW shell->local candidate '{local_path}' "
                            f"reason=user-activation remain={remain_ms}ms"
                        )
                        shell_activation_alive = True
                    # shell 虚拟路径（如此电脑）下，LocationURL 可能短暂返回上一次文件路径。
                    # 这里要求同一候选路径连续出现两次再纠正，避免瞬态误判导致路径回跳。
                    if str(current_path).lower().startswith('shell:'):
                        if shell_activation_alive:
                            self._keepalive_candidate_path = None
                            self._keepalive_candidate_count = 0
                        else:
                            candidate = getattr(self, '_keepalive_candidate_path', None)
                            count = int(getattr(self, '_keepalive_candidate_count', 0) or 0)
                            if candidate == local_path:
                                count += 1
                            else:
                                candidate = local_path
                                count = 1
                            self._keepalive_candidate_path = candidate
                            self._keepalive_candidate_count = count
                            if count < 2:
                                debug_print(f"[Keepalive] Candidate path from shell view: {local_path!r} (count={count})")
                                return
                    else:
                        self._keepalive_candidate_path = None
                        self._keepalive_candidate_count = 0

                    debug_print(f"[Keepalive] Navigation detected: {current_path!r} -> {local_path!r}, updating path bar")
                    # 保活轮询仅在无程序化导航（主同步定时器停止）时运行，
                    # 因此此处检测到的差异必为真实用户导航（如双击进入子目录）。
                    # 直接更新路径栏，绕过 _suppress_auto_refresh 门控，
                    # 避免 NavigateComplete2 偶发遗漏时路径栏长时间停留在旧路径。
                    self._keepalive_candidate_path = None
                    self._keepalive_candidate_count = 0
                    self.current_path = local_path
                    if hasattr(self, 'path_bar') and self.path_bar:
                        self.path_bar.set_path(local_path)
                    self._force_details_view_for_local_path(local_path)
                    self.update_tab_title()
                    self._schedule_status_update(track_selection=True)
                    if not getattr(self, '_navigating_programmatically', False) and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    if hasattr(self, '_keepalive_sync_timer') and self._keepalive_sync_timer:
                        self._keepalive_sync_timer.stop()
                    self.start_path_sync_timer(duration_ms=3000)
                else:
                    self._keepalive_candidate_path = None
                    self._keepalive_candidate_count = 0
        except Exception as e:
            debug_print(f"[Keepalive] ERROR: {e}")

    def sync_path_bar_with_explorer(self):
        # 通过QAxWidget的LocationURL属性获取当前路径
        # 若当前标签页不是活跃标签（_refresh_active=False），直接停止同步（避免背景标签持续COM轮询）
        if not getattr(self, '_refresh_active', True):
            self._stop_path_sync_timer()
            return
        # 抗高负载①：卡顿后堆积的爆发式定时器触发直接跳过，避免连续慢 COM 调用进一步压死 UI 线程
        if self._poll_burst_guard():
            return
        # 抗高负载②：高负载退避窗口内拉长轮询间隔并跳过本次 COM 读取，给 UI 线程喘息；
        # 路径仍由 NavigateComplete2 信号与保活轮询兜底更新，窗口到期后自动恢复正常频率。
        if self._com_under_stress():
            t = getattr(self, '_path_sync_timer', None)
            if t and t.interval() < COM_POLL_STRESS_BACKOFF_MS:
                t.setInterval(COM_POLL_STRESS_BACKOFF_MS)
            return
        try:
            url = self._read_location_url_timed()
            if url:
                url_str = str(url)
                local_path = None
                
                # 处理 file:/// 本地路径
                if url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = unquote(url_str[8:])
                    if os.name == 'nt' and local_path.startswith('/'):
                        local_path = local_path[1:]
                    local_path = self._normalize_local_path(local_path)
                # 处理 file://server/share 网络(UNC)路径
                elif url_str.startswith('file://') and not url_str.startswith('file:///'):
                    from urllib.parse import unquote
                    local_path = '\\\\' + unquote(url_str[7:]).replace('/', '\\')
                    local_path = self._normalize_local_path(local_path)
                # 处理 shell: 特殊路径
                elif url_str.startswith('shell:') or '::' in url_str:
                    # 尝试从 CLSID URL 中提取盘符路径（如映射网络驱动器 ::{...}\X:\）
                    import re
                    drive_match = re.search(r'([A-Za-z]:\\[^"]*)', url_str.replace('/', '\\'))
                    if not drive_match:
                        drive_match = re.search(r'([A-Za-z]:/[^"]*)', url_str)
                    if drive_match:
                        local_path = self._normalize_local_path(drive_match.group(1))
                        debug_print(f"[PathSync] Extracted drive path from CLSID URL: {local_path}")
                    else:
                        # 纯 shell: 路径，无法提取盘符，不更新
                        return
                
                current_path = self._normalize_local_path(self.current_path)

                if local_path and local_path != current_path:
                    if self._is_mycomputer_shell_target(current_path):
                        allow_until = float(getattr(self, '_shell_item_activation_until', 0.0) or 0.0)
                        now = time.monotonic()
                        if allow_until <= now:
                            debug_print(
                                f"[PathSync] REJECT shell->local '{local_path}' "
                                f"reason=no-user-activation current='{current_path}'"
                            )
                            return
                        remain_ms = int((allow_until - now) * 1000.0)
                        debug_print(
                            f"[PathSync] ACCEPT shell->local '{local_path}' "
                            f"reason=user-activation remain={remain_ms}ms"
                        )
                        # 仅允许消费一次，避免旧事件连锁放行。
                        self._clear_shell_item_activation("path-sync-accept")
                    # 程序化导航期间（Navigate2 已发出但 Shell.Explorer 尚未完成），
                    # LocationURL 可能仍返回旧路径，此时不回写，避免路径栏倒退。
                    # NavigateComplete2 信号会在导航真正完成后更新路径。
                    if getattr(self, '_suppress_auto_refresh', False):
                        return
                    # 窗口恢复后树面板自动展开的虚假导航，抑制
                    restore_guard = getattr(self, '_restore_guard_until', 0)
                    if restore_guard and time.monotonic() < restore_guard:
                        return
                    self._path_sync_stable_hits = 0
                    self._path_sync_interval_ms = 120
                    if hasattr(self, '_path_sync_timer') and self._path_sync_timer and self._path_sync_timer.interval() != self._path_sync_interval_ms:
                        self._path_sync_timer.setInterval(self._path_sync_interval_ms)
                    self.current_path = local_path
                    if hasattr(self, 'path_bar'):
                        self.path_bar.set_path(local_path)
                    self._force_details_view_for_local_path(local_path)
                    self.update_tab_title()
                    self._schedule_status_update(track_selection=True)
                    # 只在非程序化导航时添加到历史记录
                    if not self._navigating_programmatically and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    # 导航完成后清理标志并恢复同步
                    if getattr(self, '_navigating_folder', False):
                        self._navigating_folder = False
                    self._resume_path_sync_after_navigation()
                elif local_path and local_path == current_path:
                    self._path_sync_stable_hits = int(getattr(self, '_path_sync_stable_hits', 0)) + 1
                    # 即使路径未变化，也做一次轻量UI回写，修复偶发地址栏未重绘。
                    # 优化：若面包屑已存在，只做 repaint 而非全量重建，减少控件析构/创建开销。
                    if hasattr(self, 'path_bar'):
                        pb = self.path_bar
                        if not getattr(pb, '_in_edit', False):
                            # 单控件自绘面包屑：轻量更新分段并同步重绘，无子控件重建开销
                            try:
                                pb.set_path(local_path)
                                pb.repaint()
                            except Exception:
                                pass
                    if hasattr(self, '_path_sync_timer') and self._path_sync_timer:
                        target_interval = 220 if self._path_sync_stable_hits == 1 else 360
                        if self._path_sync_timer.interval() != target_interval:
                            self._path_sync_timer.setInterval(target_interval)
                    if self._path_sync_stable_hits >= 2:
                        # 若停止定时器剩余时间较长（说明调用方设置了长窗口，例如 SelectFile 后的 8s），
                        # 不提前退出——继续以 360ms 低频轮询，等待用户按 Enter 进入子目录。
                        # 若剩余时间较短（正常 2s 窗口快到期），才做提前停止优化。
                        _remaining_ms = 0
                        try:
                            if (hasattr(self, '_path_sync_stop_timer') and self._path_sync_stop_timer
                                    and self._path_sync_stop_timer.isActive()):
                                _remaining_ms = self._path_sync_stop_timer.remainingTime()
                        except Exception:
                            pass
                        if _remaining_ms <= 2000:
                            # 导航已确认完成（Shell.Explorer 稳定在 current_path），
                            # 立即解除抑制标志，避免用户随后双击子目录时路径栏无法更新。
                            self._suppress_auto_refresh = False
                            self._stop_path_sync_timer()
        except Exception as e:
            debug_print(f"[PathSync] ERROR: {e}")

    def _resume_path_sync_after_navigation(self):
        """导航后重启路径同步定时器（短窗口轮询，自动停止）"""
        try:
            self.start_path_sync_timer(duration_ms=2000)
        except Exception as e:
            debug_print(f"[PathSync] ERROR resuming timer: {e}")

    def _schedule_status_update(self, delay_ms=STATUS_UPDATE_DEFER_MS, track_selection=False):
        if track_selection:
            self._start_status_tracking()
        if delay_ms <= 0:
            if self.status_update_timer.isActive():
                self.status_update_timer.stop()
            self.update_explorer_status()
            return
        if self.status_update_timer.isActive() and self.status_update_timer.remainingTime() <= delay_ms:
            return
        self.status_update_timer.start(delay_ms)

    def _start_status_tracking(self, duration_ms=STATUS_TRACKING_WINDOW_MS):
        self._status_tracking_deadline_ms = int(time.time() * 1000) + int(duration_ms)
        if not self.status_tracking_timer.isActive():
            self.status_tracking_timer.start()

    def _stop_status_tracking(self):
        self._status_tracking_deadline_ms = 0
        if self.status_tracking_timer.isActive():
            self.status_tracking_timer.stop()

    def _poll_status_during_interaction(self):
        self.update_explorer_status()
        now_ms = int(time.time() * 1000)
        if now_ms >= int(getattr(self, '_status_tracking_deadline_ms', 0) or 0):
            self._stop_status_tracking()

    def setup_ui(self):
        from PyQt5.QtWidgets import QLabel
        # 设置FileExplorerTab背景为白色
        _theme.bind_style(self, "background: white;")
        self.setAutoFillBackground(True)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 路径栏（极简单行输入框）
        self.path_bar = SimplePathBar(self)
        self.path_bar.pathChanged.connect(self.on_path_bar_changed)
        self.path_bar.activated.connect(self._activate_path_bar)
        layout.addWidget(self.path_bar)

        # 位置不可用提示条（平时隐藏）：导航失败或超时时说明原因，提供重试与重新连接
        self.net_banner = NetworkErrorBanner(self)
        self.net_banner.retryRequested.connect(self._retry_failed_location)
        self.net_banner.reconnectRequested.connect(self._reconnect_failed_location)
        layout.addWidget(self.net_banner)
        
        # 加载指示器（初始隐藏）
        # 注意：改为悬浮覆盖层，不加入垂直布局，避免显示/隐藏时顶部区域高度跳动
        # （以前进度条占用布局高度，show 时路径栏下方多出 20px，看起来像路径栏“变两倍高又恢复”）
        self.loading_bar = QProgressBar(self)
        self.loading_bar.setMaximum(0)  # 不确定进度模式
        self.loading_bar.setTextVisible(True)
        self.loading_bar.setFormat(tr("正在加载大文件夹..."))
        self.loading_bar.setFixedHeight(20)
        self.loading_bar.hide()

        # 优先使用 IExplorerBrowser（真实 Windows 资源管理器外壳，支持 TortoiseGit 图标覆盖）
        # 回退到 Shell.Explorer（IE/WebBrowser ActiveX，不支持 TortoiseGit 图标覆盖）
        _ieb_ok = False
        if _COMTYPES_AVAILABLE:
            try:
                self.explorer = IExplorerBrowserWidget(self)
                _ieb_ok = True
                debug_print("[WindowsShellExplorer] Using IExplorerBrowser (TortoiseGit overlay supported)")
            except Exception as _ieb_err:
                debug_print(f"[WindowsShellExplorer] IExplorerBrowserWidget failed: {_ieb_err}")
        if not _ieb_ok:
            from PyQt5.QAxContainer import QAxWidget
            self.explorer = QAxWidget(self)
            if not self.explorer.setControl("Shell.Explorer"):
                raise RuntimeError("Shell.Explorer control initialization failed")
            debug_print("[WindowsShellExplorer] Using Shell.Explorer (no TortoiseGit overlay support)")
        
        # 设置为NoFocus，防止QAxWidget拦截键盘事件
        self.explorer.setFocusPolicy(Qt.NoFocus)
        # 允许Explorer控件横向压缩，减小右侧面板最小宽度
        try:
            self.explorer.setMinimumWidth(0)
        except Exception:
            pass
        layout.addWidget(self.explorer)
        # 新建的 Shell 视图导航失败时只有一片空白（深色主题下是强烈白块），改显示空占位
        self._explorer_placeholder = QWidget(self)
        self._explorer_placeholder.hide()
        layout.addWidget(self._explorer_placeholder)

        # 状态栏（参考系统 Explorer 样式：细高、浅底色、顶部分割线）
        self.status_bar = QLabel(tr("就绪"))
        self.status_bar.setFixedHeight(20)
        self.status_bar.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        _theme.bind_style(
            self.status_bar,
            "QLabel { padding: 2px 8px; background: white; border-top: 1px solid #e0e0e0; font-size: 12px; color: #444; }"
        )
        # 支持在状态栏上按住拖动整个软件窗口
        self._status_drag_pos = None
        self.status_bar.setCursor(Qt.SizeAllCursor)
        self.status_bar.mousePressEvent = self._status_bar_mouse_press
        self.status_bar.mouseMoveEvent = self._status_bar_mouse_move
        self.status_bar.mouseReleaseEvent = self._status_bar_mouse_release
        self.status_bar.mouseDoubleClickEvent = self._status_bar_mouse_double_click
        # 右侧资源占用标签（CPU/内存），默认隐藏，可在设置中开启
        self.resource_label = QLabel("")
        self.resource_label.setFixedHeight(20)
        self.resource_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        _theme.bind_style(
            self.resource_label,
            "QLabel { padding: 2px 10px; background: white; border-top: 1px solid #e0e0e0; font-size: 12px; }"
        )
        self.resource_label.mouseDoubleClickEvent = self._status_bar_mouse_double_click
        self.resource_label.hide()
        self.cancel_file_op_btn = QPushButton(tr("取消"), self)
        self.cancel_file_op_btn.setFixedHeight(20)
        self.cancel_file_op_btn.setFixedWidth(52)
        _theme.bind_style(
            self.cancel_file_op_btn,
            "QPushButton { background: #f8f8f8; border: 1px solid #dddddd; border-top: 1px solid #e0e0e0;"
            " color: #444; font-size: 11px; padding: 0 6px; }"
            "QPushButton:hover { background: #efefef; }"
            "QPushButton:pressed { background: #e5e5e5; }"
            "QPushButton:disabled { color: #9a9a9a; background: #f4f4f4; }"
        )
        self.cancel_file_op_btn.setToolTip(_hotkey_hint(getattr(getattr(self, 'main_window', None), 'config', {}),
                                                        tr("取消当前后台复制/删除"), 'cancel_file_op'))
        self.cancel_file_op_btn.clicked.connect(self.cancel_current_file_batch_op)
        self.cancel_file_op_btn.hide()
        status_row = QHBoxLayout()
        status_row.setContentsMargins(0, 0, 0, 0)
        status_row.setSpacing(0)
        status_row.addWidget(self.cancel_file_op_btn, 0)
        status_row.addWidget(self.status_bar, 1)
        status_row.addWidget(self.resource_label, 0)
        layout.addLayout(status_row)
        self._bottom_statusbar_visible = True
        self.set_bottom_statusbar_visible(
            bool(getattr(getattr(self, 'main_window', None), 'config', {}).get('show_bottom_statusbar', True))
        )

        
        # 异步加载相关
        self.folder_checker = None  # 文件夹大小检查线程
        self.pending_navigation = None  # 待处理的导航请求

        # 慢盘异步导航状态：_nav_in_progress 表示后台 PIDL 解析未完成（用于显示 loading
        # 与拦截重复点击）；_nav_in_progress_path 记录当前正在解析的目标路径。
        self._nav_in_progress = False
        self._nav_in_progress_path = None
        # 导航完成一致性保护：记录最近一次程序化导航目标，忽略短窗口内的迟到完成事件
        self._expected_nav_path = None
        self._expected_nav_is_shell = False
        self._expected_nav_until = 0.0
        
        # Explorer 基础配置：保留必要项，避免重复 COM 调用拖慢初始化
        self.explorer.dynamicCall('Visible', True)
        self.explorer.dynamicCall('RegisterAsBrowser', True)
        self.explorer.dynamicCall('RegisterAsDropTarget', True)
        self.explorer.dynamicCall('TheaterMode', False)
        self.explorer.dynamicCall('ToolBar', False)
        self.explorer.dynamicCall('StatusBar', False)
        self.explorer.dynamicCall('MenuBar', False)
        self.explorer.dynamicCall('AddressBar', False)
        self.explorer.dynamicCall('Resizable', True)
        self.explorer.dynamicCall('FullScreen', False)
        self.explorer.dynamicCall('Offline', False)
        self.explorer.dynamicCall('Silent', True)
        # 预绑定关键导航事件签名，避免首次导航时事件分发延迟
        self.explorer.dynamicCall('NavigateComplete2(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('DocumentComplete(QVariant,QVariant)', None, None)
        self.explorer.dynamicCall('BeforeNavigate2(QVariant,QVariant,QVariant,QVariant,QVariant,QVariant,QVariant)', None, None, None, None, None, None, None)


        # 兼容原有空白双击（保留控件但不占用空间，避免底部留白）
        self.blank = QLabel()
        self.blank.setFixedHeight(0)
        self.blank.setStyleSheet("background: transparent;")
        self.blank.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.blank.mouseDoubleClickEvent = self.blank_double_click
        # 不再额外增加可见高度
        layout.addWidget(self.blank)

        # 安装事件过滤器以捕获 Explorer 的鼠标按下与双击事件
        try:
            self.explorer.installEventFilter(self)
        except Exception:
            pass

        # 直接连接 NavigateComplete2 信号，实现路径栏即时更新（无 polling 延迟）
        try:
            self.explorer.NavigateComplete2.connect(self._on_shell_navigate_complete)
            debug_print("[Explorer] NavigateComplete2 signal connected")
        except (AttributeError, TypeError) as e:
            debug_print(f"[Explorer] NavigateComplete2 direct signal unavailable: {e}")

        # 连接异步导航（慢盘）状态信号：显示 loading + 进入/解除“导航中”锁定
        try:
            self.explorer.navigationStarted.connect(self._on_async_nav_started)
            self.explorer.navigationFinished.connect(self._on_async_nav_finished)
            self.explorer.navigationFailed.connect(self._on_navigation_failed)
        except (AttributeError, TypeError):
            pass  # Shell.Explorer 回退控件无此信号，忽略

        # 初始设置路径栏（确保路径栏显示初始路径）
        if hasattr(self, 'path_bar'):
            self.path_bar.set_path(self.current_path)
        
        # 启动路径同步定时器
        self.start_path_sync_timer()

        self.update_explorer_status()
        
        # 初始导航到当前路径（在setup_ui最后调用，确保所有设置已应用）
        self.explorer.dynamicCall('Navigate(const QString&)', QDir.toNativeSeparators(self.current_path))

    def set_bottom_statusbar_visible(self, visible):
        """显示/隐藏标签页底部状态栏（状态文本、取消按钮、资源信息）。"""
        visible = bool(visible)
        self._bottom_statusbar_visible = visible
        if hasattr(self, 'status_bar') and self.status_bar:
            self.status_bar.setVisible(visible)
        if hasattr(self, 'cancel_file_op_btn') and self.cancel_file_op_btn:
            if visible and getattr(self, '_file_op_worker', None) and self._file_op_worker.isRunning():
                self.cancel_file_op_btn.show()
            else:
                self.cancel_file_op_btn.hide()
        if hasattr(self, 'resource_label') and self.resource_label:
            if not visible:
                self.resource_label.hide()
                return
            mw = getattr(self, 'main_window', None)
            show_res = bool(mw and mw.config.get("show_resource_usage_in_statusbar", False))
            if show_res and self.resource_label.text().strip():
                self.resource_label.show()
            else:
                self.resource_label.hide()

    def _on_async_nav_started(self, path):
        """慢盘异步导航开始：显示 loading 并进入导航中锁定状态。"""
        self._nav_in_progress = True
        self._nav_in_progress_path = self._normalize_local_path(path)
        self._show_loading_indicator()
        # 安全兜底：网络中断时 NavigateComplete2 可能永不触发，超时后强制解锁并隐藏 loading
        QTimer.singleShot(
            ASYNC_NAV_TIMEOUT_MS,
            lambda p=self._nav_in_progress_path: self._async_nav_safety_timeout(p),
        )

    def _on_async_nav_finished(self, path, ok):
        """慢盘异步导航结束（仅解析失败/无法访问时触发）：解锁并隐藏 loading；失败原因由提示条展示。"""
        self._nav_in_progress = False
        self._hide_loading_indicator()

    def _async_nav_safety_timeout(self, path):
        """导航中锁定的兜底超时：仍在等待同一目标时强制解锁，并提示连接超时。

        后台解析若之后成功，导航完成会自动关闭提示条。"""
        if self._is_cleaning_up:
            return
        if (getattr(self, '_nav_in_progress', False) and
                getattr(self, '_nav_in_progress_path', None) == path):
            debug_print(f"[AsyncNav] Safety timeout, clearing nav lock: {path}")
            self._nav_in_progress = False
            self._hide_loading_indicator()
            self._show_network_error(path, NAV_TIMEOUT)

    def _on_navigation_failed(self, path, hr):
        """导航失败：hr 非 0 为路径解析失败；hr 为 0 是 Shell 内部导航失败，仅对网络位置提示。"""
        # 可能在 Shell 的 COM 回调内触发，延后改布局；其间若有导航成功则忽略本次失败
        seq = getattr(self, '_nav_success_seq', 0)
        QTimer.singleShot(0, lambda: self._handle_navigation_failure(path, hr, seq))

    def _handle_navigation_failure(self, path, hr, seq):
        if self._is_cleaning_up or getattr(self, '_nav_success_seq', 0) != seq:
            return
        if not hr and not is_network_path(path):
            return
        self._nav_in_progress = False
        self._hide_loading_indicator()
        self._show_network_error(path, win32_error(hr) if hr else 0)

    def _on_navigation_succeeded(self):
        self._nav_success_seq = getattr(self, '_nav_success_seq', 0) + 1
        banner = getattr(self, 'net_banner', None)
        placeholder = getattr(self, '_explorer_placeholder', None)
        if (banner is not None and banner.isVisible()) or (placeholder is not None and not placeholder.isHidden()):
            QTimer.singleShot(0, self._dismiss_network_banner)

    def _show_network_error(self, path, code):
        banner = getattr(self, 'net_banner', None)
        if banner is None:
            return
        debug_print(f"[NetBanner] Location unavailable: {path} (code={code})")
        banner.show_failure(path, code, network=is_network_path(path))
        if not getattr(getattr(self, 'explorer', None), '_has_view', True):
            self._set_explorer_placeholder(True)

    def _set_explorer_placeholder(self, active):
        placeholder = getattr(self, '_explorer_placeholder', None)
        if placeholder is None or placeholder.isHidden() != active:
            return
        placeholder.setVisible(active)
        self.explorer.setVisible(not active)

    def _dismiss_network_banner(self):
        banner = getattr(self, 'net_banner', None)
        if banner is not None and banner.isVisible():
            banner.hide()
        self._set_explorer_placeholder(False)

    def _retry_failed_location(self, path):
        self.net_banner.set_busy(tr("正在重试…"))
        self.navigate_to(path)

    def _reconnect_failed_location(self, path):
        if getattr(self, '_reconnect_worker', None) is not None:
            return
        self.net_banner.set_busy(tr("正在连接网络位置…"))
        window = self.window()
        worker = ReconnectWorker(path, int(window.winId()) if window is not None else 0, self)
        worker.completed.connect(self._on_reconnect_finished)
        self._reconnect_worker = worker
        worker.start()
        _retain_thread_until_finished(worker)

    def _on_reconnect_finished(self, path, code):
        self._reconnect_worker = None
        banner = getattr(self, 'net_banner', None)
        # 期间已导航到别处或关闭了提示条：不再把用户带回该位置
        if self._is_cleaning_up or banner is None or not banner.isVisible() or banner.path != path:
            return
        if code in RECONNECTED_CODES:
            banner.set_busy(tr("已连接，正在打开…"))
            self.navigate_to(path)
        elif code == ERROR_CANCELLED:
            banner.show_failure(path, banner.code, network=True, note=tr("已取消连接。"))
        else:
            self._show_network_error(path, code)

    def _clear_async_nav_lock(self):
        """导航真正完成（NavigateComplete2）时解除导航中锁定并隐藏 loading。"""
        if getattr(self, '_nav_in_progress', False):
            self._nav_in_progress = False
            self._hide_loading_indicator()

    def _mark_expected_navigation(self, path, is_shell):
        """记录短时导航期望，用于过滤乱序/迟到的 NavigateComplete2。"""
        try:
            self._expected_nav_path = str(path or '')
            self._expected_nav_is_shell = bool(is_shell)
            self._expected_nav_until = time.monotonic() + 3.0
        except Exception:
            self._expected_nav_path = None
            self._expected_nav_is_shell = False
            self._expected_nav_until = 0.0

    def _mark_shell_item_activation(self, ttl_sec=2.5):
        """标记用户刚在“此电脑”内激活了一个条目（通常是盘符）。"""
        try:
            self._shell_item_activation_until = time.monotonic() + float(ttl_sec)
            remain_ms = int(max(0.0, (self._shell_item_activation_until - time.monotonic()) * 1000.0))
            debug_print(f"[ShellActivation] armed ttl={remain_ms}ms")
        except Exception:
            self._shell_item_activation_until = 0.0
            debug_print("[ShellActivation] arm failed, reset to 0")

    def _clear_shell_item_activation(self, reason=""):
        """清理“此电脑条目激活”许可窗口。"""
        self._shell_item_activation_until = 0.0
        if reason:
            debug_print(f"[ShellActivation] cleared reason={reason}")

    def _should_accept_nav_completion(self, local_path, raw_url):
        """判断本次导航完成事件是否应被接收。"""
        exp = getattr(self, '_expected_nav_path', None)
        if not exp:
            return True

        until = float(getattr(self, '_expected_nav_until', 0.0) or 0.0)
        now = time.monotonic()
        if until > 0 and now > until:
            self._expected_nav_path = None
            self._expected_nav_is_shell = False
            self._expected_nav_until = 0.0
            return True

        exp_is_shell = bool(getattr(self, '_expected_nav_is_shell', False))
        raw = str(raw_url or '')
        if exp_is_shell:
            # 期望是 shell 目标时，短窗口内拒绝 file 路径完成，避免“此电脑 -> D:\”回跳。
            if local_path is not None:
                if self._is_mycomputer_shell_target(exp):
                    allow_until = float(getattr(self, '_shell_item_activation_until', 0.0) or 0.0)
                    now = time.monotonic()
                    if allow_until > now:
                        remain_ms = int((allow_until - now) * 1000.0)
                        debug_print(
                            f"[ShellActivation] ACCEPT shell->local '{local_path}' "
                            f"reason=user-item-activation remain={remain_ms}ms exp='{exp}' raw='{raw}'"
                        )
                        self._expected_nav_path = None
                        self._expected_nav_is_shell = False
                        self._expected_nav_until = 0.0
                        self._clear_shell_item_activation("accept-shell-to-local")
                        return True
                    debug_print(
                        f"[ShellActivation] REJECT shell->local '{local_path}' "
                        f"reason=activation-expired exp='{exp}' raw='{raw}'"
                    )
                debug_print(f"[NavigateComplete2] Ignored stale file completion during shell target: {local_path}")
                return False
            if raw.startswith('shell:') or '::' in raw:
                debug_print(f"[ShellActivation] keep shell completion accepted raw='{raw}'")
                self._expected_nav_path = None
                self._expected_nav_is_shell = False
                self._expected_nav_until = 0.0
            return True

        if local_path is None:
            debug_print(f"[NavigateComplete2] Ignored non-file completion while expecting file path: {exp}")
            return False

        expected_path = self._normalize_local_path(exp)
        actual_path = self._normalize_local_path(local_path)
        if actual_path == expected_path:
            self._expected_nav_path = None
            self._expected_nav_is_shell = False
            self._expected_nav_until = 0.0
            return True

        debug_print(
            f"[NavigateComplete2] Ignored stale completion: got {actual_path}, expected {expected_path}"
        )
        return False

    def _on_shell_navigate_complete(self, *args):
        """Shell.Explorer NavigateComplete2 直接信号处理：路径栏即时更新"""
        try:
            # NavigateComplete2 签名: (IDispatch* pDisp, VARIANT* URL)
            # PyQt5 传入参数可能是 (dispatch, url) 或仅 (url,)
            url = None
            for arg in args:
                s = str(arg)
                if s.startswith('file:///') or s.startswith('file://') or s.startswith('shell:') or '::' in s:
                    url = s
                    break
            if url is None and args:
                url = str(args[-1])
            if not url:
                return
            # 每次进入“此电脑”都强制切换到平铺视图。
            if self._is_mycomputer_shell_target(url):
                self._force_tile_view_for_this_pc()
            local_path = None
            if url.startswith('file:'):
                # 统一解析 file: URL，正确还原本地盘符路径与 UNC 网络路径。
                # UNC 路径经 _on_nav_complete 编码后形如 file://///server/share，
                # 直接剥离固定前缀会丢失一个反斜杠，破坏 UNC 前缀，故按前导斜杠数判定。
                from urllib.parse import unquote
                rest = unquote(url[5:]).replace('/', '\\')
                stripped = rest.lstrip('\\')
                if len(stripped) >= 2 and stripped[1] == ':':
                    # 本地盘符路径，如 C:\...
                    local_path = self._normalize_local_path(stripped)
                elif stripped:
                    # UNC 网络路径，如 \\server\share\...
                    local_path = self._normalize_local_path('\\\\' + stripped)
            # 处理 CLSID 路径（如 ::{20D04FE0-...}\X:\）：尝试提取盘符路径
            if local_path is None and '::' in url:
                import re
                drive_match = re.search(r'([A-Za-z]:\\[^"]*)', url.replace('/', '\\'))
                if not drive_match:
                    drive_match = re.search(r'([A-Za-z]:/[^"]*)', url)
                if drive_match:
                    local_path = self._normalize_local_path(drive_match.group(1))
                    debug_print(f"[NavigateComplete2] Extracted drive path from CLSID URL: {local_path}")
            if local_path is not None:
                # 保护“此电脑”场景：没有明确用户激活时，不接受 shell->盘符的意外完成。
                current_before = self._normalize_local_path(getattr(self, 'current_path', ''))
                if self._is_mycomputer_shell_target(current_before):
                    exp_path = getattr(self, '_expected_nav_path', None)
                    exp_is_shell = bool(getattr(self, '_expected_nav_is_shell', False))
                    expecting_local_target = bool(exp_path) and (not exp_is_shell)
                    if not expecting_local_target:
                        allow_until = float(getattr(self, '_shell_item_activation_until', 0.0) or 0.0)
                        now = time.monotonic()
                        if allow_until <= now:
                            debug_print(
                                f"[ShellActivation] REJECT nav-complete shell->local '{local_path}' "
                                f"reason=no-user-activation current='{current_before}'"
                            )
                            return
                        remain_ms = int((allow_until - now) * 1000.0)
                        debug_print(
                            f"[ShellActivation] ACCEPT nav-complete shell->local '{local_path}' "
                            f"reason=user-item-activation remain={remain_ms}ms"
                        )
                        # 已确认用户意图：清空 shell 期望，避免后续 _should_accept 再次误拒绝。
                        self._expected_nav_path = None
                        self._expected_nav_is_shell = False
                        self._expected_nav_until = 0.0
                if not self._should_accept_nav_completion(local_path, url):
                    return
                current = self._normalize_local_path(getattr(self, 'current_path', ''))
                # 抑制窗口恢复后 IEB 树面板自动展开导致的虚假导航
                # （树面板可能自动展开到之前记忆的子目录，导致 NavigateComplete2 连续触发错误路径）
                restore_guard = getattr(self, '_restore_guard_until', 0)
                if (restore_guard and time.monotonic() < restore_guard and
                        local_path != current and current and
                        not current.startswith('shell:')):
                    # 虚假导航：忽略并强制 IEB 回到正确路径
                    debug_print(f"[NavigateComplete2] Suppressed spurious post-restore nav: {local_path}")
                    if hasattr(self, 'explorer') and hasattr(self.explorer, '_navigate'):
                        self._restore_guard_until = 0  # 防止无限循环
                        self.explorer._navigate(current)
                    return
                # 无条件更新路径栏（即使路径未变也刷新显示）
                if hasattr(self, 'path_bar'):
                    self.path_bar.set_path(local_path)
                # 导航真正完成：解除慢盘异步导航的“导航中”锁定并隐藏 loading
                self._clear_async_nav_lock()
                self._on_navigation_succeeded()
                self._force_details_view_for_local_path(local_path)
                # 延迟安装/更新 IExplorerBrowser 双击钩子（SysListView32 在首次导航后才创建）
                QTimer.singleShot(200, self._install_listview_dblclick_hook)
                if local_path and local_path != current:
                    self.current_path = local_path
                    self.update_tab_title()
                    self._schedule_status_update(track_selection=True)
                    if not getattr(self, '_navigating_programmatically', False) and hasattr(self, '_add_to_history'):
                        self._add_to_history(local_path)
                    if getattr(self, '_navigating_folder', False):
                        self._navigating_folder = False
                    # 记录导航时间戳，用于抑制 WH_MOUSE_LL 双击误判
                    self._last_nav_complete_time = time.monotonic()
                    if self.main_window and hasattr(self.main_window, 'update_chat_context'):
                        try:
                            if self.main_window.get_current_tab_widget() is self:
                                self.main_window.update_chat_context()
                        except Exception:
                            pass
                    self._clear_shell_item_activation("nav-complete-committed")
                    debug_print(f"[NavigateComplete2] Path updated: {local_path}")
        except Exception as ex:
            debug_print(f"[NavigateComplete2] Error: {ex}")

    def _is_mycomputer_shell_target(self, raw_path_or_url):
        """判断目标是否为“此电脑”（shell:MyComputerFolder）。"""
        s = str(raw_path_or_url or '').strip().lower()
        if not s:
            return False
        if s == 'shell:mycomputerfolder':
            return True
        # This PC CLSID
        return '{20d04fe0-3aea-1069-a2d8-08002b30309d}' in s

    def _set_ieb_current_view_mode(self, view_mode):
        """通过 IFolderView::SetCurrentViewMode 设置 IExplorerBrowser 视图模式。"""
        try:
            if not isinstance(self.explorer, IExplorerBrowserWidget):
                return False
            if not getattr(self.explorer, '_browser', None):
                return False

            from comtypes import GUID as _GUID
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_sv = self.explorer._browser.GetCurrentView(ctypes.byref(iid_sv))
            sv_ptr = int(ppv_sv) if ppv_sv else 0
            if not sv_ptr:
                return False

            _vp_size = ctypes.sizeof(ctypes.c_void_p)
            try:
                # QI(IShellView -> IFolderView)
                iid_fv = _GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                qi_addr = ctypes.c_void_p.from_address(vtable_ptr).value
                _QI = ctypes.WINFUNCTYPE(
                    ctypes.c_long,
                    ctypes.c_void_p,
                    ctypes.POINTER(type(iid_fv)),
                    ctypes.POINTER(ctypes.c_void_p),
                )
                qi_fn = _QI(qi_addr)

                fv_ptr = ctypes.c_void_p(0)
                hr_qi = qi_fn(sv_ptr, ctypes.byref(iid_fv), ctypes.byref(fv_ptr))
                if hr_qi != 0 or not fv_ptr.value:
                    return False

                try:
                    # IFolderView vtable: QI(0), AddRef(1), Release(2),
                    # GetCurrentViewMode(3), SetCurrentViewMode(4)
                    fv_vtbl = ctypes.c_void_p.from_address(fv_ptr.value).value
                    set_mode_addr = ctypes.c_void_p.from_address(fv_vtbl + 4 * _vp_size).value
                    _SET_VIEW_MODE = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_uint)
                    hr_set = _SET_VIEW_MODE(set_mode_addr)(fv_ptr.value, int(view_mode))
                    return hr_set == 0
                finally:
                    rel_addr = ctypes.c_void_p.from_address(fv_vtbl + 2 * _vp_size).value
                    _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                    _REL(rel_addr)(fv_ptr.value)
            finally:
                sv_vtbl = ctypes.c_void_p.from_address(sv_ptr).value
                rel_addr = ctypes.c_void_p.from_address(sv_vtbl + 2 * _vp_size).value
                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _REL(rel_addr)(sv_ptr)
        except Exception as e:
            debug_print(f"[ViewMode] _set_ieb_current_view_mode error: {e}")
            return False

    def _force_tile_view_for_this_pc(self):
        """将“此电脑”视图强制切换到平铺（FVM_TILE=6）。"""
        FVM_TILE = 6
        delays_ms = (0, 140, 320, 620, 980, 1450)

        def _try_force():
            ok = self._set_ieb_current_view_mode(FVM_TILE)
            debug_print(f"[ViewMode] force tile attempt ok={ok}")

        for d in delays_ms:
            QTimer.singleShot(d, _try_force)

    def _force_details_view_for_local_path(self, local_path):
        """进入本地/UNC目录时强制详细视图（FVM_DETAILS=4）。"""
        try:
            p = self._normalize_local_path(local_path)
            s = str(p or '')
            if not s:
                return
            # 仅对真实文件系统路径生效；shell:MyComputerFolder 仍保持平铺模式。
            if s.startswith('shell:') or '::' in s:
                return
            is_drive = len(s) >= 3 and s[1:3] in (':\\', ':/')
            is_unc = s.startswith('\\\\') or s.startswith('//')
            if not (is_drive or is_unc):
                return
        except Exception:
            return

        FVM_DETAILS = 4
        delays_ms = (0, 120, 280, 520)

        def _try_force():
            ok = self._set_ieb_current_view_mode(FVM_DETAILS)
            debug_print(f"[ViewMode] force details attempt ok={ok} path='{local_path}'")

        for d in delays_ms:
            QTimer.singleShot(d, _try_force)

    def event(self, e):
        # 捕获QAxWidget的NavigateComplete2事件（MetaCall备用通道，type=43）
        if e.type() == 43:  # QEvent.MetaCall
            if hasattr(e, 'arguments') and hasattr(e, 'signal'):
                if 'NavigateComplete2' in str(e.signal):
                    url = str(e.arguments[1])
                    # 检查是否为控制面板及其子目录，若是则用原生窗口打开
                    if self._is_control_panel_path(url):
                        try:
                            import subprocess
                            launch_detached(['explorer.exe', url])
                            show_toast(self, tr("已打开"), tr("控制面板已在新窗口打开"), level="info", duration=2000)
                        except Exception as ex:
                            show_toast(self, tr("错误"), tr("无法打开控制面板: {}").format(ex), level="error")
                        # 关闭当前标签页（定位其所属标签组，左右分屏均可）
                        mw = self.main_window
                        if mw and hasattr(mw, '_all_groups'):
                            for cand_tw, cand_cs in mw._all_groups():
                                if cand_cs is not None:
                                    j = cand_cs.indexOf(self)
                                    if j != -1:
                                        mw.close_tab(j, target_tabwidget=cand_tw)
                                        break
                        elif mw and hasattr(mw, 'tab_widget'):
                            idx = mw.tab_widget.indexOf(self)
                            if idx != -1:
                                mw.close_tab(idx)
                        return True
                    if url.startswith('file:///'):
                        from urllib.parse import unquote
                        local_path = unquote(url[8:])
                        if os.name == 'nt' and local_path.startswith('/'):
                            local_path = local_path[1:]
                        local_path = self._normalize_local_path(local_path)
                    elif url.startswith('file://') and not url.startswith('file:///'):
                        from urllib.parse import unquote
                        local_path = self._normalize_local_path('\\\\' + unquote(url[7:]).replace('/', '\\'))
                    else:
                        local_path = None
                    if local_path is not None:
                        if not self._should_accept_nav_completion(local_path, url):
                            return True
                        current = self._normalize_local_path(getattr(self, 'current_path', ''))
                        if local_path != current:
                            self.current_path = local_path
                        if hasattr(self, 'path_bar'):
                            self.path_bar.set_path(local_path)
                        self._schedule_status_update(track_selection=True)
                        # 导航完成，清除导航标志（hit-test在双击时已完成判断，100ms后清除即可）
                        if hasattr(self, '_navigating_folder') and self._navigating_folder:
                            from PyQt5.QtCore import QTimer
                            QTimer.singleShot(100, lambda: setattr(self, '_navigating_folder', False))
                        if self.main_window and hasattr(self.main_window, 'update_chat_context'):
                            try:
                                if self.main_window.get_current_tab_widget() is self:
                                    self.main_window.update_chat_context()
                            except Exception:
                                pass
        return super().event(e)

    def open_tortoisegit_log(self):
        """打开 TortoiseGit 日志查看器"""
        try:
            current_path = self.current_path
            if not current_path or not os.path.exists(current_path):
                show_toast(self, tr("提示"), tr("当前路径无效"), level="warning")
                return

            repo_root = self._find_git_root(current_path)
            if not repo_root:
                show_toast(self, tr("提示"), tr("当前目录不是 Git 仓库，未找到 .git"), level="warning")
                return
            
            # TortoiseGit 命令行：TortoiseGitProc.exe /command:log /path:"路径"
            # 尝试找到 TortoiseGitProc.exe
            tortoisegit_paths = [
                r"C:\Program Files\TortoiseGit\bin\TortoiseGitProc.exe",
                r"C:\Program Files (x86)\TortoiseGit\bin\TortoiseGitProc.exe",
            ]
            
            tortoisegit_exe = None
            for path in tortoisegit_paths:
                if os.path.exists(path):
                    tortoisegit_exe = path
                    break
            
            if not tortoisegit_exe:
                show_toast(
                    self,
                    tr("提示"),
                    tr("未找到 TortoiseGit，请确认已安装 TortoiseGit\n下载地址: https://tortoisegit.org/download/"),
                    level="warning",
                )
                return
            
            # 启动 TortoiseGit Log
            launch_detached_async([tortoisegit_exe, '/command:log', f'/path:{repo_root}'])
            debug_print(f"[TortoiseGit] Opened log for: {repo_root}")
            
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开 TortoiseGit Log: {}").format(e), level="error")
            debug_print(f"[TortoiseGit] Failed to open log: {e}")
    
    def open_tortoisegit_commit(self):
        """打开 TortoiseGit 提交窗口"""
        try:
            current_path = self.current_path
            if not current_path or not os.path.exists(current_path):
                show_toast(self, tr("提示"), tr("当前路径无效"), level="warning")
                return

            repo_root = self._find_git_root(current_path)
            if not repo_root:
                show_toast(self, tr("提示"), tr("当前目录不是 Git 仓库，未找到 .git"), level="warning")
                return
            
            # TortoiseGit 命令行：TortoiseGitProc.exe /command:commit /path:"路径"
            tortoisegit_paths = [
                r"C:\Program Files\TortoiseGit\bin\TortoiseGitProc.exe",
                r"C:\Program Files (x86)\TortoiseGit\bin\TortoiseGitProc.exe",
            ]
            
            tortoisegit_exe = None
            for path in tortoisegit_paths:
                if os.path.exists(path):
                    tortoisegit_exe = path
                    break
            
            if not tortoisegit_exe:
                show_toast(
                    self,
                    tr("提示"),
                    tr("未找到 TortoiseGit，请确认已安装 TortoiseGit\n下载地址: https://tortoisegit.org/download/"),
                    level="warning",
                )
                return
            
            # 启动 TortoiseGit Commit
            launch_detached_async([tortoisegit_exe, '/command:commit', f'/path:{repo_root}'])
            debug_print(f"[TortoiseGit] Opened commit for: {repo_root}")
            
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开 TortoiseGit Commit: {}").format(e), level="error")
            debug_print(f"[TortoiseGit] Failed to open commit: {e}")

    def _find_git_root(self, start_path):
        """向上查找包含 .git 的目录，找到则返回仓库根路径，否则返回 None"""
        if not start_path:
            return None
        path = os.path.abspath(start_path)
        while True:
            git_marker = os.path.join(path, '.git')
            if os.path.isdir(git_marker):
                return path
            if os.path.isfile(git_marker):
                try:
                    with open(git_marker, 'r', encoding='utf-8', errors='ignore') as f:
                        line = f.readline().strip()
                    if line.lower().startswith('gitdir:'):
                        gitdir_path = line[7:].strip()
                        if not os.path.isabs(gitdir_path):
                            gitdir_path = os.path.abspath(os.path.join(path, gitdir_path))
                        if os.path.exists(gitdir_path):
                            return path
                except Exception:
                    pass
            parent = os.path.dirname(path)
            if parent == path:
                break
            path = parent
        return None

    def _request_git_status_async(self, dir_path):
        """异步请求 Git 状态（不阻塞 UI 线程）"""
        if not dir_path or dir_path.startswith('shell:') or '::' in dir_path:
            return
        # 慢盘（网络/UNC/映射盘）：_find_git_root 向上逐级 os.path.isdir/isfile 探测，
        # 在挂起的网络路径上会同步阻塞 UI 线程导致卡死。网络盘一般非 Git 仓库，直接跳过。
        if self._is_slow_path(dir_path):
            self._git_status_cache = {'path': dir_path, 'result': None,
                                      'ts_ms': int(time.time() * 1000)}
            return
        # 防止重复请求同一路径
        if getattr(self, '_git_status_pending_path', None) == dir_path:
            return
        cache = getattr(self, '_git_status_cache', None)
        now_ms = int(time.time() * 1000)
        if cache and cache.get('path') == dir_path and (now_ms - cache.get('ts_ms', 0)) < 5000:
            return  # 缓存仍有效，无需重新查询
        repo_root = self._find_git_root(dir_path)
        if not repo_root:
            self._git_status_cache = {'path': dir_path, 'result': None, 'ts_ms': now_ms}
            return
        git_exe = 'git.exe'
        git_root = find_git_install_root()
        if git_root:
            candidate = os.path.join(git_root, 'cmd', 'git.exe')
            if os.path.isfile(candidate):
                git_exe = candidate
        self._git_status_pending_path = dir_path
        worker = GitStatusWorker(dir_path, repo_root, git_exe, parent=self)
        worker.completed.connect(self._on_git_status_finished)
        self._git_status_worker = worker  # prevent GC
        worker.start()
        _retain_thread_until_finished(worker)

    def _on_git_status_finished(self, dir_path, repo_root, summary):
        """Git 状态查询完成回调（主线程）"""
        now_ms = int(time.time() * 1000)
        self._git_status_cache = {'path': dir_path, 'result': summary, 'ts_ms': now_ms}
        if getattr(self, '_git_status_pending_path', None) == dir_path:
            self._git_status_pending_path = None
        branch = getattr(self.sender(), 'branch', '') if summary else ''
        if (dir_path, branch) != getattr(self, '_git_branch_info', ('', '')):
            self._git_branch_info = (dir_path, branch)
            mw = getattr(self, 'main_window', None)
            if mw is not None and hasattr(mw, '_schedule_tab_labels'):
                mw._schedule_tab_labels()
        # 刷新状态栏（仅当当前路径匹配时）
        if getattr(self, 'current_path', None) == dir_path:
            self.update_explorer_status()

    def on_path_bar_changed(self, path):
        """处理面包屑路径栏的路径变化，支持特殊shell路径自动跳转"""
        import os
        path = path.strip()

        # 处理 file: URL（从邮件/浏览器/SharePoint等复制的链接）
        # 支持 file:///C:/path、file://server/share、file:\\server\share、file:\server\share 等格式
        if path.lower().startswith('file:'):
            path = _file_url_to_path(path)
            debug_print(f"[PathBar] file: URL converted to: {path}")

        lower_path = path.lower()
        debug_print(f"[PathBar] on_path_bar_changed: '{path}'")

        if lower_path in ('terminal', 'term'):
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    preferred_tool = 'cmd'
                    if self.main_window and hasattr(self.main_window, 'get_preferred_terminal_tool'):
                        preferred_tool = self.main_window.get_preferred_terminal_tool()
                    launch_shell_tool(preferred_tool, current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开默认终端"), level="error")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开默认终端: {}").format(e), level="error")
            return

        # 处理cmd命令
        if lower_path == 'cmd':
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    launch_shell_tool('cmd', current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开命令行"), level="error")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开命令行: {}").format(e), level="error")
            return

        if lower_path in ('powershell', 'ps', 'pwsh'):
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    launch_shell_tool('powershell', current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开 PowerShell"), level="error")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开 PowerShell: {}").format(e), level="error")
            return

        if lower_path in ('gitbash', 'git-bash', 'bash'):
            try:
                current_dir = self.current_path
                if current_dir and os.path.exists(current_dir):
                    launch_shell_tool('git-bash', current_dir)
                    self.path_bar.set_path(current_dir)
                else:
                    show_toast(self, tr("错误"), tr("当前路径无效，无法打开 Git Bash"), level="error")
            except FileNotFoundError as e:
                show_toast(self, tr("提示"), str(e), level="warning")
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开 Git Bash: {}").format(e), level="error")
            return

        # 支持特殊shell路径映射
        special_map = {
            tr('回收站'): 'shell:RecycleBinFolder',
            tr('此电脑'): 'shell:MyComputerFolder',
            tr('我的电脑'): 'shell:MyComputerFolder',
            '桌面': 'shell:Desktop',
            tr('网络'): 'shell:NetworkPlacesFolder',
            tr('启动项'): 'shell:Startup',
            tr('开机启动项'): 'shell:Startup',
            tr('启动文件夹'): 'shell:Startup',
            'Startup': 'shell:Startup',
            'OneDrive': 'shell:OneDrive',
            'onedrive': 'shell:OneDrive',
        }
        # shell:Startup 等特殊路径，自动解析为真实系统路径
        shell_path_map = {
            'shell:Startup': lambda: os.path.join(os.environ.get('APPDATA', ''), r'Microsoft\Windows\Start Menu\Programs\Startup'),
            'shell:OneDrive': lambda: os.environ.get('OneDrive', ''),
        }

        if path in special_map:
            shell_path = special_map[path]
            # 如果是 shell:Startup，自动跳转到真实路径
            if shell_path in shell_path_map:
                real_path = shell_path_map[shell_path]()
                if os.path.exists(real_path):
                    self.navigate_to(real_path)
                else:
                    show_toast(self, tr("路径错误"), tr("启动文件夹不存在: {}").format(real_path), level="warning")
                    if hasattr(self, 'current_path') and self.current_path:
                        self.path_bar.set_path(self.current_path)
                return
            else:
                self.navigate_to(shell_path, is_shell=True)
                return

        # 允许直接输入 shell:XXX 跳转
        if path.lower().startswith('shell:'):
            # shell:Startup 也做特殊处理
            if path.lower() == 'shell:startup' and 'shell:Startup' in shell_path_map:
                real_path = shell_path_map['shell:Startup']()
                if os.path.exists(real_path):
                    self.navigate_to(real_path)
                else:
                    show_toast(self, tr("路径错误"), tr("启动文件夹不存在: {}").format(real_path), level="warning")
                    if hasattr(self, 'current_path') and self.current_path:
                        self.path_bar.set_path(self.current_path)
                return
            # shell:OneDrive 解析为真实路径
            if path.lower() == 'shell:onedrive' and 'shell:OneDrive' in shell_path_map:
                real_path = shell_path_map['shell:OneDrive']()
                if real_path and os.path.exists(real_path):
                    self.navigate_to(real_path)
                else:
                    show_toast(self, tr("路径错误"), tr("未找到 OneDrive 文件夹"), level="warning")
                    if hasattr(self, 'current_path') and self.current_path:
                        self.path_bar.set_path(self.current_path)
                return
            self.navigate_to(path, is_shell=True)
            return

        # 路径与当前路径相同时（如 resizeEvent 误触发 pathChanged），跳过导航，防止循环刷新
        current_norm = getattr(self, 'current_path', '').replace('/', '\\').lower().rstrip('\\')
        path_norm = path.replace('/', '\\').lower().rstrip('\\')
        if path_norm == current_norm:
            debug_print(f"[PathBar] on_path_bar_changed: same as current_path, skip navigate")
            return

        # UNC 路径不做中英文目录互转：互转会同步探测路径是否存在，服务器不可达时长时间卡住界面
        if path.startswith('\\\\') or path.startswith('//'):
            self.navigate_to(path)
            return

        # 尝试中英文目录互转
        path2 = translate_common_path(path)
        # 慢盘（网络/UNC/映射盘/OneDrive）：os.path.exists 在挂起路径上会阻塞 UI 线程，
        # 直接交给 navigate_to（其内部对慢盘走后台异步解析，不做同步文件系统探测）。
        if self._is_slow_path(path2):
            self.navigate_to(path2)
        elif os.path.exists(path2) or is_network_path(path2):
            # 已记住但未连接的映射盘也交给 navigate_to，由其显示重新连接提示
            self.navigate_to(path2)
        else:
            show_toast(self, tr("路径错误"), tr("路径不存在: {}").format(path2), level="warning")
            if hasattr(self, 'current_path') and self.current_path:
                self.path_bar.set_path(self.current_path)

    def explorer_mouse_press(self, event):
        # 在鼠标按下时记录当时的选中项数量和鼠标位置点击测试结果
        try:
            cnt = self._get_selected_count_safe()
            self._selected_before_click = int(cnt) if cnt is not None else None
            
            # 使用原生点击测试来判断是否点击在项目上（更准确）
            if HAS_PYWIN:
                try:
                    from PyQt5.QtGui import QCursor
                    gx = QCursor.pos().x()
                    gy = QCursor.pos().y()
                    self._clicked_on_item = self._native_listview_hit_test(gx, gy)
                except Exception:
                    self._clicked_on_item = False
            else:
                self._clicked_on_item = False
                
        except Exception:
            self._selected_before_click = None
            self._clicked_on_item = False
        # 继续默认处理（不阻止控件行为）
        # 直接返回 None — 不尝试调用 ActiveX 的原始处理（事件仍会被控件处理）
        return None

    def _get_selected_count_safe(self):
        """安全地获取当前选中项数量，避免触发 ActiveX 属性不存在的警告"""
        try:
            # IExplorerBrowser 模式：使用 IFolderView COM 接口
            if isinstance(self.explorer, IExplorerBrowserWidget) and getattr(self.explorer, '_browser', None):
                return self._get_ieb_selection_count()

            # 优先通过 Document 接口获取 SelectedItems（避免 WebBrowser 直接调用警告）
            doc = None
            try:
                doc = self.explorer.querySubObject('Document') if hasattr(self, 'explorer') else None
            except Exception:
                doc = None

            sel = None
            if doc:
                try:
                    sel = doc.querySubObject('SelectedItems()')
                except Exception:
                    sel = None

            if sel is None:
                return None

            count = None
            try:
                if hasattr(sel, 'property'):
                    count = sel.property('Count')
            except Exception:
                count = None

            if count is None:
                try:
                    count = len(sel)
                except Exception:
                    count = None

            return int(count) if count is not None else None
        except Exception:
            return None

    def _get_ieb_selection_count(self):
        """通过 IFolderView COM 接口获取 IExplorerBrowser 的选中项数量。
        先获取 IShellView（已验证可用），再 QueryInterface 获取 IFolderView。"""
        try:
            from comtypes import GUID as _GUID
            # 先获取 IShellView（已验证此路径可靠工作）
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_sv = self.explorer._browser.GetCurrentView(ctypes.byref(iid_sv))
            sv_ptr = int(ppv_sv) if ppv_sv else 0
            if not sv_ptr:
                return None
            try:
                _vp_size = ctypes.sizeof(ctypes.c_void_p)
                # 通过 IShellView 的 QueryInterface 获取 IFolderView
                iid_fv = _GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                qi_addr = ctypes.c_void_p.from_address(vtable_ptr).value  # QI at index 0
                _QI = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                         ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
                qi_fn = _QI(qi_addr)
                fv_ptr = ctypes.c_void_p(0)
                hr_qi = qi_fn(sv_ptr, ctypes.byref(iid_fv), ctypes.byref(fv_ptr))
                if hr_qi != 0 or not fv_ptr.value:
                    return None
                try:
                    # IFolderView vtable (inherits IUnknown directly):
                    # QI(0), AddRef(1), Release(2), GetCurrentViewMode(3),
                    # SetCurrentViewMode(4), GetFolder(5), Item(6), ItemCount(7)
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value
                    fn_addr = ctypes.c_void_p.from_address(fv_vtable + 7 * _vp_size).value
                    _IC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                             ctypes.c_uint, ctypes.POINTER(ctypes.c_int))
                    item_count_fn = _IC(fn_addr)
                    count = ctypes.c_int(0)
                    hr = item_count_fn(fv_ptr.value, 0x1, ctypes.byref(count))  # SVGIO_SELECTION=1
                    if hr == 0:
                        return count.value
                finally:
                    # Release IFolderView
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value
                    rel_addr = ctypes.c_void_p.from_address(fv_vtable + 2 * _vp_size).value
                    _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                    _REL(rel_addr)(fv_ptr.value)
            finally:
                # Release IShellView
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                rel_addr = ctypes.c_void_p.from_address(vtable_ptr + 2 * _vp_size).value
                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _REL(rel_addr)(sv_ptr)
        except Exception as e:
            debug_print(f"[IEB] _get_ieb_selection_count failed: {e}")
        return None

    def _get_ieb_selected_paths(self):
        """通过 IFolderView::Items(SVGIO_SELECTION) 获取 IExplorerBrowser 选中文件路径列表。"""
        paths = []
        try:
            from comtypes import GUID as _GUID
            _vp_size = ctypes.sizeof(ctypes.c_void_p)
            # 获取 IShellView
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_sv = self.explorer._browser.GetCurrentView(ctypes.byref(iid_sv))
            sv_ptr = int(ppv_sv) if ppv_sv else 0
            if not sv_ptr:
                return paths
            try:
                # QI for IFolderView
                iid_fv = _GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                qi_addr = ctypes.c_void_p.from_address(vtable_ptr).value
                _QI = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                         ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
                qi_fn = _QI(qi_addr)
                fv_ptr = ctypes.c_void_p(0)
                hr_qi = qi_fn(sv_ptr, ctypes.byref(iid_fv), ctypes.byref(fv_ptr))
                if hr_qi != 0 or not fv_ptr.value:
                    return paths
                try:
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value
                    # IFolderView::Items at vtable index 8
                    # HRESULT Items(UINT uFlags, REFIID riid, void** ppv)
                    iid_sia = _GUID("{B63EA76D-1F85-456F-A19C-48159EFA858B}")  # IShellItemArray
                    fn_addr = ctypes.c_void_p.from_address(fv_vtable + 8 * _vp_size).value
                    _ITEMS = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                ctypes.c_uint, ctypes.POINTER(_GUID),
                                                ctypes.POINTER(ctypes.c_void_p))
                    items_fn = _ITEMS(fn_addr)
                    sia_ptr = ctypes.c_void_p(0)
                    hr = items_fn(fv_ptr.value, 0x1, ctypes.byref(iid_sia), ctypes.byref(sia_ptr))
                    if hr != 0 or not sia_ptr.value:
                        return paths
                    try:
                        # IShellItemArray::GetCount at vtable index 7
                        sia_vtable = ctypes.c_void_p.from_address(sia_ptr.value).value
                        gc_addr = ctypes.c_void_p.from_address(sia_vtable + 7 * _vp_size).value
                        _GC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                  ctypes.POINTER(ctypes.c_uint))
                        gc_fn = _GC(gc_addr)
                        count = ctypes.c_uint(0)
                        hr = gc_fn(sia_ptr.value, ctypes.byref(count))
                        if hr != 0 or count.value == 0:
                            return paths
                        # IShellItemArray::GetItemAt at vtable index 8
                        gia_addr = ctypes.c_void_p.from_address(sia_vtable + 8 * _vp_size).value
                        _GIA = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                   ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p))
                        gia_fn = _GIA(gia_addr)
                        for i in range(count.value):
                            si_ptr = ctypes.c_void_p(0)
                            hr = gia_fn(sia_ptr.value, i, ctypes.byref(si_ptr))
                            if hr != 0 or not si_ptr.value:
                                continue
                            try:
                                # IShellItem::GetDisplayName at vtable index 5
                                # SIGDN_FILESYSPATH = 0x80058000
                                si_vtable = ctypes.c_void_p.from_address(si_ptr.value).value
                                gdn_addr = ctypes.c_void_p.from_address(si_vtable + 5 * _vp_size).value
                                _GDN = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                          ctypes.c_uint, ctypes.POINTER(ctypes.c_wchar_p))
                                gdn_fn = _GDN(gdn_addr)
                                name_ptr = ctypes.c_wchar_p()
                                hr = gdn_fn(si_ptr.value, 0x80058000, ctypes.byref(name_ptr))
                                if hr == 0 and name_ptr.value:
                                    paths.append(name_ptr.value)
                                    ctypes.windll.ole32.CoTaskMemFree(name_ptr)
                            finally:
                                # Release IShellItem
                                si_vt = ctypes.c_void_p.from_address(si_ptr.value).value
                                rel_addr = ctypes.c_void_p.from_address(si_vt + 2 * _vp_size).value
                                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                                _REL(rel_addr)(si_ptr.value)
                    finally:
                        # Release IShellItemArray
                        sia_vt = ctypes.c_void_p.from_address(sia_ptr.value).value
                        rel_addr = ctypes.c_void_p.from_address(sia_vt + 2 * _vp_size).value
                        _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                        _REL(rel_addr)(sia_ptr.value)
                finally:
                    # Release IFolderView
                    fv_vt = ctypes.c_void_p.from_address(fv_ptr.value).value
                    rel_addr = ctypes.c_void_p.from_address(fv_vt + 2 * _vp_size).value
                    _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                    _REL(rel_addr)(fv_ptr.value)
            finally:
                # Release IShellView
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                rel_addr = ctypes.c_void_p.from_address(vtable_ptr + 2 * _vp_size).value
                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _REL(rel_addr)(sv_ptr)
        except Exception as e:
            debug_print(f"[IEB] _get_ieb_selected_paths failed: {e}")
        return paths

    def _select_ieb_item_by_name(self, filename):
        """通过 IFolderView 直接选中 IExplorerBrowser 当前视图中的项目。"""
        try:
            if not filename:
                return False
            if not isinstance(self.explorer, IExplorerBrowserWidget) or not getattr(self.explorer, '_browser', None):
                return False

            from comtypes import GUID as _GUID

            _vp_size = ctypes.sizeof(ctypes.c_void_p)
            iid_sv = _GUID("{000214E3-0000-0000-C000-000000000046}")
            ppv_sv = self.explorer._browser.GetCurrentView(ctypes.byref(iid_sv))
            sv_ptr = int(ppv_sv) if ppv_sv else 0
            if not sv_ptr:
                debug_print("[IEB Select] IShellView unavailable")
                return False

            def _release_interface(ptr_value):
                if not ptr_value:
                    return
                vt = ctypes.c_void_p.from_address(ptr_value).value
                rel_addr = ctypes.c_void_p.from_address(vt + 2 * _vp_size).value
                _REL = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
                _REL(rel_addr)(ptr_value)

            def _get_display_name(si_ptr_value, sigdn):
                si_vtable = ctypes.c_void_p.from_address(si_ptr_value).value
                gdn_addr = ctypes.c_void_p.from_address(si_vtable + 5 * _vp_size).value
                _GDN = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                          ctypes.c_uint, ctypes.POINTER(ctypes.c_wchar_p))
                gdn_fn = _GDN(gdn_addr)
                name_ptr = ctypes.c_wchar_p()
                hr = gdn_fn(si_ptr_value, sigdn, ctypes.byref(name_ptr))
                if hr == 0 and name_ptr.value:
                    value = name_ptr.value
                    ctypes.windll.ole32.CoTaskMemFree(name_ptr)
                    return value
                return None

            try:
                iid_fv = _GUID("{CDE725B0-CCC9-4519-917E-325D72FAB4CE}")
                vtable_ptr = ctypes.c_void_p.from_address(sv_ptr).value
                qi_addr = ctypes.c_void_p.from_address(vtable_ptr).value
                _QI = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                         ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
                qi_fn = _QI(qi_addr)
                fv_ptr = ctypes.c_void_p(0)
                hr_qi = qi_fn(sv_ptr, ctypes.byref(iid_fv), ctypes.byref(fv_ptr))
                if hr_qi != 0 or not fv_ptr.value:
                    debug_print(f"[IEB Select] QI IFolderView failed: hr=0x{hr_qi & 0xFFFFFFFF:08X}")
                    return False

                try:
                    fv_vtable = ctypes.c_void_p.from_address(fv_ptr.value).value

                    iid_sia = _GUID("{B63EA76D-1F85-456F-A19C-48159EFA858B}")
                    items_addr = ctypes.c_void_p.from_address(fv_vtable + 8 * _vp_size).value
                    _ITEMS = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                ctypes.c_uint, ctypes.POINTER(_GUID),
                                                ctypes.POINTER(ctypes.c_void_p))
                    items_fn = _ITEMS(items_addr)
                    sia_ptr = ctypes.c_void_p(0)
                    # SVGIO_ALLVIEW(0x2) | SVGIO_FLAG_VIEWORDER(0x80000000)：
                    # 必须带 VIEWORDER，使枚举顺序与 IFolderView::SelectItem(iItem) 的
                    # 视图索引一致；否则匹配到的索引会指向显示顺序中的另一个文件，
                    # 表现为“定位到了但选中了错误的文件”。
                    SVGIO_ALLVIEW_VIEWORDER = 0x2 | 0x80000000
                    hr_items = items_fn(fv_ptr.value, SVGIO_ALLVIEW_VIEWORDER, ctypes.byref(iid_sia), ctypes.byref(sia_ptr))
                    if hr_items != 0 or not sia_ptr.value:
                        debug_print(f"[IEB Select] Items(SVGIO_ALLVIEW) failed: hr=0x{hr_items & 0xFFFFFFFF:08X}")
                        return False

                    try:
                        sia_vtable = ctypes.c_void_p.from_address(sia_ptr.value).value
                        gc_addr = ctypes.c_void_p.from_address(sia_vtable + 7 * _vp_size).value
                        _GC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                  ctypes.POINTER(ctypes.c_uint))
                        gc_fn = _GC(gc_addr)
                        count = ctypes.c_uint(0)
                        hr_count = gc_fn(sia_ptr.value, ctypes.byref(count))
                        if hr_count != 0 or count.value == 0:
                            debug_print(f"[IEB Select] View items unavailable: hr=0x{hr_count & 0xFFFFFFFF:08X}, count={count.value}")
                            return False

                        gia_addr = ctypes.c_void_p.from_address(sia_vtable + 8 * _vp_size).value
                        _GIA = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                   ctypes.c_uint, ctypes.POINTER(ctypes.c_void_p))
                        gia_fn = _GIA(gia_addr)

                        target_index = -1
                        filename_cf = filename.casefold()
                        for i in range(count.value):
                            si_ptr = ctypes.c_void_p(0)
                            hr_item = gia_fn(sia_ptr.value, i, ctypes.byref(si_ptr))
                            if hr_item != 0 or not si_ptr.value:
                                continue
                            try:
                                item_name = _get_display_name(si_ptr.value, 0)
                                if not item_name:
                                    item_path = _get_display_name(si_ptr.value, 0x80058000)
                                    item_name = os.path.basename(item_path) if item_path else None
                                if item_name and item_name.casefold() == filename_cf:
                                    target_index = i
                                    break
                            finally:
                                _release_interface(si_ptr.value)

                        if target_index < 0:
                            debug_print(f"[IEB Select] Target not found in current view: {filename}")
                            return False

                        select_addr = ctypes.c_void_p.from_address(fv_vtable + 15 * _vp_size).value
                        _SELECT = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p,
                                                     ctypes.c_int, ctypes.c_uint)
                        select_fn = _SELECT(select_addr)
                        svsi_flags = 0x1 | 0x4 | 0x8 | 0x10
                        hr_select = select_fn(fv_ptr.value, target_index, svsi_flags)
                        if hr_select != 0:
                            debug_print(f"[IEB Select] SelectItem failed: hr=0x{hr_select & 0xFFFFFFFF:08X}, index={target_index}")
                            return False

                        self.activateWindow()
                        self.raise_()
                        if hasattr(self, 'explorer') and self.explorer:
                            self.explorer.setFocus(Qt.OtherFocusReason)

                        debug_print(f"[IEB Select] Successfully selected item via IFolderView: {filename}")
                        self._suppress_auto_refresh = False
                        self._arm_selection_guard()
                        self.start_path_sync_timer(duration_ms=8000)
                        return True
                    finally:
                        _release_interface(sia_ptr.value)
                finally:
                    _release_interface(fv_ptr.value)
            finally:
                _release_interface(sv_ptr)
        except Exception as e:
            debug_print(f"[IEB Select] Failed to select item '{filename}': {e}")
        return False

    # --- Windows native helpers for listview hit-testing ---
    def _find_syslistview_hwnd(self):
        if not HAS_PYWIN:
            return None
        try:
            parent = int(self.explorer.winId())
        except Exception:
            return None

        def find_in_tree(hwnd):
            try:
                cls = win32gui.GetClassName(hwnd)
            except Exception:
                return None
            if cls == 'SysListView32':
                return hwnd
            result = None
            try:
                def cb(h, lparam):
                    nonlocal result
                    if result:
                        return False
                    found = find_in_tree(h)
                    if found:
                        result = found
                        return False
                    return True
                win32gui.EnumChildWindows(hwnd, cb, None)
            except Exception:
                return None
            return result

        return find_in_tree(parent)

    def _install_listview_dblclick_hook(self):
        """Install a single process-wide WH_MOUSE_LL hook (singleton) for
        double-click detection across all IExplorerBrowser tabs.

        Only one hook exists; it dispatches to the current active tab.
        """
        if not HAS_PYWIN:
            return
        if not isinstance(getattr(self, 'explorer', None), IExplorerBrowserWidget):
            return
        if not getattr(self.explorer, '_init_ok', False):
            return

        # Singleton: if a global hook already exists, skip
        if getattr(FileExplorerTab, '_global_mouse_hook_handle', None):
            return

        _self_ref = self  # only used to locate main_window

        class _POINT(ctypes.Structure):
            _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]

        class _MSLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [
                ('pt',          _POINT),
                ('mouseData',   ctypes.c_ulong),
                ('flags',       ctypes.c_ulong),
                ('time',        ctypes.c_ulong),
                ('dwExtraInfo', ctypes.c_size_t),
            ]

        _HOOKPROC = ctypes.WINFUNCTYPE(
            ctypes.c_ssize_t,
            ctypes.c_int,
            ctypes.wintypes.WPARAM,
            ctypes.wintypes.LPARAM,
        )

        WM_LBUTTONDOWN = 0x0201
        WM_MOUSEMOVE   = 0x0200
        WM_RBUTTONDOWN = 0x0204
        WM_RBUTTONUP   = 0x0205
        _GESTURE_MARKER    = 0x47455354  # 'GEST'：标记本程序合成的右键事件，避免被本钩子再次拦截
        _GESTURE_THRESHOLD = 30          # 触发鼠标手势的最小位移（像素）
        _state = {'t': 0, 'x': -9999, 'y': -9999, 'in_exp': False}
        # 鼠标手势（类似 Mouse Gestures）：按住右键画线，支持多笔画序列
        #   ←后退 / →前进 / ↓关闭标签 / ↑新建标签
        #   ↑↓刷新 / ↓↑上级目录 / ↓→恢复关闭的标签页
        _gesture = {'active': False, 'sx': 0, 'sy': 0, 'cx': 0, 'cy': 0,
                    'ax': 0, 'ay': 0, 'seq': []}
        _STROKE_THRESHOLD = 36  # 单笔画方向判定的最小位移（像素）
        _ARROW = {'left': '←', 'right': '→', 'up': '↑', 'down': '↓'}
        # 手势序列 → (动作键, 动作名)
        _GESTURE_TABLE = {
            ('left',):           ('back',    tr('后退')),
            ('right',):          ('forward', tr('前进')),
            ('down',):           ('close',   tr('关闭标签页')),
            ('up',):             ('new',     tr('新建标签页')),
            ('up', 'down'):      ('refresh', tr('刷新')),
            ('down', 'up'):      ('up_dir',  tr('上级目录')),
            ('down', 'right'):   ('reopen',  tr('恢复关闭的标签页')),
        }

        def _seq_arrows(seq):
            return ''.join(_ARROW.get(d, '') for d in seq)

        def _lookup_gesture(seq):
            """返回 (动作键, 动作名)；未匹配返回 (None, '')。"""
            return _GESTURE_TABLE.get(tuple(seq), (None, ''))

        def _get_current_tab():
            """获取当前活动的浏览面板（含分屏面板，由最近一次鼠标按下位置决定）"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                if mw and hasattr(mw, 'get_active_pane'):
                    tab = mw.get_active_pane()
                    if tab and isinstance(getattr(tab, 'explorer', None), IExplorerBrowserWidget):
                        return tab
            except Exception:
                pass
            return None

        def _gestures_enabled():
            """鼠标手势是否启用（config.json: enable_mouse_gestures，默认开启）"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                if mw and hasattr(mw, 'config'):
                    return bool(mw.config.get('enable_mouse_gestures', True))
            except Exception:
                pass
            return True

        def _cursor_over_current_explorer(px, py):
            """光标是否位于某个浏览面板（当前标签或分屏面板）的资源管理器区域内，且本程序窗口为前台。

            仅在满足条件时才接管右键，避免影响其他程序或非文件区域的右键菜单。
            """
            try:
                mw = getattr(_self_ref, 'main_window', None)
                if mw is None:
                    return False
                # 仅当本程序为前台窗口时接管右键
                try:
                    fg = ctypes.windll.user32.GetForegroundWindow()
                    if int(fg) != int(mw.winId()):
                        return False
                except Exception:
                    pass
                return mw.pane_at_global_pos(px, py) is not None
            except Exception:
                return False

        def _point_in_active_explorer(px, py):
            """按下当下坐标是否位于活动标签的 Explorer 视图区。"""
            try:
                from PyQt5.QtCore import QPoint
                tab = _get_current_tab()
                if not tab:
                    return False
                pt = tab.explorer.mapFromGlobal(QPoint(px, py))
                return bool(tab.explorer.rect().contains(pt))
            except Exception:
                return False

        def _run_gesture_action(action):
            """在主线程执行手势对应的操作"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                tab = _get_current_tab()
                if tab is not None:
                    mw = getattr(tab, 'main_window', None) or mw
                if not mw:
                    return
                if action == 'back':
                    mw.go_back_current_tab()
                elif action == 'forward':
                    mw.go_forward_current_tab()
                elif action == 'close':
                    mw.close_current_tab()
                elif action == 'new':
                    # 新建标签作用于手势所在的一侧（左/右分屏组）
                    mw.add_new_tab(target_tabwidget=mw.get_active_group_tabwidget())
                elif action == 'refresh':
                    mw.refresh_current_tab()
                elif action == 'up_dir':
                    mw.go_up_current_tab()
                elif action == 'reopen':
                    mw.reopen_closed_tab()
            except Exception as _e:
                debug_print(f"[Gesture] action error: {_e}")

        def _synth_right_click():
            """合成一次真实右键点击以弹出系统右键菜单（带标记，避免被本钩子再次拦截）"""
            try:
                MOUSEEVENTF_RIGHTDOWN = 0x0008
                MOUSEEVENTF_RIGHTUP   = 0x0010
                ctypes.windll.user32.mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, _GESTURE_MARKER)
                ctypes.windll.user32.mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, _GESTURE_MARKER)
            except Exception as _e:
                debug_print(f"[Gesture] synth right click error: {_e}")

        def _get_overlay():
            """获取/懒创建 MainWindow 上的手势覆盖层单例。"""
            try:
                mw = getattr(_self_ref, 'main_window', None)
                tab = _get_current_tab()
                if tab is not None:
                    mw = getattr(tab, 'main_window', None) or mw
                if mw is None:
                    return None
                ov = getattr(mw, '_gesture_overlay', None)
                if ov is None:
                    ov = GestureOverlay()
                    mw._gesture_overlay = ov
                return ov
            except Exception:
                return None

        def _hook_proc(nCode, wParam, lParam):
            # ── 鼠标手势处理（右键画线）─────────────────────────────────
            if nCode >= 0:
                try:
                    g_info = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                    if int(g_info.dwExtraInfo) == _GESTURE_MARKER:
                        pass  # 本程序合成的事件：直接放行
                    elif wParam == WM_RBUTTONDOWN and _gestures_enabled():
                        gx, gy = int(g_info.pt.x), int(g_info.pt.y)
                        # 手势起点处的面板即为本次手势目标（同时供后续键盘定位）
                        mw_g = getattr(_self_ref, 'main_window', None)
                        if mw_g is not None:
                            try:
                                mw_g.set_active_pane_from_global_pos(gx, gy)
                            except Exception:
                                pass
                        if _cursor_over_current_explorer(gx, gy):
                            _gesture['active'] = True
                            _gesture['sx'], _gesture['sy'] = gx, gy
                            _gesture['cx'], _gesture['cy'] = gx, gy
                            _gesture['ax'], _gesture['ay'] = gx, gy
                            _gesture['seq'] = []
                            ov = _get_overlay()
                            if ov is not None:
                                ov.start()
                                ov.add_point(gx, gy)
                            return 1  # 吞掉按下，松开时再决定执行手势或弹出菜单
                    elif wParam == WM_MOUSEMOVE and _gesture['active']:
                        nx, ny = int(g_info.pt.x), int(g_info.pt.y)
                        _gesture['cx'], _gesture['cy'] = nx, ny
                        # 多笔画方向识别：相对当前笔画锚点的位移超过阈值则提交一个方向
                        adx = nx - _gesture['ax']
                        ady = ny - _gesture['ay']
                        if max(abs(adx), abs(ady)) >= _STROKE_THRESHOLD:
                            if abs(adx) >= abs(ady):
                                d = 'right' if adx > 0 else 'left'
                            else:
                                d = 'down' if ady > 0 else 'up'
                            seq = _gesture['seq']
                            if not seq or seq[-1] != d:
                                seq.append(d)
                            _gesture['ax'], _gesture['ay'] = nx, ny
                        ov = _get_overlay()
                        if ov is not None:
                            ov.add_point(nx, ny)
                            _action, _label = _lookup_gesture(_gesture['seq'])
                            ov.set_hint(_seq_arrows(_gesture['seq']) if _label else '', _label)
                    elif wParam == WM_RBUTTONUP and _gesture['active']:
                        _gesture['active'] = False
                        ov = _get_overlay()
                        if ov is not None:
                            ov.finish()
                        action, label = _lookup_gesture(_gesture['seq'])
                        if action:
                            debug_print(f"[Gesture] {_seq_arrows(_gesture['seq'])} → {action}")
                            QTimer.singleShot(0, lambda a=action: _run_gesture_action(a))
                        else:
                            # 未识别为有效手势 → 合成真实右键以弹出菜单
                            QTimer.singleShot(0, _synth_right_click)
                        _gesture['seq'] = []
                        return 1  # 吞掉这次松开（已执行手势或即将合成右键）
                except Exception:
                    pass

            if nCode >= 0 and wParam == WM_LBUTTONDOWN:
                try:
                    info = ctypes.cast(lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)).contents
                    px = int(info.pt.x)
                    py = int(info.pt.y)
                    t  = int(info.time)

                    # 左键按下位置所在面板即为活动面板（供键盘快捷键/双击定位）
                    mw_l = getattr(_self_ref, 'main_window', None)
                    if mw_l is not None:
                        try:
                            mw_l.set_active_pane_from_global_pos(px, py)
                        except Exception:
                            pass

                    # 点击 explorer 区域内时，通知路径栏退出编辑模式
                    try:
                        from PyQt5.QtCore import QPoint
                        tab = _get_current_tab()
                        if tab:
                            pt_chk = tab.explorer.mapFromGlobal(QPoint(px, py))
                            if tab.explorer.rect().contains(pt_chk):
                                if hasattr(tab, 'path_bar') and getattr(tab.path_bar, '_in_edit', False):
                                    tab.path_bar.exit_edit_mode()
                                elif hasattr(tab, 'path_bar') and getattr(tab.path_bar, 'edit_mode', False):
                                    tab.path_bar.exit_edit_mode()
                    except Exception:
                        pass

                    dct = ctypes.windll.user32.GetDoubleClickTime()
                    dt  = (t - _state['t']) & 0xFFFFFFFF
                    in_exp = _point_in_active_explorer(px, py)

                    is_dblclick = (
                        _state['t'] != 0 and
                        dt <= dct and
                        abs(px - _state['x']) <= 4 and
                        abs(py - _state['y']) <= 4 and
                        in_exp and bool(_state.get('in_exp', False))
                    )

                    if is_dblclick:
                        _state['t'] = 0

                        # 记录双击瞬间的路径：若 150ms 后路径已变化，说明这次双击其实是
                        # 进入了盘符/文件夹（Explorer 内部导航），绝不能再 go_up，否则会出现
                        # “双击 D 盘进入后又立即退回此电脑”的抖动。
                        _tab_at_click = _get_current_tab()
                        _path_at_click = getattr(_tab_at_click, 'current_path', None) if _tab_at_click else None
                        if _tab_at_click and _tab_at_click._is_mycomputer_shell_target(_path_at_click):
                            _tab_at_click._mark_shell_item_activation(ttl_sec=1.6)
                            debug_print("[ShellActivation] tentative arm from low-level dblclick on MyComputer")

                        def _check(_px=px, _py=py, _path_at_click=_path_at_click):
                            try:
                                tab = _get_current_tab()
                                if not tab:
                                    return

                                # 路径已变化 → 双击进入了文件夹/盘符，取消 go_up
                                if _path_at_click is not None and getattr(tab, 'current_path', None) != _path_at_click:
                                    return

                                # Bounds check
                                from PyQt5.QtCore import QPoint
                                pt_local = tab.explorer.mapFromGlobal(QPoint(_px, _py))
                                if not tab.explorer.rect().contains(pt_local):
                                    return

                                # Navigation guard
                                last_nav = getattr(tab, '_last_nav_complete_time', 0)
                                if time.monotonic() - last_nav < 0.8:
                                    return

                                # LVM_HITTEST
                                lv = tab._find_syslistview_hwnd()
                                if lv:
                                    try:
                                        cpt = win32gui.ScreenToClient(lv, (_px, _py))

                                        class _PT(ctypes.Structure):
                                            _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]
                                        class _LVHI(ctypes.Structure):
                                            _fields_ = [('pt', _PT), ('flags', ctypes.c_uint),
                                                         ('iItem', ctypes.c_int), ('iSubItem', ctypes.c_int)]

                                        hi = _LVHI()
                                        hi.pt.x, hi.pt.y = int(cpt[0]), int(cpt[1])
                                        res = ctypes.windll.user32.SendMessageW(
                                            lv, 0x1012, 0, ctypes.byref(hi))
                                        if int(res) != -1:
                                            if tab._is_mycomputer_shell_target(getattr(tab, 'current_path', '')):
                                                tab._mark_shell_item_activation()
                                            tab._suppress_auto_refresh = False
                                            tab._resume_path_sync_after_navigation()
                                            return
                                    except Exception:
                                        pass

                                cnt = tab._get_selected_count_safe()
                                if cnt is not None and int(cnt) > 0:
                                    return

                                debug_print("[DoubleClick/IEB] Blank area → go_up")
                                if tab._is_mycomputer_shell_target(getattr(tab, 'current_path', '')):
                                    tab._clear_shell_item_activation("blank-dblclick-go-up")
                                tab.go_up(force=True)
                            except Exception as _e:
                                debug_print(f"[DoubleClick/IEB] check: {_e}")

                        QTimer.singleShot(150, _check)
                    else:
                        _state['t'] = t
                        _state['x'] = px
                        _state['y'] = py
                        _state['in_exp'] = bool(in_exp)
                except Exception:
                    pass

            try:
                hh = getattr(FileExplorerTab, '_global_mouse_hook_handle', None) or 0
                return ctypes.windll.user32.CallNextHookEx(hh, nCode, wParam, lParam)
            except Exception:
                return 0

        cb = _HOOKPROC(_hook_proc)
        handle = ctypes.windll.user32.SetWindowsHookExW(14, cb, None, 0)
        if handle:
            FileExplorerTab._global_mouse_hook_cb     = cb       # prevent GC
            FileExplorerTab._global_mouse_hook_handle = handle
            self._lv_hook_handle = handle  # backward compat for logs
            debug_print("[IEB] WH_MOUSE_LL hook installed (singleton)")
        else:
            err = ctypes.windll.kernel32.GetLastError()
            debug_print(f"[IEB] SetWindowsHookExW(WH_MOUSE_LL) failed err={err}")

    def _uninstall_listview_dblclick_hook(self):
        """Remove the singleton WH_MOUSE_LL hook only when the app is closing."""
        self._lv_hook_handle = None
        # Only actually unhook when the MainWindow is closing (no more tabs)
        handle = getattr(FileExplorerTab, '_global_mouse_hook_handle', None)
        if not handle:
            return
        # Check if any other IEB tab still exists
        mw = getattr(self, 'main_window', None)
        if mw and hasattr(mw, 'content_stack'):
            for i in range(mw.content_stack.count()):
                w = mw.content_stack.widget(i)
                if w is not self and isinstance(getattr(w, 'explorer', None), IExplorerBrowserWidget):
                    return  # other IEB tabs still alive, keep hook
        try:
            ctypes.windll.user32.UnhookWindowsHookEx(handle)
            debug_print("[IEB] WH_MOUSE_LL hook removed (singleton)")
        except Exception as e:
            debug_print(f"[IEB] UnhookWindowsHookEx error: {e}")
        finally:
            FileExplorerTab._global_mouse_hook_handle = None
            FileExplorerTab._global_mouse_hook_cb = None

    def _native_listview_hit_test(self, screen_x, screen_y):
        if not HAS_PYWIN:
            return False
        try:
            lv = self._find_syslistview_hwnd()
            if not lv:
                return False
            # convert screen -> client
            pt = (int(screen_x), int(screen_y))
            try:
                cx, cy = win32gui.ScreenToClient(lv, pt)
            except Exception:
                return False

            class POINT(ctypes.Structure):
                _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]

            class LVHITTESTINFO(ctypes.Structure):
                _fields_ = [('pt', POINT), ('flags', ctypes.c_uint), ('iItem', ctypes.c_int), ('iSubItem', ctypes.c_int)]

            info = LVHITTESTINFO()
            info.pt.x = int(cx)
            info.pt.y = int(cy)
            LVM_FIRST = 0x1000
            LVM_HITTEST = LVM_FIRST + 18
            res = ctypes.windll.user32.SendMessageW(lv, LVM_HITTEST, 0, ctypes.byref(info))
            try:
                if int(res) == -1:
                    return False
                return True
            except Exception:
                return False
        except Exception:
            return False

    def eventFilter(self, obj, event):
        # 通过事件过滤器捕获 Explorer 的鼠标按下与双击事件
        from PyQt5.QtCore import QEvent, QTimer, Qt
        
        # 注意：快捷键处理现在由MainWindow的轮询定时器处理，不在这里处理
        
        try:
            if obj is self.explorer:
                if event.type() == QEvent.MouseButtonPress:
                    # 记录按下时的选中项数
                    try:
                        cnt = self._get_selected_count_safe()
                        self._selected_before_click = int(cnt) if cnt is not None else None
                    except Exception:
                        self._selected_before_click = None
                    self._schedule_status_update(track_selection=True)
                elif event.type() == QEvent.ContextMenu:
                    self._schedule_status_update(track_selection=True)
                    global_pos = event.globalPos() if hasattr(event, 'globalPos') else QCursor.pos()
                    if self.show_selected_item_context_menu(global_pos):
                        return True
                elif event.type() == QEvent.MouseButtonRelease:
                    self._schedule_status_update(track_selection=True)
                elif event.type() == QEvent.MouseButtonDblClick:
                    self._schedule_status_update(track_selection=True)
                    # 取消所有之前待处理的双击检查
                    for _t in getattr(self, '_pending_double_click_timers', []):
                        try: _t.stop()
                        except Exception: pass
                    self._pending_double_click_timers = []

                    # 立即在双击位置做 native hit-test —— 这是最可靠的判断
                    # True  = 点中了列表项（文件夹/文件），不触发 go_up
                    # False = 点在空白区，触发 go_up
                    double_click_pos = QCursor.pos()
                    path_before = getattr(self, 'current_path', None)
                    if self._is_mycomputer_shell_target(path_before):
                        self._mark_shell_item_activation(ttl_sec=1.6)
                        debug_print("[ShellActivation] tentative arm from Qt dblclick on MyComputer")

                    if HAS_PYWIN:
                        hit = self._native_listview_hit_test(double_click_pos.x(), double_click_pos.y())
                        debug_print(f"[DoubleClick] hit-test={hit}, path_before='{path_before}'")
                        if hit:
                            if self._is_mycomputer_shell_target(path_before):
                                self._mark_shell_item_activation()
                            # 点中了项目，让 Explorer 自己处理（打开文件夹/文件）
                            # 无论是否能立即检测到选中项，都启动路径同步定时器，
                            # 确保进入子目录后地址栏及时更新（Explorer可能在双击瞬间已清除选中状态）
                            sel = self._get_selected_count_safe()
                            if sel and int(sel) > 0:
                                self._navigating_folder = True
                            # 解除导航抑制标志，确保本次用户主动双击触发的 Explorer 内部
                            # 导航能被路径同步定时器检测到，不被之前 navigate_to 的 3s 抑制窗口阻断。
                            self._suppress_auto_refresh = False
                            # 50ms 后启动 polling 兜底（NavigateComplete2 直连信号优先触发）
                            QTimer.singleShot(50, self._resume_path_sync_after_navigation)
                        else:
                            # 空白区域双击 —— 用极短延迟（50ms）执行 go_up，
                            # 50ms 仅为让 Explorer 完成 dblclick 内部处理，不会有可见延迟
                            def _do_go_up_blank():
                                try:
                                    # 二次安全确认：路径未变且无选中，再 go_up
                                    if getattr(self, 'current_path', None) != path_before:
                                        return
                                    cnt = self._get_selected_count_safe()
                                    if cnt and int(cnt) > 0:
                                        return
                                    debug_print(f"[DoubleClick] Blank area confirmed, executing go_up")
                                    if self._is_mycomputer_shell_target(getattr(self, 'current_path', '')):
                                        self._clear_shell_item_activation("qt-blank-dblclick-go-up")
                                    self.go_up(force=True)
                                except Exception as e:
                                    debug_print(f"[DoubleClick] go_up exception: {e}")
                            t = QTimer(self)
                            t.setSingleShot(True)
                            t.timeout.connect(_do_go_up_blank)
                            t.start(50)
                            self._pending_double_click_timers.append(t)
                    else:
                        # 无 pywin32：退回到 150ms 路径/选中检查（比原 400ms+700ms 仍快很多）
                        selected_before = getattr(self, '_selected_before_click', None)
                        def _fallback_check():
                            try:
                                if getattr(self, '_navigating_folder', False):
                                    self._navigating_folder = False
                                    return
                                cur_path = getattr(self, 'current_path', None)
                                if path_before is not None and cur_path != path_before:
                                    return
                                cnt = self._get_selected_count_safe()
                                if cnt is not None and int(cnt) > 0:
                                    return
                                if selected_before is not None and int(selected_before) > 0:
                                    return
                                debug_print(f"[DoubleClick] Fallback: blank area confirmed, executing go_up")
                                self.go_up(force=True)
                            except Exception as e:
                                debug_print(f"[DoubleClick] Fallback exception: {e}")
                            finally:
                                self._selected_before_click = None
                        t = QTimer(self)
                        t.setSingleShot(True)
                        t.timeout.connect(_fallback_check)
                        t.start(150)
                        self._pending_double_click_timers.append(t)
                elif event.type() in (QEvent.KeyRelease, QEvent.FocusIn, QEvent.Wheel):
                    self._schedule_status_update(track_selection=True)
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def blank_double_click(self, event):
        self.go_up(force=True)

    def _status_bar_mouse_press(self, event):
        """在状态栏按下左键时，开始拖动整个软件窗口。"""
        if event.button() == Qt.LeftButton:
            win = self.window()
            # 优先使用系统原生窗口移动（Qt 5.15+，最大化等情况处理更可靠）
            handle = win.windowHandle()
            if handle is not None and hasattr(handle, 'startSystemMove'):
                try:
                    if handle.startSystemMove():
                        self._status_drag_pos = None
                        event.accept()
                        return
                except Exception:
                    pass
            self._status_drag_pos = event.globalPos() - win.frameGeometry().topLeft()
            event.accept()
            return
        QLabel.mousePressEvent(self.status_bar, event)

    def _status_bar_mouse_move(self, event):
        """拖动状态栏时移动整个软件窗口。"""
        if event.buttons() == Qt.LeftButton and self._status_drag_pos is not None:
            win = self.window()
            if not win.isMaximized():
                win.move(event.globalPos() - self._status_drag_pos)
            event.accept()
            return
        QLabel.mouseMoveEvent(self.status_bar, event)

    def _status_bar_mouse_release(self, event):
        """结束状态栏拖动。"""
        self._status_drag_pos = None
        QLabel.mouseReleaseEvent(self.status_bar, event)

    def _status_bar_mouse_double_click(self, event):
        """双击底部状态栏：手动触发一次“缩放式”重排刷新，兜底修复路径栏偶发卡死。"""
        if event.button() != Qt.LeftButton:
            QLabel.mouseDoubleClickEvent(self.status_bar, event)
            return
        try:
            mw = getattr(self, 'main_window', None)
            if mw and hasattr(mw, '_manual_statusbar_reflow_refresh'):
                mw._manual_statusbar_reflow_refresh()
            else:
                self._manual_pathbar_rebuild()
            event.accept()
        except Exception as e:
            debug_print(f"[StatusBar] manual reflow refresh failed: {e}")
            QLabel.mouseDoubleClickEvent(self.status_bar, event)

    def _manual_pathbar_rebuild(self):
        """仅重建当前标签路径栏（MainWindow 不可用时的兜底）。"""
        pb = getattr(self, 'path_bar', None)
        current_path = getattr(self, 'current_path', '')
        if pb and current_path and not getattr(pb, '_in_edit', False):
            pb.set_path(current_path)
            try:
                pb.force_refresh()
            except Exception:
                pass

    # 移除 on_document_complete 和 eventFilter 相关内容

    def go_up(self, force=False):
        # 返回上一级目录，盘符根目录时导航到"此电脑"
        # 如果 force=True，则绕过鼠标位置检查（用于按钮或程序化调用）
        if not self.current_path:
            return
        
        debug_print(f"[go_up] Called with force={force}, current_path='{self.current_path}'")
        
        # 快速检查：如果是force模式，直接跳过widget检查
        if not force:
            # 仅在明确来自空白区域或路径栏的触发时执行，避免误由文件双击触发
            try:
                pos = QCursor.pos()
                w = QApplication.widgetAt(pos.x(), pos.y())
                # 允许的触发源：底部空白标签或路径栏
                if w is not self.blank and w is not getattr(self, 'path_bar', None):
                    debug_print(f"[go_up] Rejected: not from valid source")
                    return
            except Exception:
                debug_print(f"[go_up] Rejected: exception in widget check")
                return
        
        path = self.current_path
        # 判断是否为盘符根目录，导航到"此电脑"
        if path.endswith(":\\") or path.endswith(":/"):
            debug_print(f"[go_up] Root directory, navigate to MyComputer")
            # 避免残留许可窗口把“回到此电脑”又立即放行回盘符。
            self._clear_shell_item_activation("go_up-root-to-mycomputer")
            self.navigate_to('shell:MyComputerFolder', is_shell=True, skip_async_check=True)
            return
        
        parent_path = os.path.dirname(path)
        if parent_path and os.path.exists(parent_path):
            debug_print(f"[go_up] Navigate to parent: {parent_path}")
            # 返回上一级时跳过异步检查，直接导航，提升响应速度
            self.navigate_to(parent_path, skip_async_check=True)
        else:
            debug_print(f"[go_up] Invalid parent path")

    def _normalize_local_path(self, path):
        if not isinstance(path, str) or not path:
            return path
        if path.startswith('shell:') or '::' in path:
            return path
        # 自愈：修复历史会话/缓存中被破坏的 UNC 路径。旧版本把 \\server\share
        # 误存成单反斜杠 \server\share（丢了一个 \），导致 SHParseDisplayName 失败。
        # 统一分隔符后，若以单个分隔符开头（非双）且至少有两段（\主机\共享…），
        # 判定为被破坏的 UNC 并还原 \\ 前缀。本应用只存绝对路径，不会有盘符相对路径。
        unified = path.replace('/', '\\')
        if unified.startswith('\\') and not unified.startswith('\\\\'):
            rest = unified.lstrip('\\')
            if rest and rest[1:2] != ':':  # 排除形如 \C:\ 的异常
                segs = [s for s in rest.split('\\') if s]
                if len(segs) >= 2:
                    unified = '\\\\' + rest
        try:
            return os.path.normpath(unified)
        except Exception:
            return unified

    def __init__(self, parent=None, path="", is_shell=False, select_file=None, defer_nav=False):
        super().__init__(parent)
        self.main_window = parent
        initial_path = path if path else QDir.homePath()
        self.current_path = self._normalize_local_path(initial_path)
        self.select_file = select_file  # 要选中的文件名
        # 延迟首次导航：会话恢复时后台标签用此模式，避免启动瞬间 N 个 IExplorerBrowser
        # 同时创建 COM/导航/overlay 预加载/scandir 造成的 CPU 洪峰。首次可见（showEvent）时才导航。
        self._deferred_nav = None  # (path, is_shell) 待首次可见时执行；None 表示无待处理导航
        self.notepad_plus_plus_path = detect_notepad_plus_plus()
        # 浏览历史记录
        self.history = []
        self.history_index = -1
        # 标志：是否正在程序化导航（用于防止sync时重复添加历史）
        self._navigating_programmatically = False
        # “此电脑”内项目激活许可窗口：仅在窗口内允许 shell -> 盘符切换。
        self._shell_item_activation_until = 0.0
        # 用于跟踪待处理的双击检查定时器
        self._pending_double_click_timers = []
        # 双击事件唯一ID，用于区分不同的双击操作
        self._double_click_id = 0
        # Win32 WH_MOUSE thread hook for IExplorerBrowser double-click detection
        self._lv_hook_wndproc = None   # ctypes callback (must stay alive to prevent GC)
        self._lv_hook_handle  = None   # HHOOK handle returned by SetWindowsHookExW
        self._is_cleaning_up = False
        
        # 文件系统监控（监控当前路径的变化）
        self.file_watcher = QFileSystemWatcher(self)
        self.file_watcher.directoryChanged.connect(self.on_directory_changed)
        self.file_watcher.fileChanged.connect(self.on_file_changed)
        # 延迟刷新定时器（避免频繁刷新）
        self.refresh_timer = QTimer(self)
        self.refresh_timer.setSingleShot(True)
        self.refresh_timer.timeout.connect(self.delayed_refresh)
        self.refresh_delay_ms = 500  # 500ms延迟
        self._refresh_min_interval_ms = 3000  # 连续刷新最小间隔，避免COM刷新风暴卡界面
        self._last_refresh_ts_ms = 0
        # 防抖机制：记录最近处理的路径和时间，避免事件风暴
        self._last_watcher_event = {}  # {path: timestamp}
        self._watcher_debounce_ms = 3000  # 3s 内的重复事件会被忽略，进一步抑制风暴
        self._refresh_burst_times = []  # 刷新时间戳列表，用于检测风暴
        self._refresh_burst_suppressed_until = 0  # 风暴期间暂停刷新的截止时间戳(ms)
        self._last_file_event = {}  # {file_path: timestamp_ms}
        self._file_event_debounce_ms = 400
        self._watched_files = set()
        self._last_dir_snapshot = None
        # 后台目录快照（兜底轮询）：懒创建的信号对象 + in-flight 去重标志
        self._snapshot_signals = None
        self._snapshot_inflight = False
        self._refresh_active = False
        self._refresh_pending = False
        self._refresh_pending_reason = None
        self._manual_refresh_frozen = False
        # Ctrl+G 选中保护：选中文件后短时间内抑制自动刷新，避免后台目录变动触发的
        # Refresh()（完整重导航）立即清掉刚设置的选中项，导致“跳转选中总是失败”。
        # 注意：仅抑制自动刷新，不影响路径同步（回车进入子目录仍能更新地址栏）。
        self._selection_guard_until = 0.0
        self._status_tracking_deadline_ms = 0
        self.status_update_timer = QTimer(self)
        self.status_update_timer.setSingleShot(True)
        self.status_update_timer.timeout.connect(self.update_explorer_status)
        self.status_tracking_timer = QTimer(self)
        self.status_tracking_timer.setInterval(STATUS_TRACKING_INTERVAL_MS)
        self.status_tracking_timer.timeout.connect(self._poll_status_during_interaction)

        # 目录轮询刷新兜底：处理部分编辑器改文件但目录 watcher 不触发的情况
        self.dir_mtime = None
        self.dir_poll_timer = QTimer(self)
        self.dir_poll_timer.setInterval(8000)  # 8s 低频轮询，后台标签降低I/O开销
        self.dir_poll_timer.timeout.connect(self._poll_directory_changes)
        
        self.setup_ui()
        
        # 安装事件过滤器来处理快捷键（让Ctrl键能穿透到主窗口）
        self.installEventFilter(self)
        if hasattr(self, 'explorer'):
            self.explorer.installEventFilter(self)
        
        if defer_nav:
            # 延迟首次导航：仅在路径栏显示占位路径，暂存导航参数；
            # 不创建 IExplorerBrowser、不导航、不 scandir、不 overlay 预加载，
            # 直到该标签首次可见（showEvent）时才真正导航。
            deferred_is_shell = bool(
                is_shell or str(self.current_path).lower().startswith('shell:')
            )
            self._deferred_nav = (self.current_path, deferred_is_shell)
            if hasattr(self, 'path_bar') and self.path_bar:
                try:
                    self.path_bar.set_path(self.current_path)
                except Exception:
                    pass
            self.update_tab_title()
        else:
            self.navigate_to(self.current_path, is_shell=is_shell)
        # 路径同步定时器已在setup_ui中启动，这里不需要重复启动
        
        # 如果指定了要选中的文件，延迟选中（等待导航完成）
        # 增加延迟时间确保文件夹完全加载
        if self.select_file:
            QTimer.singleShot(1500, lambda: self.select_file_in_explorer(self.select_file))

    def _release_folder_checker(self, wait_ms=150):
        checker = getattr(self, 'folder_checker', None)
        self.folder_checker = None
        if not checker:
            return
        try:
            checker.completed.disconnect()
        except Exception:
            pass
        try:
            checker.stop()
        except Exception:
            pass
        try:
            if checker.isRunning():
                checker.wait(wait_ms)
        except Exception:
            pass
        try:
            _retain_thread_until_finished(checker)
        except Exception:
            pass

    def cleanup(self):
        if self._is_cleaning_up:
            return
        self._is_cleaning_up = True
        # Remove Win32 subclass hook before any other cleanup
        self._uninstall_listview_dblclick_hook()

        if hasattr(self, '_pending_double_click_timers'):
            for timer in self._pending_double_click_timers:
                try:
                    timer.stop()
                    timer.deleteLater()
                except Exception:
                    pass
            self._pending_double_click_timers = []

        for timer_name in (
            'refresh_timer',
            'status_update_timer',
            'status_tracking_timer',
            'dir_poll_timer',
            '_path_sync_timer',
            '_path_sync_stop_timer',
            '_keepalive_sync_timer',
        ):
            timer = getattr(self, timer_name, None)
            if timer:
                try:
                    timer.stop()
                    timer.deleteLater()
                except Exception:
                    pass
                setattr(self, timer_name, None)

        watcher = getattr(self, 'file_watcher', None)
        if watcher:
            try:
                watched_paths = watcher.directories() + watcher.files()
                if watched_paths:
                    watcher.removePaths(watched_paths)
            except Exception:
                pass

        self._release_folder_checker(wait_ms=150)

        # 关闭 LocationURL 看门狗线程池
        ex = getattr(self, '_com_executor', None)
        if ex is not None:
            try:
                ex.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass
            self._com_executor = None

        explorer = getattr(self, 'explorer', None)
        if explorer:
            try:
                explorer.removeEventFilter(self)
            except Exception:
                pass
            try:
                explorer.dynamicCall('Stop()')
            except Exception:
                pass
            try:
                explorer.clear()
            except Exception:
                pass

        try:
            self.removeEventFilter(self)
        except Exception:
            pass

        self._watched_files.clear()
        self._last_watcher_event.clear()
        self._last_file_event.clear()
        self._last_dir_snapshot = None
        # 断开后台快照信号，避免延迟到达的 QThreadPool 结果回调已销毁的标签
        sig = getattr(self, '_snapshot_signals', None)
        if sig is not None:
            try:
                sig.done.disconnect()
            except Exception:
                pass
            try:
                sig.deleteLater()
            except Exception:
                pass
            self._snapshot_signals = None
        self._snapshot_inflight = False

    # 移除重复的setup_ui，保留带路径栏的实现

    def get_selected_filenames(self):
        """获取选中的文件名列表（仅文件名，含后缀）"""
        filenames = []
        try:
            # IExplorerBrowser 模式：通过 COM 接口直接获取选中路径
            if isinstance(self.explorer, IExplorerBrowserWidget):
                paths = self._get_ieb_selected_paths()
                for p in paths:
                    if p:
                        base = os.path.basename(str(p).rstrip('\\/'))
                        if base:
                            filenames.append(base)
                return filenames
            # 旧 QAxWidget 模式：通过Document接口获取SelectedItems
            doc = self.explorer.querySubObject('Document')
            if doc:
                selected = doc.querySubObject('SelectedItems()')
                if selected:
                    count = selected.dynamicCall('Count()')
                    if count and count > 0:
                        for i in range(count):
                            item = selected.querySubObject('Item(int)', i)
                            if item:
                                name = item.dynamicCall('Name()')
                                if name:
                                    filenames.append(str(name))
        except Exception as e:
            debug_print(f"[get_selected_filenames] Error: {e}")
        return filenames

    def showEvent(self, event):
        """标签首次可见时消费延迟导航：真正创建 IExplorerBrowser 并导航。

        会话恢复时后台标签以 defer_nav 模式创建（不导航），切换过去首次可见即在此导航，
        从而把 N 个 Shell 视图的创建/导航分摊到用户实际访问时，消除启动 CPU 洪峰。"""
        super().showEvent(event)
        deferred = getattr(self, '_deferred_nav', None)
        if deferred is not None:
            self._deferred_nav = None
            path, is_shell = deferred
            debug_print(f"[navigate_to] Deferred first navigation on show: '{path}' (is_shell={is_shell})")
            try:
                self.navigate_to(path, is_shell=is_shell)
            except Exception as e:
                debug_print(f"[navigate_to] Deferred navigation failed: {e}")

    def navigate_to(self, path, is_shell=False, add_to_history=True, skip_async_check=False):
        if not is_shell:
            path = self._normalize_local_path(path)
        debug_print(f"[navigate_to] To '{path}' (is_shell={is_shell}, skip_async={skip_async_check})")
        # 导航进行中（慢盘异步解析未完成）：忽略对同一目标的重复点击，避免频繁操作堆积后台线程。
        # 仅拦截相同目标；切换到不同路径仍放行（旧解析结果由导航代号自动作废）。
        if (not is_shell and getattr(self, '_nav_in_progress', False) and
                path == getattr(self, '_nav_in_progress_path', None)):
            debug_print(f"[navigate_to] Ignoring duplicate nav while in progress: {path}")
            return
        # 控制面板及其子目录用原生窗口打开，不嵌入
        if self._is_control_panel_path(path):
            try:
                import subprocess
                launch_detached(['explorer.exe', path])
                show_toast(self, tr("已打开"), tr("控制面板已在新窗口打开"), level="info", duration=2000)
            except Exception as e:
                show_toast(self, tr("错误"), tr("无法打开控制面板: {}").format(e), level="error")
            if hasattr(self, 'path_bar'):
                self.path_bar.set_path(self.current_path)
            return

        # 快速取消所有待处理的双击检查定时器
        if hasattr(self, '_pending_double_click_timers'):
            for timer in self._pending_double_click_timers:
                try:
                    timer.stop()
                except Exception:
                    pass
            self._pending_double_click_timers = []

        # 停止之前的文件夹检查线程（减少等待时间）
        if hasattr(self, 'folder_checker') and self.folder_checker and self.folder_checker.isRunning():
            self._release_folder_checker(wait_ms=50)  # 减少等待时间从100ms到50ms

        # 支持本地路径和shell特殊路径
        if is_shell:
            # shell:OneDrive 解析为真实路径（Shell.Explorer无法正确显示内容）
            if path.lower() == 'shell:onedrive':
                onedrive_path = os.environ.get('OneDrive', '')
                if onedrive_path and os.path.exists(onedrive_path):
                    self.navigate_to(onedrive_path, is_shell=False, add_to_history=add_to_history, skip_async_check=True)
                    return
            self._mark_expected_navigation(path, is_shell=True)
            self._hide_loading_indicator()
            self.explorer.dynamicCall("Navigate(const QString&)", path)
            self.current_path = path
            if hasattr(self, 'path_bar'):
                self.path_bar.set_path(path)
            if self._is_mycomputer_shell_target(path):
                self._force_tile_view_for_this_pc()
            self.update_tab_title()
            # 添加到历史记录
            if add_to_history:
                self._add_to_history(path)
            self.update_explorer_status()
        elif self._is_slow_path(path):
            # 慢盘（网络/UNC/映射盘/OneDrive）：os.path.exists / os.path.isdir 在挂起的
            # 网络路径上会同步阻塞 UI 线程，导致整个程序卡死。提前判定并跳过所有同步
            # 文件系统探测，直接进入导航流程（PIDL 解析已在后台线程异步执行）。
            debug_print(f"[navigate_to] Slow/network path, skipping sync fs checks: {path}")
            self._perform_navigation(path, add_to_history)
        elif os.path.exists(path):
            # 如果skip_async_check=True或禁用异步，直接导航
            # OneDrive/网络路径的os.scandir()可能永久阻塞，直接跳过异步检查
            is_slow = self._is_slow_path(path)
            if is_slow:
                debug_print(f"[navigate_to] OneDrive/network path detected, skipping async check: {path}")
            if skip_async_check or not ASYNC_LOAD_ENABLED or not os.path.isdir(path) or is_slow:
                self._perform_navigation(path, add_to_history)
            else:
                # 异步检查文件夹大小
                self._check_folder_size_async(path, add_to_history)
        elif path.startswith('\\\\') or path.startswith('//'):
            # UNC 网络路径可能因网络延迟导致 os.path.exists 返回 False，直接尝试导航
            debug_print(f"[navigate_to] UNC path, attempting navigation despite exists=False: {path}")
            self._perform_navigation(path, add_to_history)
        else:
            debug_print(f"[navigate_to] Path does not exist: {path}")
            if is_network_path(path):
                # 盘符仍记录为网络驱动器但当前未连接
                self._show_network_error(path, ERROR_CONNECTION_UNAVAIL)

    def rebuild_shell_view(self):
        """切换深浅色后重建 Shell 视图（原生视图只在新建时完整应用配色）。不可见的标签延迟到下次显示。"""
        explorer = getattr(self, 'explorer', None)
        path = getattr(self, 'current_path', '')
        if not isinstance(explorer, IExplorerBrowserWidget) or not explorer._init_ok or not path:
            return False
        if getattr(explorer, '_created_dark', None) == _theme.is_dark():
            return False
        if not explorer.hibernate(path):
            return False
        if self.isVisible():
            # 销毁视图会作废在途的后台解析，需解除导航中锁定才能重新导航到同一路径
            self._nav_in_progress = False
            self._hide_loading_indicator()
            self.navigate_to(path, is_shell=str(path).startswith('shell:'), add_to_history=False)
        return True

    def _is_slow_path(self, path):
        """检测OneDrive/网络/映射网络驱动器路径——这类路径上os.scandir()可能阻塞UI线程"""
        if not path:
            return False
        # 网络UNC路径
        if path.startswith('\\\\') or path.startswith('//'):
            return True
        # OneDrive同步文件夹（路径中含OneDrive关键字）
        path_lower = path.replace('\\', '/').lower()
        if 'onedrive' in path_lower:
            return True
        # 映射网络驱动器（盘符类型 DRIVE_REMOTE=4）——网络抖动时os.scandir()同样永久阻塞
        try:
            import ctypes
            drive = os.path.splitdrive(path)[0]  # e.g. 'D:'
            if drive:
                DRIVE_REMOTE = 4
                if ctypes.windll.kernel32.GetDriveTypeW(drive + '\\') == DRIVE_REMOTE:
                    return True
        except Exception:
            pass
        return False

    def _is_control_panel_path(self, path):
        """判断路径是否为控制面板或其子目录"""
        if not path:
            return False
        s = path.lower()
        # shell:ControlPanelFolder
        if s.startswith('shell:controlpanelfolder'):
            return True
        # 控制面板 CLSID
        if '::{26ee0668-a00a-44d7-9371-beb064c98683}' in s:
            return True
        # 控制面板的子项目通常以 control panel/ 或 control panel\\ 开头
        if s.startswith('control panel') or s.startswith('control panel/') or s.startswith('control panel\\'):
            return True
        # 也可能是 file:///C:/Windows/System32/control.exe 或类似
        if 'control.exe' in s:
            return True
        # 也可能是 shell:::{26ee0668-a00a-44d7-9371-beb064c98683} 或其子路径
        if s.startswith('shell:::{26ee0668-a00a-44d7-9371-beb064c98683}'):
            return True
        # 也可能是 explorer.exe 打开的控制面板子页面，带有 control panel 字样
        if '/control panel/' in s or '\\control panel\\' in s:
            return True
        return False
    
    def _check_folder_size_async(self, path, add_to_history):
        """异步检查文件夹大小并决定是否显示加载指示器"""
        # 显示加载指示器
        self._show_loading_indicator()
        self._folder_checker_done = False

        self._release_folder_checker(wait_ms=50)

        # 创建并启动检查线程
        self.folder_checker = FolderSizeChecker(path, self)
        checker = self.folder_checker

        def _on_checker_finished(p, count, is_large):
            if self._is_cleaning_up or self.folder_checker is not checker:
                return
            if getattr(self, '_folder_checker_done', False):
                self._release_folder_checker(wait_ms=0)
                return  # 已由超时保护处理
            self._folder_checker_done = True
            self._on_folder_size_checked(p, count, is_large, add_to_history)
            self._release_folder_checker(wait_ms=0)

        self.folder_checker.completed.connect(_on_checker_finished)
        self.folder_checker.start()

        # 超时保护：若线程在 FOLDER_CHECK_TIMEOUT+500ms 内未完成则强制导航
        # 防止云存储/网络路径的os.scandir()永久阻塞
        def _checker_timeout():
            if self._is_cleaning_up or self.folder_checker is not checker:
                return
            if getattr(self, '_folder_checker_done', True):
                return  # 线程已正常结束
            debug_print(f"[AsyncLoad] FolderSizeChecker timeout, forcing navigation: {path}")
            if hasattr(self, 'folder_checker') and self.folder_checker:
                self._release_folder_checker(wait_ms=50)
            self._folder_checker_done = True
            self._on_folder_size_checked(path, 0, False, add_to_history)

        QTimer.singleShot(FOLDER_CHECK_TIMEOUT + 500, _checker_timeout)
        debug_print(f"[AsyncLoad] Started checking folder: {path}")
    
    def _on_folder_size_checked(self, path, file_count, is_large, add_to_history):
        """文件夹大小检查完成的回调"""
        debug_print(f"[AsyncLoad] Folder checked: {path}, files={file_count}, large={is_large}")
        
        # 执行导航
        self._perform_navigation(path, add_to_history)
        
        # 隐藏加载指示器
        if is_large:
            # 大文件夹延迟隐藏指示器（等待Explorer加载）
            QTimer.singleShot(1000, self._hide_loading_indicator)
        else:
            # 小文件夹立即隐藏
            self._hide_loading_indicator()
    
    def _perform_navigation(self, path, add_to_history):
        """执行实际的导航操作"""
        path = self._normalize_local_path(path)
        self._mark_expected_navigation(path, is_shell=False)
        old_path = getattr(self, 'current_path', None)
        url = QDir.toNativeSeparators(path)
        
        # 立即更新路径栏（不等到最后，确保先更新 UI）
        if hasattr(self, 'path_bar'):
            self.path_bar.set_path(path)
        
        # 更新当前路径
        self.current_path = path
        
        # 导航到新目录时清除 Git 状态缓存
        self._git_status_cache = None
        
        # 停止目录轮询定时器，避免误触发刷新
        if hasattr(self, 'dir_poll_timer') and self.dir_poll_timer.isActive():
            self.dir_poll_timer.stop()
            debug_print(f"[Navigation] Stopped dir polling timer")
        
        # 设置标志，防止导航期间的自动刷新
        self._suppress_auto_refresh = True
        
        # 使用Navigate2来获得更好的控制
        try:
            # 尝试使用Navigate2获得更好的刷新效果
            self.explorer.dynamicCall("Navigate2(QVariant,QVariant,QVariant,QVariant,QVariant)", 
                                     url, 0, "", None, None)
        except Exception:
            # 回退到普通Navigate
            self.explorer.dynamicCall("Navigate(const QString&)", url)
        
        # 导航完成后清理标志并恢复路径同步
        if getattr(self, '_navigating_folder', False):
            self._navigating_folder = False
        self._resume_path_sync_after_navigation()

        # 更新状态栏
        self.update_explorer_status()
        
        # 更新文件系统监控（只监控真实文件系统路径）
        # 慢盘（网络/UNC/映射盘）：os.path.exists/os.path.isdir/addPath/_build_dir_snapshot
        # 均为同步文件系统调用，在挂起的网络路径上会阻塞 UI 线程导致整个程序卡死，
        # 故对慢盘完全跳过 watcher 注册与快照构建（此类路径的自动刷新本就依赖轮询兜底，
        # 而 _update_dir_polling 已对慢盘跳过轮询）。
        path_is_slow = self._is_slow_path(path)
        if hasattr(self, 'file_watcher') and not path_is_slow:
            # 移除旧路径的监控（旧路径若为慢盘同样跳过，避免 os.path.exists 阻塞）
            if (old_path and not self._is_slow_path(old_path) and
                    os.path.exists(old_path) and os.path.isdir(old_path) and
                    not old_path.startswith('shell:')):
                self._force_remove_watcher(old_path)
            # 添加新路径的监控
            if os.path.isdir(path):
                self._force_remove_watcher(path)
                if self.file_watcher.addPath(path):
                    debug_print(f"[FileWatcher] Now watching: {path}")
                else:
                    debug_print(f"[FileWatcher] Failed to watch: {path}")
                self._refresh_file_watch_paths(path)
                self._last_dir_snapshot = self._build_dir_snapshot(path)
                debug_print(f"[FileWatcher] Now watching: {self.file_watcher.directories()}")
        elif path_is_slow:
            # 慢盘不注册 watcher，清空上次快照，避免下次轮询用旧快照误判
            self._last_dir_snapshot = None
            debug_print(f"[FileWatcher] Skipped watcher for slow path: {path}")

        # 启用低频轮询兜底，处理 watcher 未报告的文件修改时间变化
        self._update_dir_polling(path)
        
        self.update_tab_title()
        if self.main_window and hasattr(self.main_window, 'get_current_tab_widget'):
            try:
                if self.main_window.get_current_tab_widget() is self:
                    self.main_window.update_chat_context()
            except Exception:
                pass
        # 添加到历史记录
        if add_to_history:
            self._add_to_history(path)
        
        # 延迟3秒后允许自动刷新（避免刚导航完就因为文件监视器触发刷新）
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(3000, lambda: setattr(self, '_suppress_auto_refresh', False))
    
    def _show_loading_indicator(self):
        """显示加载指示器（延迟显示）。

        快速导航（绝大多数情况）会在 hide 前取消该定时器，因此进度条不会出现，
        避免路径栏下方进度条一闪而过导致顶部区域"变高又恢复"的布局抖动。
        """
        if not hasattr(self, 'loading_bar'):
            return
        from PyQt5.QtCore import QTimer
        timer = getattr(self, '_loading_show_timer', None)
        if timer is None:
            timer = QTimer(self)
            timer.setSingleShot(True)
            timer.timeout.connect(self._do_show_loading_indicator)
            self._loading_show_timer = timer
        # 250ms 内完成的导航不显示进度条，消除闪烁
        timer.start(250)

    def _position_loading_bar(self):
        """将悬浮加载进度条定位到文件区域顶部（路径栏正下方），覆盖在 explorer 之上。"""
        if not hasattr(self, 'loading_bar'):
            return
        try:
            top = self.explorer.y() if hasattr(self, 'explorer') else 30
            self.loading_bar.setGeometry(0, top, self.width(), 20)
        except Exception:
            pass

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 悬浮加载进度条不在布局中，需随窗口尺寸变化手动跟随宽度
        if getattr(self, 'loading_bar', None) is not None and self.loading_bar.isVisible():
            self._position_loading_bar()

    def _do_show_loading_indicator(self):
        if hasattr(self, 'loading_bar'):
            banner = getattr(self, 'net_banner', None)
            if banner is not None and banner.isVisible():
                return  # 提示条已显示重试/连接进度
            self._position_loading_bar()
            self.loading_bar.show()
            self.loading_bar.raise_()
            debug_print("[AsyncLoad] Loading indicator shown")

    def _hide_loading_indicator(self):
        """隐藏加载指示器"""
        timer = getattr(self, '_loading_show_timer', None)
        if timer is not None:
            timer.stop()
        if hasattr(self, 'loading_bar'):
            self.loading_bar.hide()
            debug_print("[AsyncLoad] Loading indicator hidden")
    
    def _add_to_history(self, path):
        """添加路径到历史记录（应用内存优化限制）"""
        # 如果当前不在历史末尾，删除当前位置之后的所有历史
        if self.history_index < len(self.history) - 1:
            self.history = self.history[:self.history_index + 1]
        # 添加新路径（避免重复添加相同路径）
        if not self.history or self.history[-1] != path:
            self.history.append(path)
            self.history_index = len(self.history) - 1
            
            # 内存优化：限制历史记录长度
            if len(self.history) > MAX_NAVIGATION_HISTORY:
                # 删除最旧的记录
                remove_count = len(self.history) - MAX_NAVIGATION_HISTORY
                self.history = self.history[remove_count:]
                self.history_index = len(self.history) - 1
                debug_print(f"[Navigation History] Trimmed to {MAX_NAVIGATION_HISTORY} entries")
                
        # 更新主窗口按钮状态
        if self.main_window and hasattr(self.main_window, 'update_navigation_buttons'):
            self.main_window.update_navigation_buttons()
    
    def can_go_back(self):
        """是否可以后退"""
        return self.history_index > 0
    
    def can_go_forward(self):
        """是否可以前进"""
        return self.history_index < len(self.history) - 1
    
    def go_back(self):
        """后退到上一个位置"""
        if self.can_go_back():
            self.history_index -= 1
            path = self.history[self.history_index]
            is_shell = path.startswith('shell:')
            # 设置标志，防止sync时重复添加历史
            self._navigating_programmatically = True
            self.navigate_to(path, is_shell=is_shell, add_to_history=False)
            # 延迟重置标志，确保sync不会在导航完成前被触发
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1000, lambda: setattr(self, '_navigating_programmatically', False))
            # 更新主窗口按钮状态
            if self.main_window and hasattr(self.main_window, 'update_navigation_buttons'):
                self.main_window.update_navigation_buttons()
    
    def go_forward(self):
        """前进到下一个位置"""
        if self.can_go_forward():
            self.history_index += 1
            path = self.history[self.history_index]
            is_shell = path.startswith('shell:')
            # 设置标志，防止sync时重复添加历史
            self._navigating_programmatically = True
            self.navigate_to(path, is_shell=is_shell, add_to_history=False)
            # 延迟重置标志，确保sync不会在导航完成前被触发
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1000, lambda: setattr(self, '_navigating_programmatically', False))
            # 更新主窗口按钮状态
            if self.main_window and hasattr(self.main_window, 'update_navigation_buttons'):
                self.main_window.update_navigation_buttons()
    
    def on_directory_changed(self, path):
        """文件系统监控：目录内容发生变化（带防抖+风暴检测）"""
        import time, os
        current_time = time.time() * 1000  # 转为毫秒
        if self._is_slow_path(path):
            return
        _search_cache.clear()

        # ── 风暴计数：记录短时间内收到的事件数 ──
        storm_times = getattr(self, '_watcher_storm_times', [])
        storm_times = [t for t in storm_times if current_time - t < 10000]  # 10s 窗口
        storm_times.append(current_time)
        self._watcher_storm_times = storm_times
        is_storm = len(storm_times) > 5  # 10s 内超过 5 个事件视为风暴

        if hasattr(self, '_last_watcher_event'):
            last_time = self._last_watcher_event.get(path, 0)
            time_since_last = current_time - last_time
            if path == getattr(self, 'current_path', None) and self.refresh_timer.isActive():
                return
            # 风暴期间使用更长的防抖时间（5s），正常情况 3s
            debounce = 5000 if is_storm else self._watcher_debounce_ms
            if time_since_last < debounce:
                return
            self._last_watcher_event[path] = current_time
            if len(self._last_watcher_event) > 10:
                sorted_items = sorted(self._last_watcher_event.items(), key=lambda x: x[1], reverse=True)
                self._last_watcher_event = dict(sorted_items[:10])
        else:
            self._last_watcher_event = {path: current_time}
        # 事件风暴（批量拷贝/解压/删除等）期间：跳过 UI 线程上的同步快照扫描。
        # 内嵌视图在拷贝过程中会自行刷新；兜底轮询会在风暴平息后用快照确认变化。
        # 不能仅凭目录事件调用 Refresh()：调试日志等内部文件也会产生事件，并形成
        # “写日志 -> watcher -> Refresh -> 写日志”的循环，周期性清空用户选择。
        if not is_storm and path == self.current_path and getattr(self, '_refresh_active', True):
            self._poll_directory_changes()
            return
        debug_print(f"[FileWatcher] Directory changed: {path} (storm={is_storm})")
        if is_storm and path == self.current_path and getattr(self, '_refresh_active', True):
            debug_print("[FileWatcher] Storm refresh deferred to directory snapshot")
            return
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[FileWatcher] Auto-refresh suppressed during navigation")
            return
        if not getattr(self, '_refresh_active', True):
            self._request_refresh(reason="watcher")
            debug_print(f"[FileWatcher] Background tab marked dirty: {path}")
            return
        if path == self.current_path:
            self._request_refresh(reason="watcher")

    def _schedule_refresh(self, reason="manual"):
        """统一的刷新调度，避免重复代码"""
        import time
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[AutoRefresh] Suppressed during navigation (reason={reason})")
            return

        delay_ms = int(getattr(self, 'refresh_delay_ms', 500))
        min_interval_ms = int(getattr(self, '_refresh_min_interval_ms', 0))
        last_refresh_ms = float(getattr(self, '_last_refresh_ts_ms', 0) or 0)
        if min_interval_ms > 0 and last_refresh_ms > 0:
            elapsed_ms = (time.time() * 1000) - last_refresh_ms
            if elapsed_ms < min_interval_ms:
                delay_ms = max(delay_ms, int(min_interval_ms - elapsed_ms))

        if self.refresh_timer.isActive():
            remaining = self.refresh_timer.remainingTime()
            # 已有更早的刷新计划时不延后，避免连续事件导致“永远不刷新”
            if remaining > 0 and remaining <= delay_ms:
                return
            self.refresh_timer.stop()
            self.refresh_timer.start(delay_ms)
            debug_print(f"[AutoRefresh] Refresh timer adjusted to {delay_ms}ms (reason={reason})")
        else:
            debug_print(f"[AutoRefresh] Scheduling refresh in {delay_ms}ms (reason={reason})")
            self.refresh_timer.start(delay_ms)
    
    def delayed_refresh(self):
        """延迟刷新：避免频繁刷新。风暴期间跳过或进一步延后。"""
        import time
        if getattr(self, '_suppress_auto_refresh', False):
            debug_print(f"[FileWatcher] Auto-refresh suppressed during navigation")
            return
        if self._selection_guard_active():
            debug_print(f"[FileWatcher] Auto-refresh suppressed during selection guard")
            return
        now_ms = time.time() * 1000

        # 风暴检测：批量拷贝/解压/删除等会在短时间内触发大量目录事件
        storm_times = getattr(self, '_watcher_storm_times', [])
        storm_count = sum(1 for t in storm_times if now_ms - t < 10000)
        is_storm = storm_count > 5
        if is_storm:
            # 风暴进行中：不在 UI 线程调用同步 COM Refresh()，避免卡住所有标签。
            self.refresh_timer.start(2000)
            debug_print("[AutoRefresh] Storm active, skip sync COM refresh; will settle after storm")
            return

        if getattr(self, '_manual_refresh_frozen', False):
            debug_print(f"[AutoRefresh] Manually frozen, skipping refresh execution")
            return
        self._last_refresh_ts_ms = now_ms
        self._refresh_pending = False
        self._refresh_pending_reason = None
        debug_print(f"[FileWatcher] Auto-refreshing: {self.current_path}")
        if hasattr(self, 'explorer') and self.current_path:
            try:
                try:
                    self.explorer.dynamicCall('Refresh()')
                except Exception:
                    is_shell = self.current_path.startswith('shell:')
                    if is_shell:
                        self.explorer.dynamicCall('Navigate(const QString&)', self.current_path)
                    elif self.current_path.startswith('\\\\'):
                        self.explorer.dynamicCall('Navigate(const QString&)', self.current_path)
                    else:
                        url = 'file:///' + self.current_path.replace('\\', '/')
                        self.explorer.dynamicCall('Navigate2(const QVariant&)', url)
                debug_print(f"[FileWatcher] Refresh completed")
            except Exception as e:
                debug_print(f"[FileWatcher] Refresh error: {e}")
        # 刷新后更新快照（此处已排除风暴：风暴时前面已提前 return）
        self._poll_directory_changes()
        self.update_explorer_status()

    def _poll_directory_changes(self):
        """兜底轮询目录元数据，解决文件编辑后目录 watcher 不触发的问题"""
        # 如果设置了抑制标志，不触发刷新
        if getattr(self, '_suppress_auto_refresh', False):
            return
        if self._selection_guard_active():
            return
        if not getattr(self, '_refresh_active', True):
            return

        # 风暴期间跳过轮询（watcher 已在处理，避免额外 scandir 阻塞）
        import time
        now_ms = time.time() * 1000
        storm_times = getattr(self, '_watcher_storm_times', [])
        storm_count = sum(1 for t in storm_times if now_ms - t < 10000)
        if storm_count > 5:
            return

        path = self.current_path
        if not path or path.startswith(('shell:', '::')):
            return
        # OneDrive/网络路径的os.scandir()会阻塞UI线程，不做轮询
        if self._is_slow_path(path):
            return
        # 快照计算（scandir + 逐项 stat）放到后台线程池，避免大目录每 8s 在 UI 线程卡顿。
        # 上一次计算尚未返回时跳过本次，防止慢目录任务堆积。
        if getattr(self, '_snapshot_inflight', False):
            return
        try:
            if not hasattr(self, '_snapshot_signals') or self._snapshot_signals is None:
                self._snapshot_signals = _DirSnapshotSignals(self)
                self._snapshot_signals.done.connect(self._on_dir_snapshot_ready)
            self._snapshot_inflight = True
            runnable = _DirSnapshotRunnable(path, self._should_ignore_internal_dir_entry, self._snapshot_signals)
            QThreadPool.globalInstance().start(runnable)
        except Exception as e:
            self._snapshot_inflight = False
            debug_print(f"[DirPoll] Failed to start snapshot worker: {e}")

    def _on_dir_snapshot_ready(self, snap_path, current_snapshot):
        """后台快照计算完成（回到 UI 线程）：与上次快照比较，变化则调度刷新。"""
        self._snapshot_inflight = False
        if getattr(self, '_is_cleaning_up', False):
            return
        # 计算期间已切换目录：丢弃过期结果
        if snap_path != getattr(self, 'current_path', None):
            return
        if current_snapshot is None:
            return
        if self._last_dir_snapshot is None:
            self._last_dir_snapshot = current_snapshot
            self._request_refresh(reason="snapshot_ready")
            return
        if current_snapshot != self._last_dir_snapshot:
            debug_print(f"[DirPoll] Detected snapshot change for {snap_path}, scheduling refresh")
            self._last_dir_snapshot = current_snapshot
            self._request_refresh(reason="poll")

    def _update_dir_polling(self, path):
        """根据当前路径启动或停止兜底轮询"""
        if not hasattr(self, 'dir_poll_timer'):
            return
        # OneDrive/网络路径不做轮询，避免os.scandir()阻塞UI线程
        if path and self._is_slow_path(path):
            if self.dir_poll_timer.isActive():
                self.dir_poll_timer.stop()
                debug_print(f"[DirPoll] Stopped polling for slow path: {path}")
            return
        if path and os.path.isdir(path):
            # 立即更新快照，避免刚导航时误判为变化
            self.dir_mtime = self._get_dir_mtime(path)
            self._last_dir_snapshot = self._build_dir_snapshot(path)
            debug_print(f"[DirPoll] Updated mtime for {path}: {self.dir_mtime}")
            if getattr(self, '_refresh_active', False) and not self.dir_poll_timer.isActive():
                self.dir_poll_timer.start()
                debug_print(f"[DirPoll] Started polling {path}")
        else:
            if self.dir_poll_timer.isActive():
                self.dir_poll_timer.stop()
                debug_print(f"[DirPoll] Stopped polling")
            self.dir_mtime = None
            self._last_dir_snapshot = None

    def _get_dir_mtime(self, path):
        """安全获取目录修改时间"""
        try:
            return os.stat(path).st_mtime
        except Exception:
            return None

    def update_explorer_status(self):
        """更新嵌入 Explorer 下方状态栏（仅显示 Git 状态）"""
        if not hasattr(self, 'status_bar'):
            return
        try:
            worker = getattr(self, '_file_op_worker', None)
            if worker and worker.isRunning():
                # 后台复制/删除进行中时，保留进度文案，避免被 Git 状态刷新覆盖。
                return
        except Exception:
            pass
        path = getattr(self, 'current_path', None)
        if not path or path.startswith('shell:') or '::' in path:
            self.status_bar.setText('')
            return

        # 只显示 Git 状态摘要（仅读缓存，不阻塞 UI）
        cache = getattr(self, '_git_status_cache', None)
        git_summary = cache.get('result') if (cache and cache.get('path') == path) else None
        self.status_bar.setText(git_summary or '')
        # 异步刷新 Git 状态（后台线程）
        self._request_git_status_async(path)

    def _get_selection_entries(self):
        """返回选中条目列表，每项包含 is_file 与 size"""
        try:
            if isinstance(self.explorer, IExplorerBrowserWidget):
                paths = self._get_ieb_selected_paths()
                if not paths:
                    return []
                entries = []
                collect_file_sizes = len(paths) <= STATUS_SELECTION_METADATA_LIMIT
                for p in paths:
                    if not p:
                        continue
                    path_str = str(p)
                    is_file = os.path.isfile(path_str)
                    size = None
                    if is_file and collect_file_sizes:
                        try:
                            size = os.path.getsize(path_str)
                        except Exception:
                            size = None
                    entries.append({
                        'path': path_str,
                        'is_file': is_file,
                        'size': size,
                    })
                return entries

            doc = self.explorer.querySubObject('Document')
            if not doc:
                return []
            selected = doc.querySubObject('SelectedItems()')
            if not selected:
                return []
            count = selected.dynamicCall('Count()')
            if not count or count <= 0:
                return []
            count_int = int(count)
            collect_file_sizes = count_int <= STATUS_SELECTION_METADATA_LIMIT
            entries = []
            for i in range(count_int):
                item = selected.querySubObject('Item(int)', i)
                if not item:
                    continue
                path = item.dynamicCall('Path()')
                name = item.dynamicCall('Name()')
                if not path and name:
                    # 部分场景仅返回名称，尝试拼接
                    path = os.path.join(self.current_path, str(name))
                if not path:
                    continue
                path_str = str(path)
                is_file = os.path.isfile(path_str)
                size = None
                if is_file and collect_file_sizes:
                    try:
                        size = os.path.getsize(path_str)
                    except Exception:
                        size = None
                entries.append({
                    'path': path_str,
                    'is_file': is_file,
                    'size': size,
                })
            return entries
        except Exception:
            return None

    def _get_selected_paths(self):
        entries = self._get_selection_entries()
        if not entries:
            return []
        return [entry.get('path') for entry in entries if entry and entry.get('path')]

    def _get_single_selected_path(self):
        # 先用廉价的计数（单次 COM Count）短路：仅在恪好选中 1 项时才逐项枚举，
        # 避免大目录 Ctrl+A 后右键时对成千上万项做 COM Item()+isfile 的无谓枚举。
        cnt = self._get_selected_count_safe()
        if cnt is not None and cnt != 1:
            return None
        paths = self._get_selected_paths()
        if len(paths) == 1:
            return paths[0]
        return None

    def _launch_selected_file_with_program(self, file_path, program_path, display_name):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("当前选中项不是可打开的文件"), level="warning")
            return False
        if not program_path:
            show_toast(self, tr("提示"), tr("未找到 {}").format(display_name), level="warning")
            return False
        try:
            launch_detached_async([program_path, file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法使用 {} 打开文件: {}").format(display_name, e), level="error")
            return False

    def open_selected_with_default_app(self, file_path):
        if not file_path or not os.path.exists(file_path):
            show_toast(self, tr("提示"), tr("当前选中项不可打开"), level="warning")
            return False
        try:
            if os.path.isdir(file_path):
                if self.main_window and hasattr(self.main_window, 'add_new_tab'):
                    self.main_window.add_new_tab(file_path)
            elif os.name == 'nt':
                os.startfile(file_path)
            else:
                launch_detached([file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开选中项: {}").format(e), level="error")
            return False

    def open_selected_with_notepad(self, file_path):
        return self._launch_selected_file_with_program(file_path, 'notepad.exe', tr('记事本'))

    def open_selected_with_notepad_plus_plus(self, file_path):
        return self._launch_selected_file_with_program(file_path, getattr(self, 'notepad_plus_plus_path', None), 'Notepad++')

    def open_selected_with_system_dialog(self, file_path):
        if not file_path or not os.path.isfile(file_path):
            show_toast(self, tr("提示"), tr("当前选中项不是可打开的文件"), level="warning")
            return False
        if os.name != 'nt':
            show_toast(self, tr("提示"), tr("当前系统不支持打开“选择其他应用”对话框"), level="warning")
            return False
        try:
            launch_detached_async(['rundll32.exe', 'shell32.dll,OpenAs_RunDLL', file_path], cwd=os.path.dirname(file_path) or None)
            return True
        except Exception as e:
            show_toast(self, tr("错误"), tr("无法打开“选择其他应用”对话框: {}").format(e), level="error")
            return False

    def open_selected_parent_folder(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return False
        if os.path.isdir(file_path):
            folder_path = file_path
            select_file = None
        else:
            folder_path = os.path.dirname(file_path)
            select_file = os.path.basename(file_path)
        if self.main_window and hasattr(self.main_window, 'add_new_tab'):
            self.main_window.add_new_tab(folder_path, select_file=select_file)
            return True
        return False

    def show_selected_item_context_menu(self, global_pos):
        selected_paths = self._get_selected_paths()
        if not selected_paths:
            return False
        selected_paths = [p for p in selected_paths if p and os.path.exists(p)]
        if not selected_paths:
            return False

        file_path = selected_paths[0] if len(selected_paths) == 1 else None
        if not file_path:
            return False


        menu = QMenu(self)
        if file_path:
            open_action = menu.addAction(tr("打开"))
            open_action.triggered.connect(lambda: self.open_selected_with_default_app(file_path))

            open_folder_action = menu.addAction(tr("打开所在目录"))
            open_folder_action.triggered.connect(lambda: self.open_selected_parent_folder(file_path))

            if os.path.isfile(file_path):
                default_action = menu.addAction(tr("用系统默认程序打开"))
                default_action.triggered.connect(lambda: self.open_selected_with_default_app(file_path))

                menu.addSeparator()

                notepad_action = menu.addAction(tr("用记事本打开"))
                notepad_action.triggered.connect(lambda: self.open_selected_with_notepad(file_path))

                notepadpp_action = menu.addAction(tr("用 Notepad++ 打开"))
                notepadpp_action.setEnabled(bool(getattr(self, 'notepad_plus_plus_path', None)))
                if not getattr(self, 'notepad_plus_plus_path', None):
                    notepadpp_action.setToolTip(tr("未检测到 Notepad++"))
                notepadpp_action.triggered.connect(lambda: self.open_selected_with_notepad_plus_plus(file_path))

                menu.addSeparator()

                system_dialog_action = menu.addAction(tr("选择其他应用..."))
                system_dialog_action.triggered.connect(lambda: self.open_selected_with_system_dialog(file_path))

            menu.addSeparator()

        menu.exec_(global_pos)
        return True

    def quick_copy_selected_paths(self, selected_paths):
        if not selected_paths:
            show_toast(self, tr("提示"), tr("未选择"), level="warning")
            return
        if getattr(self, '_file_op_worker', None) and self._file_op_worker.isRunning():
            show_toast(self, tr("提示"), tr("已有后台文件操作进行中，请稍后"), level="warning")
            return

        from PyQt5.QtWidgets import QFileDialog
        initial_dir = self.current_path if os.path.isdir(getattr(self, 'current_path', '')) else QDir.homePath()
        dst_dir = QFileDialog.getExistingDirectory(self, tr("选择目标文件夹"), initial_dir)
        if not dst_dir:
            return

        self._run_file_batch_op('copy', selected_paths, dst_dir)

    def confirm_delete_selected_paths(self, selected_paths):
        if not selected_paths:
            show_toast(self, tr("提示"), tr("未选择"), level="warning")
            return
        if getattr(self, '_file_op_worker', None) and self._file_op_worker.isRunning():
            show_toast(self, tr("提示"), tr("已有后台文件操作进行中，请稍后"), level="warning")
            return

        from PyQt5.QtWidgets import QMessageBox
        total = len(selected_paths)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle(tr("确认删除"))
        box.setText(tr("将选中的 {} 项移入回收站？").format(total))
        box.setDetailedText('\n'.join(selected_paths))
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.setDefaultButton(QMessageBox.No)
        box.setWindowModality(Qt.NonModal)
        box.setAttribute(Qt.WA_DeleteOnClose, True)
        box.finished.connect(lambda result, paths=list(selected_paths): self._on_delete_confirm_finished(result, paths))
        box.open()

    def _on_delete_confirm_finished(self, result, selected_paths):
        from PyQt5.QtWidgets import QMessageBox
        if result == QMessageBox.Yes:
            self._run_file_batch_op('delete', selected_paths)

    def _run_file_batch_op(self, op_type, selected_paths, dst_dir=None, conflict_actions=None):
        if getattr(self, '_file_op_worker', None) and self._file_op_worker.isRunning():
            show_toast(self, tr("提示"), tr("已有后台文件操作进行中，请稍后"), level="warning")
            return
        paths = list(dict.fromkeys(p for p in (selected_paths or []) if isinstance(p, str) and p))
        if not paths:
            show_toast(self, tr("提示"), tr("未找到可操作的文件或文件夹"), level="warning")
            return

        configured_workers = 0
        try:
            mw = getattr(self, 'main_window', None)
            cfg = getattr(mw, 'config', {}) if mw else {}
            configured_workers = int(cfg.get('file_op_max_workers', 0) or 0)
        except Exception:
            configured_workers = 0

        panel = self.main_window.get_file_task_panel()
        panel.start_task(op_type, paths, dst_dir, configured_workers, tab=self, conflict_actions=conflict_actions)
        if hasattr(self, 'cancel_file_op_btn') and self.cancel_file_op_btn:
            self.cancel_file_op_btn.setText(tr("取消"))
            self.cancel_file_op_btn.setEnabled(True)
            if getattr(self, '_bottom_statusbar_visible', True):
                self.cancel_file_op_btn.show()
            else:
                self.cancel_file_op_btn.hide()
        if op_type == 'copy':
            key = _hotkey_text(getattr(getattr(self, 'main_window', None), 'config', {}), 'cancel_file_op')
            message = (tr("后台复制已开始，可继续操作其他标签页（{} 可取消）").format(key) if key
                       else tr("后台复制已开始，可继续操作其他标签页"))
            show_toast(self, tr("提示"), message, level="info", duration=2600)
        elif op_type in ('delete', 'permanent_delete'):
            show_toast(self, tr("提示"), tr("删除任务已开始，可在文件任务面板查看进度"), level="info", duration=2600)

    def cancel_current_file_batch_op(self):
        worker = getattr(self, '_file_op_worker', None)
        if not worker or not worker.isRunning():
            return False
        try:
            worker.request_cancel()
            if hasattr(self, 'cancel_file_op_btn') and self.cancel_file_op_btn:
                self.cancel_file_op_btn.setText(tr("取消中"))
                self.cancel_file_op_btn.setEnabled(False)
            if hasattr(self, 'status_bar') and self.status_bar:
                self.status_bar.setText(tr("正在取消后台任务，请稍候..."))
            show_toast(self, tr("提示"), tr("已发送取消请求，正在尽快停止..."), level="info", duration=1800)
            return True
        except Exception:
            return False

    def _on_file_batch_op_progress(self, op_type, done_count, total_count, current_name):
        if total_count <= 0:
            return
        percent = int((done_count * 100) / total_count)
        current_label = current_name or ''
        worker = getattr(self, '_file_op_worker', None)
        ok_count = int(getattr(worker, 'ok_count', 0) or 0)
        fail_count = int(getattr(worker, 'fail_count', 0) or 0)
        started_at = float(getattr(worker, 'started_at', 0.0) or 0.0)
        elapsed = max(0.0, time.monotonic() - started_at) if started_at > 0 else 0.0
        elapsed_text = f"{int(elapsed // 60):02d}:{int(elapsed % 60):02d}"
        remain_count = max(0, total_count - done_count)
        if op_type == 'copy':
            msg = tr("后台复制: {}/{} ({}%) 剩余{} | 用时 {} | 成功{} 失败{} | 当前: {}" ).format(
                done_count, total_count, percent, remain_count, elapsed_text, ok_count, fail_count, current_label
            )
        elif op_type in ('delete', 'permanent_delete'):
            msg = tr("后台删除: {}/{} ({}%) 剩余{} | 用时 {} | 成功{} 失败{} | 当前: {}" ).format(
                done_count, total_count, percent, remain_count, elapsed_text, ok_count, fail_count, current_label
            )
        else:
            msg = tr("后台操作进度: {}/{} ({}%) {}" ).format(done_count, total_count, percent, current_label)

        try:
            if hasattr(self, 'status_bar') and self.status_bar:
                self.status_bar.setText(msg)
        except Exception:
            pass

    def _on_file_batch_op_finished(self, op_type, ok_count, fail_count, errors):
        worker = getattr(self, '_file_op_worker', None)
        cancelled = bool(getattr(worker, 'cancelled', False))
        self._file_op_worker = None
        if hasattr(self, 'cancel_file_op_btn') and self.cancel_file_op_btn:
            self.cancel_file_op_btn.hide()
            self.cancel_file_op_btn.setText(tr("取消"))
            self.cancel_file_op_btn.setEnabled(True)

        try:
            self.update_explorer_status()
            self._request_refresh(reason='custom_file_op')
        except Exception:
            pass

    def select_file_in_explorer(self, filename, retries=6, delay_ms=250):
        """在Explorer控件中选中当前目录下指定的文件或文件夹。"""
        try:
            debug_print(f"[SelectFile] Attempting to select file: {filename}")
            
            # 构建完整路径
            full_path = os.path.join(self.current_path, filename)
            if not os.path.exists(full_path):
                debug_print(f"[SelectFile] File not found: {full_path}")
                return False

            if self._select_ieb_item_by_name(filename):
                return True
            
            # 使用Windows API选中文件（通过查找ListView控件并发送消息）
            try:
                import ctypes
                from ctypes import wintypes
                
                # Windows API常量
                LVM_SETITEMSTATE = 0x102B
                LVM_ENSUREVISIBLE = 0x1013
                LVM_GETITEMCOUNT = 0x1004
                LVM_GETITEMTEXT = 0x102D
                LVIF_STATE = 0x0008
                LVIS_SELECTED = 0x0002
                LVIS_FOCUSED = 0x0001
                
                # 获取当前窗口句柄
                hwnd = int(self.explorer.winId())
                
                # 查找ListView控件（通常类名是 SysListView32）
                user32 = ctypes.windll.user32
                
                def enum_child_windows(parent_hwnd):
                    """枚举所有子窗口"""
                    handles = []
                    def callback(hwnd, lParam):
                        handles.append(hwnd)
                        return True
                    
                    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
                    enum_proc = WNDENUMPROC(callback)
                    user32.EnumChildWindows(parent_hwnd, enum_proc, 0)
                    return handles
                
                # 查找ListView控件
                listview_hwnd = None
                for child_hwnd in enum_child_windows(hwnd):
                    class_name = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(child_hwnd, class_name, 256)
                    if 'SysListView32' in class_name.value:
                        listview_hwnd = child_hwnd
                        debug_print(f"[SelectFile] Found ListView control: {listview_hwnd}")
                        break
                
                if not listview_hwnd:
                    debug_print(f"[SelectFile] ListView control not found")
                    if retries > 0:
                        QTimer.singleShot(delay_ms, lambda: self.select_file_in_explorer(filename, retries - 1, delay_ms))
                    return False
                
                # 获取ListView中的项目数
                item_count = user32.SendMessageW(listview_hwnd, LVM_GETITEMCOUNT, 0, 0)
                debug_print(f"[SelectFile] ListView has {item_count} items")
                if item_count <= 0:
                    if retries > 0:
                        debug_print(f"[SelectFile] ListView not ready, retrying... remaining={retries}")
                        QTimer.singleShot(delay_ms, lambda: self.select_file_in_explorer(filename, retries - 1, delay_ms))
                    return False
                
                # 遍历所有项目，查找匹配的文件名
                for i in range(item_count):
                    # 获取项目文本需要使用进程间通信（因为ListView在不同进程）
                    # 这里使用简化方法：通过Document接口获取文件列表来匹配索引
                    try:
                        doc = self.explorer.querySubObject('Document')
                        if doc:
                            folder = doc.querySubObject('Folder')
                            if folder:
                                items = folder.querySubObject('Items()')
                                if items and i < items.dynamicCall('Count()'):
                                    item = items.querySubObject('Item(int)', i)
                                    if item:
                                        item_name = item.dynamicCall('Name()')
                                        if item_name == filename:
                                            debug_print(f"[SelectFile] Found file at index {i}: {filename}")
                                            
                                            # 定义LVITEM结构
                                            class LVITEM(ctypes.Structure):
                                                _fields_ = [
                                                    ('mask', wintypes.UINT),
                                                    ('iItem', ctypes.c_int),
                                                    ('iSubItem', ctypes.c_int),
                                                    ('state', wintypes.UINT),
                                                    ('stateMask', wintypes.UINT),
                                                    ('pszText', wintypes.LPWSTR),
                                                    ('cchTextMax', ctypes.c_int),
                                                    ('iImage', ctypes.c_int),
                                                    ('lParam', wintypes.LPARAM),
                                                ]
                                            
                                            # 取消所有项目的选中状态
                                            for j in range(item_count):
                                                lvi = LVITEM()
                                                lvi.mask = LVIF_STATE
                                                lvi.state = 0
                                                lvi.stateMask = LVIS_SELECTED | LVIS_FOCUSED
                                                user32.SendMessageW(listview_hwnd, LVM_SETITEMSTATE, j, ctypes.byref(lvi))
                                            
                                            # 选中并聚焦目标项
                                            lvi = LVITEM()
                                            lvi.mask = LVIF_STATE
                                            lvi.state = LVIS_SELECTED | LVIS_FOCUSED
                                            lvi.stateMask = LVIS_SELECTED | LVIS_FOCUSED
                                            result = user32.SendMessageW(listview_hwnd, LVM_SETITEMSTATE, i, ctypes.byref(lvi))
                                            debug_print(f"[SelectFile] SendMessage result: {result}")
                                            
                                            # 确保可见
                                            user32.SendMessageW(listview_hwnd, LVM_ENSUREVISIBLE, i, 0)
                                            
                                            # 设置焦点到ListView
                                            self.activateWindow()
                                            self.raise_()
                                            if hasattr(self, 'explorer') and self.explorer:
                                                self.explorer.setFocus(Qt.OtherFocusReason)
                                            user32.SetFocus(listview_hwnd)
                                            
                                            debug_print(f"[SelectFile] Successfully selected file via API: {filename}")
                                            # 用户随后可能按 Enter 进入选中的文件夹。
                                            # NavigateComplete2 信号在某些环境下不可用，
                                            # 通过启动路径同步定时器来兜底捕获导航变化。
                                            self._suppress_auto_refresh = False
                                            self._arm_selection_guard()
                                            self.start_path_sync_timer(duration_ms=8000)
                                            return True
                    except Exception as e:
                        debug_print(f"[SelectFile] Error matching item {i}: {e}")
                        continue
                
                debug_print(f"[SelectFile] File not found in ListView: {filename}")
                if retries > 0:
                    debug_print(f"[SelectFile] Target not visible yet, retrying... remaining={retries}")
                    QTimer.singleShot(delay_ms, lambda: self.select_file_in_explorer(filename, retries - 1, delay_ms))
                return False
                
            except Exception as e:
                debug_print(f"[SelectFile] Windows API method failed: {e}")
                import traceback
                traceback.print_exc()
                return False
        
        except Exception as e:
            debug_print(f"[SelectFile] Error: {e}")
            import traceback
            traceback.print_exc()
            return False

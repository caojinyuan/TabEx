"""Navigation requests, asynchronous completion and failure recovery for a file tab."""

import os
import time

from PyQt5.QtCore import QDir, QTimer

from .constants import ASYNC_LOAD_ENABLED, ASYNC_NAV_TIMEOUT_MS, FOLDER_CHECK_TIMEOUT
from .debuglog import debug_print
from .i18n import tr
from .netstatus import (
    ERROR_CANCELLED, ERROR_CONNECTION_UNAVAIL, is_network_path, NAV_TIMEOUT,
    RECONNECTED_CODES, ReconnectWorker, win32_error,
)
from .paths import is_slow_path, translate_common_path
from .shellview import _file_url_to_path
from .system import launch_detached, launch_shell_tool, _retain_thread_until_finished
from .widgets import show_toast
from .workers import FolderSizeChecker


class NavigationMixin:
    """Keep the tab's public navigation API independent of its widget construction."""

    def _on_async_nav_started(self, path):
        """慢盘异步导航开始：显示 loading 并进入导航中锁定状态。"""
        self._nav_in_progress = True
        self._nav_in_progress_path = self._normalize_local_path(path)
        self._show_loading_indicator()
        request_id = getattr(self, '_navigation_request_id', 0)
        # 安全兜底：网络中断时 NavigateComplete2 可能永不触发，超时后强制解锁并隐藏 loading
        QTimer.singleShot(
            ASYNC_NAV_TIMEOUT_MS,
            lambda p=self._nav_in_progress_path, r=request_id: self._async_nav_safety_timeout(p, r),
        )

    def _on_async_nav_finished(self, path, ok):
        """慢盘异步导航结束（仅解析失败/无法访问时触发）：解锁并隐藏 loading；失败原因由提示条展示。"""
        if self._normalize_local_path(path) != getattr(self, '_nav_in_progress_path', None):
            return
        self._nav_in_progress = False
        self._hide_loading_indicator()

    def _async_nav_safety_timeout(self, path, request_id=None):
        """导航中锁定的兜底超时：仍在等待同一目标时强制解锁，并提示连接超时。

        后台解析若之后成功，导航完成会自动关闭提示条。"""
        if (self._is_cleaning_up or (request_id is not None
                and request_id != getattr(self, '_navigation_request_id', 0))):
            return
        if (getattr(self, '_nav_in_progress', False) and
                getattr(self, '_nav_in_progress_path', None) == path):
            debug_print(f"[AsyncNav] Safety timeout, clearing nav lock: {path}")
            self._nav_in_progress = False
            self._hide_loading_indicator()
            self._show_network_error(path, NAV_TIMEOUT)

    def _on_navigation_failed(self, path, hr):
        """导航失败：hr 非 0 为路径解析失败；hr 为 0 是 Shell 内部导航失败，仅对网络位置提示。"""
        if getattr(self, '_sync_follow_target', None) == os.path.normcase(os.path.normpath(path)):
            self._sync_follow_target = None
        # 可能在 Shell 的 COM 回调内触发，延后改布局；其间若有导航成功则忽略本次失败
        seq = getattr(self, '_nav_success_seq', 0)
        request_id = getattr(self, '_navigation_request_id', 0)
        QTimer.singleShot(0, lambda: self._handle_navigation_failure(path, hr, seq, request_id))

    def _handle_navigation_failure(self, path, hr, seq, request_id=None):
        if (self._is_cleaning_up or getattr(self, '_nav_success_seq', 0) != seq
                or (request_id is not None and request_id != getattr(self, '_navigation_request_id', 0))):
            return
        if not hr and not is_network_path(path):
            return
        self._nav_in_progress = False
        self._hide_loading_indicator()
        self._show_network_error(path, win32_error(hr) if hr else 0)

    def _on_navigation_succeeded(self, path=None):
        self._nav_success_seq = getattr(self, '_nav_success_seq', 0) + 1
        checker = getattr(self, 'folder_checker', None)
        duplicate = (path is not None and self._normalize_local_path(path)
                     == self._normalize_local_path(getattr(self, 'current_path', '')))
        # Shell may finish refreshing the old view while a new folder check is queued.
        if checker is None or not duplicate:
            self._navigation_request_id = getattr(self, '_navigation_request_id', 0) + 1
            if checker is not None:
                self._release_folder_checker(wait_ms=0)
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
        # 错误占位隐藏的 Shell 控件必须先恢复，否则 IEB 会把本次导航延后到控件显示时。
        self._set_explorer_placeholder(False)
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
                self._on_navigation_succeeded(local_path)
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

        # 正常显示同一路径时避免重复刷新；失败后地址栏重输同一路径应允许重试。
        current_norm = getattr(self, 'current_path', '').replace('/', '\\').lower().rstrip('\\')
        path_norm = path.replace('/', '\\').lower().rstrip('\\')
        banner = getattr(self, 'net_banner', None)
        failed_navigation = banner is not None and not banner.isHidden()
        if path_norm == current_norm and not failed_navigation:
            debug_print(f"[PathBar] on_path_bar_changed: same as current_path, skip navigate")
            return

        # UNC 路径不做中英文目录互转：互转会同步探测路径是否存在，服务器不可达时长时间卡住界面
        if self._is_slow_path(path):
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

        self._navigation_request_id = getattr(self, '_navigation_request_id', 0) + 1
        self._nav_in_progress = False
        self._nav_in_progress_path = None
        # 快速取消所有待处理的双击检查定时器
        if hasattr(self, '_pending_double_click_timers'):
            for timer in self._pending_double_click_timers:
                try:
                    timer.stop()
                except Exception:
                    pass
            self._pending_double_click_timers = []

        # 停止之前的文件夹检查线程（减少等待时间）
        if getattr(self, 'folder_checker', None) is not None:
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
            self._sync_follow_target = None
            if is_network_path(path):
                # 盘符仍记录为网络驱动器但当前未连接
                self._show_network_error(path, ERROR_CONNECTION_UNAVAIL)

    def _is_slow_path(self, path):
        """检测OneDrive/网络/映射网络驱动器路径——这类路径上os.scandir()可能阻塞UI线程"""
        return is_slow_path(path)

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
        request_id = getattr(self, '_navigation_request_id', 0)

        def _on_checker_finished(p, count, is_large):
            if (self._is_cleaning_up or self.folder_checker is not checker
                    or request_id != getattr(self, '_navigation_request_id', 0)):
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
            if (self._is_cleaning_up or self.folder_checker is not checker
                    or request_id != getattr(self, '_navigation_request_id', 0)):
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

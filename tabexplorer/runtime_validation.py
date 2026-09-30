"""Opt-in native smoke/soak checks, restricted to isolated validation data."""

import ctypes
import json
import os
from pathlib import Path
import threading
import time

from PyQt5.QtCore import QObject, QTimer
from PyQt5.QtWidgets import QApplication

from .paths import get_app_data_path
from .system import get_process_memory_usage_mb


def _handles():
    from ctypes import wintypes
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.GetCurrentProcess.restype = wintypes.HANDLE
    kernel.GetProcessHandleCount.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    count = wintypes.DWORD()
    return count.value if kernel.GetProcessHandleCount(kernel.GetCurrentProcess(), ctypes.byref(count)) else None


class RuntimeValidation(QObject):
    def __init__(self, window, started_at):
        super().__init__(window)
        self.window = window
        self.started_at = float(os.environ.get('TABEX_VALIDATION_STARTED', started_at))
        self.root = Path(get_app_data_path())
        self.report = Path(os.environ['TABEX_VALIDATION_REPORT'])
        self.duration = max(5.0, float(os.environ.get('TABEX_VALIDATION_SECONDS', '10')))
        self.target_tabs = max(2, min(100, int(os.environ.get('TABEX_VALIDATION_TABS', '10'))))
        self.deadline = time.monotonic() + max(40, self.duration + 30)
        self.latencies = []
        self.samples = []
        self.actions = 0
        self.ready_at = None
        self.last_tick = time.monotonic()
        self.last_action = self.last_tick
        self.expected_path = str(self.root)
        self.next_path = None
        self.navigation_pending = False
        self.restored = False
        self.timer = QTimer(self)
        self.timer.setInterval(20)
        self.timer.timeout.connect(self.tick)
        self.timer.start()

    def sample(self):
        return {'elapsed_s': round(time.time() - self.started_at, 3),
                'rss_mb': get_process_memory_usage_mb(), 'handles': _handles(),
                'python_threads': len(threading.enumerate()), 'tabs': self.window.tab_widget.count()}

    def is_ready(self):
        tab = self.window.get_current_tab_widget()
        explorer = getattr(tab, 'explorer', None)
        if not getattr(explorer, '_init_ok', False) or not getattr(explorer, '_browser', None):
            return False
        path = getattr(tab, 'current_path', '')
        location = explorer._url_to_path(explorer.property('LocationURL'))
        expected = os.path.normcase(os.path.normpath(self.expected_path))
        return (os.path.normcase(os.path.normpath(path)) == expected
            and os.path.normcase(os.path.normpath(location or '')) == expected)

    def tick(self):
        now = time.monotonic()
        self.latencies.append(max(0, (now - self.last_tick) * 1000 - 20))
        if len(self.latencies) > 100000:
            self.latencies = self.latencies[-50000:]
        self.last_tick = now
        try:
            if now > self.deadline:
                self.finish('Native view did not settle before timeout')
                return
            if not self.is_ready():
                return
            if self.ready_at is None:
                self.ready_at = now
                self.first_ready_seconds = time.time() - self.started_at
                self.samples.append(self.sample())
                self.report.with_suffix('.ready').write_text('ready', encoding='ascii')
            if now - self.last_action < 0.2:
                return
            self.last_action = now
            if self.navigation_pending:
                self.actions += 1
                self.navigation_pending = False
            if now - self.ready_at >= self.duration and self.actions >= 4:
                self.finish('')
                return
            if self.next_path is not None:
                self.expected_path, self.next_path = self.next_path, None
                self.navigation_pending = True
                self.window.get_current_tab_widget().navigate_to(self.expected_path, skip_async_check=True)
                return
            if not self.restored:
                self.restored = True
                self.window.showMinimized()
                QTimer.singleShot(100, self.window.showMaximized)
                return
            if self.window.tab_widget.count() < self.target_tabs:
                self.window.add_new_tab(str(self.root), activate=False)
            else:
                index = self.actions % self.window.tab_widget.count()
                self.window.tab_widget.setCurrentIndex(index)
                self.expected_path = self.window.get_current_tab_widget().current_path
                self.next_path = str(self.root / 'navigation') if self.actions % 2 else str(self.root)
            if len(self.samples) < 600 and (len(self.samples) < 2 or now - self.ready_at > len(self.samples) * 30):
                self.samples.append(self.sample())
        except Exception as error:
            self.finish(f'{type(error).__name__}: {error}')

    def finish(self, error):
        self.timer.stop()
        self.samples.append(self.sample())
        ordered = sorted(self.latencies) or [0]
        report = {
            'ok': not error, 'error': error, 'frozen': bool(getattr(__import__('sys'), 'frozen', False)),
            'first_ready_s': round(getattr(self, 'first_ready_seconds', 0), 3),
            'actions': self.actions, 'minimize_restore': self.restored,
            'heartbeat_p95_ms': round(ordered[min(len(ordered) - 1, int(len(ordered) * .95))], 2),
            'heartbeat_p99_ms': round(ordered[min(len(ordered) - 1, int(len(ordered) * .99))], 2),
            'heartbeat_max_ms': round(max(ordered), 2), 'samples': self.samples,
            'expected_path': self.expected_path if error else '',
            'actual_path': getattr(self.window.get_current_tab_widget(), 'current_path', '') if error else '',
        }
        self.report.parent.mkdir(parents=True, exist_ok=True)
        self.window.grab().save(str(self.report.with_suffix('.png')))
        self.report.write_text(json.dumps(report, indent=2), encoding='utf-8')
        def close_after_navigation():
            self.window.close()
            if error:
                QApplication.instance().exit(1)
        QTimer.singleShot(500, close_after_navigation)
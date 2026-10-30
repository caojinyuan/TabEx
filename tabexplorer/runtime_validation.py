"""Opt-in native smoke/soak checks, restricted to isolated validation data."""

from collections import Counter
import json
import math
import os
from pathlib import Path
import threading
import time

from PyQt5.QtCore import QObject, QTimer
from PyQt5.QtWidgets import QApplication

from .paths import get_app_data_path
from .system import get_process_memory_usage_mb, get_process_resources


DEFAULT_LIMITS = {
    'commit_mb': 128, 'handles': 128, 'gdi_objects': 32, 'user_objects': 32,
    'python_threads': 2, 'dialogs': 0,
    'heartbeat_p95_ms': 200, 'heartbeat_p99_ms': 1000, 'heartbeat_max_ms': 5000,
}


def resource_violations(baseline, current, limits):
    errors = []
    for key in ('commit_mb', 'handles', 'gdi_objects', 'user_objects', 'python_threads', 'dialogs'):
        if not isinstance(baseline.get(key), (int, float)) or not isinstance(current.get(key), (int, float)):
            errors.append(f'Missing resource metric: {key}')
        elif current[key] - baseline[key] > limits[key]:
            errors.append(f'{key} growth {current[key] - baseline[key]:.2f} > {limits[key]}')
    return errors


class LatencyMetrics:
    """Bounded histogram over the entire run (10 ms buckets), not just its tail."""

    def __init__(self):
        self.buckets = Counter()
        self.count = 0
        self.maximum = 0

    def add(self, value):
        self.count += 1
        self.maximum = max(self.maximum, value)
        self.buckets[min(6000, math.ceil(value / 10))] += 1

    def summary(self):
        def percentile(fraction):
            cumulative = 0
            for bucket, count in sorted(self.buckets.items()):
                cumulative += count
                if cumulative >= max(1, self.count * fraction):
                    return bucket * 10
            return 0
        return {'heartbeat_p95_ms': percentile(.95), 'heartbeat_p99_ms': percentile(.99),
                'heartbeat_max_ms': round(self.maximum, 2)}


class RuntimeValidation(QObject):
    def __init__(self, window, started_at):
        super().__init__(window)
        self.window = window
        self.started_at = float(os.environ.get('TABEX_VALIDATION_STARTED', started_at))
        self.root = Path(get_app_data_path())
        self.report = Path(os.environ['TABEX_VALIDATION_REPORT'])
        self.duration = max(5.0, float(os.environ.get('TABEX_VALIDATION_SECONDS', '10')))
        self.target_tabs = max(2, min(100, int(os.environ.get('TABEX_VALIDATION_TABS', '10'))))
        self.scenario = os.environ.get('TABEX_VALIDATION_SCENARIO', 'smoke')
        self.minimum_cycles = int(os.environ.get('TABEX_VALIDATION_CYCLES', '3'))
        self.warmup_cycles = int(os.environ.get('TABEX_VALIDATION_WARMUP', '2'))
        self.settle_seconds = float(os.environ.get('TABEX_VALIDATION_SETTLE', '1'))
        self.limits = dict(DEFAULT_LIMITS, **json.loads(os.environ.get('TABEX_VALIDATION_LIMITS', '{}')))
        self.deadline = time.monotonic() + float(os.environ.get(
            'TABEX_VALIDATION_TIMEOUT', str(max(40, self.duration + 30))))
        self.latencies = LatencyMetrics()
        self.cycles = 0
        self.baseline = None
        self.violations = []
        self.phase = 'startup'
        self._in_tick = False
        self._finishing = False
        self._step = None
        if self.scenario == 'lifecycle':
            from .runtime_scenarios import lifecycle_steps
            self.steps = lifecycle_steps(self)
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
        from PyQt5.QtWidgets import QDialog
        return {'elapsed_s': round(time.time() - self.started_at, 3),
                'rss_mb': get_process_memory_usage_mb(), **(get_process_resources() or {}),
                'python_threads': len(threading.enumerate()), 'tabs': self.window.tab_widget.count(),
                'dialogs': len(self.window.findChildren(QDialog)), 'cycle': self.cycles}

    @staticmethod
    def same_path(actual, expected):
        return bool(actual) and os.path.normcase(os.path.normpath(actual)) == os.path.normcase(
            os.path.normpath(str(expected)))

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
        if self._in_tick or self._finishing:
            return
        self._in_tick = True
        now = time.monotonic()
        self.latencies.add(max(0, (now - self.last_tick) * 1000 - 20))
        self.last_tick = now
        try:
            if now > self.deadline:
                self.finish(f'Validation timeout: {self.phase}')
                return
            if self.scenario == 'lifecycle' and self.ready_at is not None:
                self.tick_lifecycle(now)
                return
            if not self.is_ready():
                return
            if self.ready_at is None:
                self.ready_at = now
                self.first_ready_seconds = time.time() - self.started_at
                self.samples.append(self.sample())
                self.report.with_suffix('.ready').write_text('ready', encoding='ascii')
                if self.scenario == 'lifecycle':
                    return
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
        finally:
            self._in_tick = False

    def tick_lifecycle(self, now):
        if self._step is not None:
            if now > self._step_deadline:
                raise TimeoutError(self.phase)
            if not self._step():
                return
            self.actions += 1
        self.phase, self._step = next(self.steps)
        self._step_deadline = now + 30
        if self._step is not None:
            return
        self.cycles += 1
        current = self.sample()
        if self.cycles == self.warmup_cycles:
            self.baseline = current
            self.measurement_started = now
            self.latencies = LatencyMetrics()
        elif self.baseline is not None:
            errors = resource_violations(self.baseline, current, self.limits)
            self.violations.extend(error for error in errors if error not in self.violations)
        # Preserve early and recent checkpoints while bounding multi-hour reports.
        if len(self.samples) >= 600:
            del self.samples[1]
        self.samples.append(current)
        if self.violations:
            self.finish('Resource budget exceeded')
        elif self.baseline is not None and self.cycles >= self.warmup_cycles + self.minimum_cycles \
                and now - self.measurement_started >= self.duration:
            self.finish('')

    def finish(self, error):
        if self._finishing:
            return
        self._finishing = True
        self.timer.stop()
        self.samples.append(self.sample())
        latencies = self.latencies.summary()
        if self.scenario == 'lifecycle' and self.baseline is not None:
            self.violations.extend(f'{key} {value} > {self.limits[key]}' for key, value in latencies.items()
                                   if value > self.limits[key])
            if self.violations and not error:
                error = 'Responsiveness budget exceeded'
        report = {
            'ok': not error, 'error': error, 'frozen': bool(getattr(__import__('sys'), 'frozen', False)),
            'first_ready_s': round(getattr(self, 'first_ready_seconds', 0), 3),
            'actions': self.actions, 'minimize_restore': self.restored,
            **latencies, 'samples': self.samples, 'scenario': self.scenario, 'phase': self.phase,
            'cycles': self.cycles, 'warmup_cycles': self.warmup_cycles,
            'measured_cycles': max(0, self.cycles - self.warmup_cycles),
            'baseline': self.baseline, 'limits': self.limits, 'violations': self.violations,
            'measurement_seconds': round(time.monotonic() - self.measurement_started, 2)
                if self.baseline is not None else 0,
            'resource_growth': {
                key: round(self.samples[-1][key] - self.baseline[key], 2)
                for key in ('commit_mb', 'handles', 'gdi_objects', 'user_objects', 'python_threads', 'dialogs')
                if key in self.baseline and key in self.samples[-1]
            } if self.baseline is not None else {},
            'expected_path': self.expected_path if error else '',
            'actual_path': getattr(self.window.get_current_tab_widget(), 'current_path', '') if error else '',
        }
        self.report.parent.mkdir(parents=True, exist_ok=True)
        self.window.grab().save(str(self.report.with_suffix('.png')))
        self.report.write_text(json.dumps(report, indent=2), encoding='utf-8')
        def close_after_navigation():
            panel = getattr(self.window, '_file_task_panel', None)
            if panel is not None and panel.has_running_tasks():
                for record in panel.records:
                    if not record['done']:
                        panel.table.setCurrentItem(record['row'])
                        panel.cancel_selected()
                QTimer.singleShot(100, close_after_navigation)
                return
            self.window.close()
            if error:
                QApplication.instance().exit(1)
        QTimer.singleShot(500, close_after_navigation)
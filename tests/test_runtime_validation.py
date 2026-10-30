"""Validation gates fail on measured resource/latency regressions, not only crashes."""

import unittest
import json
from pathlib import Path
import tempfile
import time
import types
from unittest.mock import Mock, patch

from app_modules import app
from tabexplorer.runtime_validation import DEFAULT_LIMITS, LatencyMetrics, RuntimeValidation, resource_violations


class RuntimeBudgetTests(unittest.TestCase):
    def test_resource_growth_boundary_and_missing_metrics(self):
        baseline = {'commit_mb': 80, 'handles': 600, 'gdi_objects': 90,
                    'user_objects': 120, 'python_threads': 3, 'dialogs': 1}
        current = {key: value + DEFAULT_LIMITS[key] for key, value in baseline.items()}
        self.assertEqual(resource_violations(baseline, current, DEFAULT_LIMITS), [])
        for key in baseline:
            with self.subTest(metric=key):
                excessive = dict(current)
                excessive[key] += 1
                self.assertTrue(any(key in error for error in resource_violations(
                    baseline, excessive, DEFAULT_LIMITS)))
        self.assertTrue(resource_violations(baseline, {}, DEFAULT_LIMITS))

    def test_latency_histogram_is_bounded_and_keeps_early_maximum(self):
        metrics = LatencyMetrics()
        metrics.add(70000)
        for _index in range(150000):
            metrics.add(1)
        summary = metrics.summary()
        self.assertEqual(summary['heartbeat_max_ms'], 70000)
        self.assertEqual(summary['heartbeat_p95_ms'], 10)
        self.assertEqual(summary['heartbeat_p99_ms'], 10)
        self.assertLessEqual(len(metrics.buckets), 6001)

    def test_percentiles_use_entire_run(self):
        metrics = LatencyMetrics()
        for value in [0] * 90 + [250] * 8 + [1500] * 2:
            metrics.add(value)
        self.assertEqual(metrics.summary(), {
            'heartbeat_p95_ms': 250, 'heartbeat_p99_ms': 1500, 'heartbeat_max_ms': 1500})

    def test_excessive_latency_marks_report_failed(self):
        metrics = LatencyMetrics()
        metrics.add(6000)
        with tempfile.TemporaryDirectory() as directory:
            host = types.SimpleNamespace(
                _finishing=False, timer=Mock(), samples=[], sample=lambda: {},
                latencies=metrics, scenario='lifecycle', baseline={}, violations=[],
                limits=DEFAULT_LIMITS, actions=2, restored=True, phase='resource checkpoint',
                cycles=3, warmup_cycles=2, expected_path='', window=Mock(),
                measurement_started=time.monotonic(),
                report=Path(directory) / 'report.json')
            host.window.get_current_tab_widget.return_value = types.SimpleNamespace(current_path='')
            with patch('tabexplorer.runtime_validation.QTimer.singleShot'):
                RuntimeValidation.finish(host, '')
            report = json.loads(host.report.read_text(encoding='utf-8'))
            self.assertFalse(report['ok'])
            self.assertEqual(report['error'], 'Responsiveness budget exceeded')
            self.assertTrue(any('heartbeat_max_ms' in error for error in report['violations']))
            self.assertEqual(report['measured_cycles'], 1)

    def test_invalid_budget_configuration_rejected_before_launch(self):
        from tools.validate_runtime import validate
        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / 'untouched.json'
            with patch('tools.validate_runtime.subprocess.Popen') as launch:
                for limits in ({'typo': 1}, {'handles': -1}, {'handles': float('nan')}, []):
                    with self.subTest(limits=limits), self.assertRaises(ValueError):
                        validate([], report, limits=limits)
                with self.assertRaises(ValueError):
                    validate([], report, timeout=-1)
                launch.assert_not_called()
                self.assertFalse(report.exists())

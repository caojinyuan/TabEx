"""Run source or packaged native checks without touching personal application data."""

import argparse
import json
import math
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def validate(command, report, seconds=10, tabs=10, scenario='smoke', cycles=3, warmup=2,
             settle=1, limits=None, timeout=None):
    from tabexplorer.runtime_validation import DEFAULT_LIMITS
    if limits is not None and (not isinstance(limits, dict) or set(limits) - set(DEFAULT_LIMITS)):
        raise ValueError('Unknown validation budget keys')
    budgets = dict(DEFAULT_LIMITS, **(limits or {}))
    if (scenario not in ('smoke', 'lifecycle') or seconds < 0 or cycles < 1 or warmup < 1
            or settle < 0 or any(not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0
                                 for value in [seconds, settle, *budgets.values()])):
        raise ValueError('Invalid validation scenario, duration, cycles or budgets')
    if timeout is not None and (not math.isfinite(timeout) or timeout <= 0):
        raise ValueError('Validation timeout must be positive and finite')
    timeout = timeout if timeout is not None else max(120, seconds + (cycles + warmup) * 45 + 60)
    report = Path(report).resolve()
    report.parent.mkdir(parents=True, exist_ok=True)
    report.unlink(missing_ok=True)
    report.with_suffix('.ready').unlink(missing_ok=True)
    with tempfile.TemporaryDirectory(prefix='tabex-validation-') as temporary:
        data = Path(temporary) / 'data'
        data.mkdir()
        (data / 'navigation' / 'nested').mkdir(parents=True)
        (data / 'fixture').mkdir()
        (data / 'fixture' / 'sample.bin').write_bytes(b'sample payload\n' * 65536)
        (data / 'destination').mkdir()
        (data / '.tabex-validation').write_text('isolated', encoding='ascii')
        (data / 'config.json').write_text(json.dumps({
            'enable_explorer_monitor': False, 'auto_update_check': False,
            'enable_cache_tabs': True, 'cached_tabs': [{'path': str(data)}],
            'pinned_tabs': [], 'debug_mode': False,
        }), encoding='utf-8')
        environment = dict(os.environ, TABEX_DATA_DIR=str(data), LOCALAPPDATA=str(Path(temporary) / 'local'),
                           TABEX_VALIDATION_REPORT=str(report), TABEX_VALIDATION_SECONDS=str(seconds),
                           TABEX_VALIDATION_TABS=str(tabs), TABEX_VALIDATION_STARTED=str(time.time()),
                           TABEX_VALIDATION_SCENARIO=scenario, TABEX_VALIDATION_CYCLES=str(cycles),
                           TABEX_VALIDATION_WARMUP=str(warmup), TABEX_VALIDATION_SETTLE=str(settle),
                           TABEX_VALIDATION_LIMITS=json.dumps(budgets), TABEX_VALIDATION_TIMEOUT=str(timeout),
                           PYTHONUTF8='1')
        with tempfile.TemporaryFile() as output:
            process = subprocess.Popen(command + ['--self-test'], cwd=ROOT, env=environment,
                                       stdout=output, stderr=subprocess.STDOUT)
            try:
                deadline = time.monotonic() + 40
                while process.poll() is None and not report.with_suffix('.ready').exists():
                    if time.monotonic() > deadline:
                        raise TimeoutError('Native startup did not become ready')
                    threading.Event().wait(.05)
                if process.poll() is None:
                    second = subprocess.run(command, cwd=ROOT, env=environment, stdin=subprocess.DEVNULL,
                                            stdout=output, stderr=subprocess.STDOUT, timeout=30)
                    if second.returncode:
                        raise RuntimeError('Second launch was not forwarded')
                code = process.wait(timeout=timeout + 15)
                if code != 0 or not report.exists():
                    output.seek(0)
                    details = report.read_text(encoding='utf-8') if report.exists() else ''
                    raise RuntimeError(f'Exit code: {code}\n' + details + '\n'
                                       + output.read(16384).decode('utf-8', errors='replace'))
                result = json.loads(report.read_text(encoding='utf-8'))
                sessions = list((Path(temporary) / 'local' / 'TabEx' / 'diagnostics').glob('*/session.json'))
                states = [json.loads(path.read_text(encoding='utf-8')).get('status') for path in sessions]
                result['clean_exit'] = bool(states) and all(state == 'closed' for state in states)
                result['single_instance'] = len(sessions) == 1
                result['ok'] = result['ok'] and result['clean_exit'] and result['single_instance']
                report.write_text(json.dumps(result, indent=2), encoding='utf-8')
                if not result['ok']:
                    raise RuntimeError(json.dumps(result))
                return result
            finally:
                if process.poll() is None:
                    subprocess.run(['taskkill', '/PID', str(process.pid), '/T', '/F'],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                    process.wait(timeout=10)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--exe', type=Path)
    parser.add_argument('--report', type=Path, default=ROOT / 'artifacts' / 'runtime.json')
    parser.add_argument('--seconds', type=float, default=10)
    parser.add_argument('--tabs', type=int, default=10)
    parser.add_argument('--scenario', choices=('smoke', 'lifecycle'), default='smoke')
    parser.add_argument('--cycles', type=int, default=3, help='Minimum measured lifecycle cycles')
    parser.add_argument('--warmup', type=int, default=2, help='Warmup cycles before resource baseline')
    parser.add_argument('--settle', type=float, default=1, help='Quiet seconds before each resource checkpoint')
    parser.add_argument('--timeout', type=float, help='Overall watchdog deadline in seconds')
    parser.add_argument('--limits', type=Path, help='JSON overrides for resource growth and heartbeat budgets')
    options = parser.parse_args()
    command = [str(options.exe.resolve())] if options.exe else [sys.executable, '-X', 'utf8', str(ROOT / 'TabEx.py')]
    limits = json.loads(options.limits.read_text(encoding='utf-8')) if options.limits else None
    result = validate(command, options.report, options.seconds, options.tabs, options.scenario,
                      options.cycles, options.warmup, options.settle, limits, options.timeout)
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()
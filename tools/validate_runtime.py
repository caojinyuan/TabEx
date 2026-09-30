"""Run source or packaged native checks without touching personal application data."""

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time


ROOT = Path(__file__).resolve().parents[1]


def validate(command, report, seconds=10, tabs=10):
    report = Path(report).resolve()
    report.parent.mkdir(parents=True, exist_ok=True)
    report.unlink(missing_ok=True)
    report.with_suffix('.ready').unlink(missing_ok=True)
    with tempfile.TemporaryDirectory(prefix='tabex-validation-') as temporary:
        data = Path(temporary) / 'data'
        data.mkdir()
        (data / 'navigation').mkdir()
        (data / '.tabex-validation').write_text('isolated', encoding='ascii')
        (data / 'config.json').write_text(json.dumps({
            'enable_explorer_monitor': False, 'auto_update_check': False,
            'enable_cache_tabs': True, 'cached_tabs': [{'path': str(data)}],
            'pinned_tabs': [], 'debug_mode': False,
        }), encoding='utf-8')
        environment = dict(os.environ, TABEX_DATA_DIR=str(data), LOCALAPPDATA=str(Path(temporary) / 'local'),
                           TABEX_VALIDATION_REPORT=str(report), TABEX_VALIDATION_SECONDS=str(seconds),
                           TABEX_VALIDATION_TABS=str(tabs), TABEX_VALIDATION_STARTED=str(time.time()),
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
                code = process.wait(timeout=max(60, seconds + 40))
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
    options = parser.parse_args()
    command = [str(options.exe.resolve())] if options.exe else [sys.executable, '-X', 'utf8', str(ROOT / 'TabEx.py')]
    result = validate(command, options.report, options.seconds, options.tabs)
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()
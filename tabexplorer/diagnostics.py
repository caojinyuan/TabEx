"""Bounded, local-only crash records and redacted support bundles."""

import faulthandler
import json
import os
from pathlib import Path
import platform
import re
import sys
import tempfile
import threading
import time
import traceback
import uuid
import zipfile


MAX_LOG_BYTES = 256 * 1024
MAX_SESSIONS = 5
SAFE_LOG_TAGS = (
    '[ClosedTabs]', '[TabSwitch]', '[DirPoll]', '[AutoRefresh]',
    '[FileWatcher]', '[FolderSizeChecker]', '[AsyncLoad]', '[IExplorerBrowser]',
    '[PidlResolver]', '[DebugLog]', 'Qt Message:', 'Qt Error:',
)


def redact(text):
    text = str(text)
    text = re.sub(r'(?i)\b(?:https?|ftp)://[^\s<>"\']+', '<URL>', text)
    text = re.sub(r'(?i)\bBearer\s+[^\s,;"\']+', 'Bearer <REDACTED>', text)
    text = re.sub(
        r'''(?ix)(["']?(?:api[_-]?key|access[_-]?token|refresh[_-]?token|token|password|passwd|secret|authorization)["']?\s*[:=]\s*)(?:"[^"]*"|'[^']*'|[^\s,;}]+)''',
        r'\1<REDACTED>', text)
    text = re.sub(r'(?:[A-Za-z]:[\\/]|\\\\)[^\r\n"\'<>]*', '<PATH>', text)
    text = re.sub(r'(?<![\w:])/(?:[^\s/]+/)+[^\s"\'<>]*', '<PATH>', text)
    text = re.sub(r'\b[^\s@]+@[^\s@]+\.[^\s@]+\b', '<EMAIL>', text)
    for value in (os.environ.get('USERNAME'), os.environ.get('USER'), os.environ.get('COMPUTERNAME')):
        if value and len(value) >= 3:
            text = re.sub(re.escape(value), '<IDENTITY>', text, flags=re.IGNORECASE)
    return text


def read_tail(path):
    with open(path, 'rb') as source:
        source.seek(0, os.SEEK_END)
        size = source.tell()
        source.seek(max(0, size - MAX_LOG_BYTES))
        data = source.read(MAX_LOG_BYTES)
    if size > MAX_LOG_BYTES:
        data = data.partition(b'\n')[2]
    return data.decode('utf-8', errors='replace')


def filtered_debug_log(path):
    lines = []
    for line in read_tail(path).splitlines():
        message = re.sub(r'^\[\d{4}-\d\d-\d\d \d\d:\d\d:\d\d\]\s*', '', line)
        if message.startswith(SAFE_LOG_TAGS):
            lines.append(redact(line))
    return '\n'.join(lines)


def _process_alive(pid):
    if not isinstance(pid, int) or pid <= 0:
        return False
    if os.name == 'nt':
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
        kernel.OpenProcess.restype = wintypes.HANDLE
        kernel.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
        kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
        handle = kernel.OpenProcess(0x1000, False, pid)
        if not handle:
            return ctypes.get_last_error() == 5
        try:
            code = wintypes.DWORD()
            return not kernel.GetExitCodeProcess(handle, ctypes.byref(code)) or code.value == 259
        finally:
            kernel.CloseHandle(handle)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True


def _write_json(path, value):
    descriptor, temporary = tempfile.mkstemp(prefix='.diagnostic-', dir=str(path.parent))
    try:
        with os.fdopen(descriptor, 'w', encoding='utf-8') as target:
            json.dump(value, target, ensure_ascii=True, indent=2)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)


class Diagnostics:
    def __init__(self, directory, version):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self.session = self.directory / (time.strftime('%Y%m%d-%H%M%S-') + uuid.uuid4().hex[:8])
        self.session.mkdir()
        self.version = version
        self.lock = threading.RLock()
        self.fault_file = None
        self.owns_fault_handler = False
        self.previous_hooks = None
        self.state = {
            'schema': 1, 'version': version, 'pid': os.getpid(),
            'started': time.strftime('%Y-%m-%dT%H:%M:%S'), 'status': 'running',
            'python': platform.python_version(), 'os': platform.system(),
            'os_release': platform.release(), 'architecture': platform.machine(),
            'frozen': bool(getattr(sys, 'frozen', False)),
        }
        _write_json(self.session / 'session.json', self.state)
        self._prune()

    def _prune(self):
        import shutil
        sessions = sorted(self.directory.glob('????????-??????-????????'), reverse=True)
        for session in sessions[MAX_SESSIONS:]:
            try:
                if session == self.session or session.is_symlink():
                    continue
                state = json.loads((session / 'session.json').read_text(encoding='utf-8'))
                if state.get('ended') or not _process_alive(state.get('pid')):
                    shutil.rmtree(session)
            except (OSError, ValueError):
                pass

    def record(self, kind, details):
        try:
            with self.lock:
                path = self.session / 'events.jsonl'
                if path.exists() and path.stat().st_size > MAX_LOG_BYTES:
                    os.replace(path, self.session / 'events.previous.jsonl')
                event = {'time': time.strftime('%Y-%m-%dT%H:%M:%S'),
                         'kind': kind, 'details': redact(str(details)[:8192])}
                with path.open('a', encoding='utf-8') as target:
                    target.write(json.dumps(event, ensure_ascii=True) + '\n')
                    target.flush()
        except OSError:
            pass

    def exception(self, kind, exception_type, exception, trace):
        frames = [{'file': Path(frame.filename).name, 'line': frame.lineno, 'function': frame.name}
                  for frame in traceback.extract_tb(trace)[-40:]]
        self.record(kind, json.dumps({'type': exception_type.__name__, 'frames': frames}))
        self.mark('error')

    def mark(self, status):
        try:
            with self.lock:
                self.state['status'] = status
                self.state['updated'] = time.strftime('%Y-%m-%dT%H:%M:%S')
                _write_json(self.session / 'session.json', self.state)
        except OSError:
            pass

    def snapshot(self, values):
        try:
            safe = json.loads(redact(json.dumps(values, ensure_ascii=False)))
            with self.lock:
                _write_json(self.session / 'snapshot.json', safe)
        except (OSError, ValueError):
            pass

    def sample(self, values):
        """追加一条资源趋势记录（仅计数），与异常事件分开滚动，避免挤掉崩溃记录。"""
        try:
            line = redact(json.dumps(dict(values, time=time.strftime('%Y-%m-%dT%H:%M:%S')), ensure_ascii=True))
            with self.lock:
                path = self.session / 'resources.jsonl'
                if path.exists() and path.stat().st_size > MAX_LOG_BYTES:
                    os.replace(path, self.session / 'resources.previous.jsonl')
                with path.open('a', encoding='utf-8') as target:
                    target.write(line + '\n')
        except (OSError, TypeError, ValueError):
            pass

    def install(self):
        if self.previous_hooks is not None:
            return
        self.previous_hooks = (sys.excepthook, threading.excepthook)

        def main_exception(exception_type, exception, trace):
            self.exception('python_exception', exception_type, exception, trace)
            self.previous_hooks[0](exception_type, exception, trace)

        def thread_exception(args):
            self.exception('thread_exception', args.exc_type, args.exc_value, args.exc_traceback)
            self.previous_hooks[1](args)

        sys.excepthook = main_exception
        threading.excepthook = thread_exception
        self.state['native_fault_capture'] = 'external-handler'
        try:
            self.fault_file = (self.session / 'fault.log').open('ab', buffering=0)
            if not faulthandler.is_enabled():
                faulthandler.enable(file=self.fault_file, all_threads=True)
                self.owns_fault_handler = True
                self.state['native_fault_capture'] = 'enabled'
        except (OSError, RuntimeError):
            self.state['native_fault_capture'] = 'unavailable'
            if self.fault_file is not None:
                self.fault_file.close()
                self.fault_file = None
        self.mark(self.state['status'])

    def capture_debug(self, path):
        try:
            (self.session / 'previous-debug-filtered.log').write_text(filtered_debug_log(path), encoding='utf-8')
        except OSError:
            pass

    def dump_threads(self):
        if self.fault_file is not None:
            try:
                faulthandler.dump_traceback(file=self.fault_file, all_threads=True)
            except (OSError, RuntimeError):
                pass

    def close(self):
        self.state['ended'] = time.strftime('%Y-%m-%dT%H:%M:%S')
        self.mark('closed' if self.state['status'] == 'running' else self.state['status'])
        if self.previous_hooks is not None:
            sys.excepthook, threading.excepthook = self.previous_hooks
            self.previous_hooks = None
        if self.fault_file is not None:
            if self.owns_fault_handler:
                faulthandler.disable()
                self.owns_fault_handler = False
            self.fault_file.close()
            self.fault_file = None

    def export(self, destination, debug_path=None):
        entries = {}
        warnings = []
        sessions = sorted(self.directory.glob('????????-??????-????????'), reverse=True)[:MAX_SESSIONS]
        if self.session not in sessions:
            sessions = [self.session] + sessions[:MAX_SESSIONS - 1]
        for session in sessions:
            if session.is_symlink():
                continue
            for name in ('session.json', 'snapshot.json', 'events.jsonl', 'events.previous.jsonl',
                         'resources.jsonl', 'resources.previous.jsonl',
                         'fault.log', 'previous-debug-filtered.log'):
                path = session / name
                if path.is_file() and not path.is_symlink():
                    try:
                        entries[f'sessions/{session.name}/{name}'] = redact(read_tail(path))
                    except OSError:
                        warnings.append(f'Unreadable: {session.name}/{name}')
        if debug_path:
            try:
                entries['debug-filtered.log'] = filtered_debug_log(debug_path)
            except OSError:
                warnings.append('Debug log unavailable')
        entries['manifest.json'] = json.dumps({
            'schema': 1, 'app_version': self.version, 'created': time.strftime('%Y-%m-%dT%H:%M:%S'),
            'files': sorted(entries), 'warnings': warnings,
            'privacy': 'Local export only. Paths, URLs and common credentials redacted. '
                       'No config, bookmarks, chat history, environment dump, file contents or memory dump. '
                       'Review before sharing; automatic redaction is not a secrecy guarantee.',
            'limits': 'Last 5 sessions; last 256 KiB per file. Running means no clean-exit marker, '
                      'not proof of a crash. Native fault capture is best effort.',
        }, ensure_ascii=True, indent=2)
        destination = Path(destination)
        descriptor, temporary = tempfile.mkstemp(prefix='.tabex-diagnostics-', dir=str(destination.parent))
        os.close(descriptor)
        try:
            with zipfile.ZipFile(temporary, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
                for name, content in entries.items():
                    archive.writestr(name, content.encode('utf-8'))
            os.replace(temporary, destination)
        finally:
            if os.path.exists(temporary):
                os.remove(temporary)
        return sorted(entries)
"""Shared JSON validation, recovery and atomic persistence."""

import json
import os
import sys
import time
import uuid

from .debuglog import _diagnostic_event


def _preserve_unreadable_file(path):
    backup = f"{path}.broken-{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}"
    try:
        os.rename(path, backup)
        return backup
    except OSError as error:
        print(f"[Data] Cannot preserve {path}: {error}", file=sys.stderr)
        return ''


class JsonFileStore:
    def __init__(self, path, root_type):
        self.path = os.fspath(path)
        self.root_type = root_type
        self.last_written = None
        self.recovered_backup = ''
        self.write_blocked = False
        self.last_error = ''
        self.error_handler = None
        self._error_notified = False

    def report_error(self, error):
        message = str(error)
        if message != self.last_error:
            print(f"[Data] {self.path}: {message}", file=sys.stderr)
            _diagnostic_event('data_error', f'{self.path}: {message}')
        self.last_error = message
        if self.error_handler is not None and not self._error_notified:
            self._error_notified = True
            self.error_handler(self.path, message)
        return False

    def read(self, default, validate=None):
        self.last_written = None
        self.recovered_backup = ''
        self.write_blocked = False
        self.last_error = ''
        self._error_notified = False
        try:
            with open(self.path, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
            if not isinstance(data, self.root_type) or (validate is not None and not validate(data)):
                raise ValueError('Invalid JSON data structure')
            return data
        except FileNotFoundError as error:
            if not os.path.lexists(self.path):
                return default
            self.write_blocked = True
            self.report_error(error)
        except (OSError, ValueError) as error:
            if isinstance(error, ValueError):
                self.recovered_backup = _preserve_unreadable_file(self.path)
            self.write_blocked = not bool(self.recovered_backup)
            self.report_error(error)
        return default

    def save(self, value, indent=2):
        if self.write_blocked:
            return self.report_error(self.last_error)
        temporary = self.path + '.tmp'
        created = False
        try:
            if not isinstance(value, self.root_type):
                raise ValueError('Invalid JSON data structure')
            content = json.dumps(value, ensure_ascii=False, indent=indent)
            previous = self.last_written
            if previous is None:
                try:
                    with open(self.path, 'r', encoding='utf-8') as handle:
                        previous = handle.read()
                except FileNotFoundError:
                    pass
            if previous != content:
                with open(temporary, 'w', encoding='utf-8') as handle:
                    created = True
                    handle.write(content)
                    handle.flush()
                    os.fsync(handle.fileno())
                os.replace(temporary, self.path)
            self.last_written = content
            self.last_error = ''
            self._error_notified = False
            return True
        except (OSError, TypeError, ValueError) as error:
            return self.report_error(error)
        finally:
            if created:
                try:
                    os.remove(temporary)
                except FileNotFoundError:
                    pass
                except OSError as error:
                    self.report_error(error)

"""Portable file-basket manifests. Paths are relative to the manifest directory."""

import hashlib
import json
import os
from pathlib import Path, PureWindowsPath
import re
import stat
import tempfile

from .i18n import tr

MAX_FILES = 10000
MAX_JSON_BYTES = 8 * 1024 * 1024
FORMAT = 'TabEx-SHA256-v1'


def _check_cancel(cancelled):
    if cancelled():
        raise InterruptedError(tr("已取消"))


def _member(root, relative):
    if not isinstance(relative, str) or not relative:
        raise ValueError(tr("清单内的路径必须是相对文件路径。"))
    winpath = PureWindowsPath(relative)
    parts = relative.replace('\\', '/').split('/')
    if winpath.anchor or any(p in ('', '.', '..') or os.path.isreserved(p) for p in parts):
        raise ValueError(tr("清单内的路径必须是相对文件路径。"))
    path = root.joinpath(*parts).resolve()
    if not path.is_relative_to(root):
        raise ValueError(tr("文件不在清单目录内：{}").format(relative))
    return path


def _digest(path, cancelled):
    _check_cancel(cancelled)
    if not path.is_file():
        raise ValueError(tr("不是可读取的普通文件：{}").format(path))
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        before = os.fstat(stream.fileno())
        if not stat.S_ISREG(before.st_mode):
            raise ValueError(tr("不是可读取的普通文件：{}").format(path))
        while True:
            _check_cancel(cancelled)
            block = stream.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
        after = path.stat()
    identity = lambda s: (s.st_dev, s.st_ino, s.st_size, s.st_mtime_ns)
    if identity(before) != identity(after):
        raise ValueError(tr("计算期间文件已改变：{}").format(path))
    return before.st_size, digest.hexdigest()


def export_manifest(destination, paths, cancelled=lambda: False):
    if not 0 < len(paths) <= MAX_FILES:
        raise ValueError(tr("清单必须包含 1 至 10000 个文件。"))
    destination = Path(destination).absolute()
    root = destination.parent.resolve()
    files, seen = [], set()
    for name in paths:
        _check_cancel(cancelled)
        path = Path(name).absolute()
        if not path.is_relative_to(destination.parent):
            raise ValueError(tr("请将清单保存到所有所选文件的共同上级目录。"))
        relative = path.relative_to(destination.parent).as_posix()
        resolved = _member(root, relative)
        if resolved == destination.resolve() or (destination.exists() and resolved.samefile(destination)):
            raise ValueError(tr("清单不能覆盖被校验的文件。"))
        key = relative.casefold()
        if key in seen:
            raise ValueError(tr("清单含有重复路径：{}").format(relative))
        seen.add(key)
        size, digest = _digest(resolved, cancelled)
        files.append({'path': relative, 'size': size, 'sha256': digest})
    encoded = json.dumps({'format': FORMAT, 'files': files}, ensure_ascii=False, indent=2).encode('utf-8')
    if len(encoded) > MAX_JSON_BYTES:
        raise ValueError(tr("清单过大（最多 8 MiB）。"))
    temp = None
    try:
        with tempfile.NamedTemporaryFile(dir=root, prefix='.tabex-sha-', delete=False) as f:
            temp = Path(f.name)
            f.write(encoded)
            f.flush()
            os.fsync(f.fileno())
        _check_cancel(cancelled)
        os.replace(temp, destination)
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)
    return len(files)


def verify_manifest(source, cancelled=lambda: False):
    source = Path(source).absolute()
    with source.open('rb') as stream:
        raw = stream.read(MAX_JSON_BYTES + 1)
    if len(raw) > MAX_JSON_BYTES:
        raise ValueError(tr("清单过大（最多 8 MiB）。"))
    data = json.loads(raw.decode('utf-8-sig'))
    files = data.get('files') if isinstance(data, dict) and data.get('format') == FORMAT else None
    if not isinstance(files, list) or not 0 < len(files) <= MAX_FILES:
        raise ValueError(tr("无效的 SHA-256 清单。"))
    root = source.parent.resolve()
    failures, seen = [], set()
    for entry in files:
        _check_cancel(cancelled)
        if (not isinstance(entry, dict) or type(entry.get('size')) is not int or entry['size'] < 0
                or not isinstance(entry.get('sha256'), str) or not re.fullmatch('[0-9a-fA-F]{64}', entry['sha256'])):
            raise ValueError(tr("无效的 SHA-256 清单。"))
        relative = entry.get('path')
        path = _member(root, relative)
        key = PureWindowsPath(relative).as_posix().casefold()
        if key in seen:
            raise ValueError(tr("清单含有重复路径：{}").format(relative))
        seen.add(key)
        try:
            size, digest = _digest(path, cancelled)
        except InterruptedError:
            raise
        except (OSError, ValueError) as error:
            failures.append(f'{relative}: {error}')
        else:
            if size != entry['size'] or digest != entry['sha256'].lower():
                failures.append(f'{relative}: ' + tr("大小或 SHA-256 不符"))
    return len(files), failures

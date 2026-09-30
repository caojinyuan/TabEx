"""后台线程：Git 状态、目录大小检查与诊断包导出。"""

import os
import subprocess
import time

from PyQt5.QtCore import pyqtSignal, QMutex, QThread

from .i18n import tr
from .constants import FOLDER_CHECK_TIMEOUT, LARGE_FOLDER_THRESHOLD
from .debuglog import debug_print


class OpenPathWorker(QThread):
    completed = pyqtSignal(str, str)

    def __init__(self, path, parent=None):
        super().__init__(parent)
        self.path = path

    def run(self):
        try:
            path = self.path
            if not path.startswith(('shell:', '::')):
                path = os.path.abspath(path)
                if os.path.isfile(path):
                    path = os.path.dirname(path)
                if not os.path.isdir(path):
                    raise FileNotFoundError(path)
            if not self.isInterruptionRequested():
                self.completed.emit(path, '')
        except OSError as error:
            if not self.isInterruptionRequested():
                self.completed.emit('', str(error))


class QuickFindWorker(QThread):
    completed = pyqtSignal(str, object, str)

    def __init__(self, directory, keyword, limit=200, parent=None):
        super().__init__(parent)
        self.directory = directory
        self.keyword = keyword.casefold()
        self.limit = limit

    def run(self):
        matches, error = [], ''
        try:
            with os.scandir(self.directory) as entries:
                for entry in entries:
                    if self.isInterruptionRequested():
                        return
                    if self.keyword in entry.name.casefold():
                        matches.append(entry.path)
                        if len(matches) >= self.limit:
                            break
            matches.sort(key=lambda path: os.path.basename(path).casefold())
        except OSError as exception:
            error = str(exception)
        if not self.isInterruptionRequested():
            self.completed.emit(self.directory, matches, error)


class DiagnosticExportWorker(QThread):
    def __init__(self, diagnostics, destination, debug_path):
        super().__init__()
        self.diagnostics = diagnostics
        self.destination = destination
        self.debug_path = debug_path
        self.error = ''

    def run(self):
        try:
            self.diagnostics.export(self.destination, self.debug_path)
        except Exception as error:
            self.error = f'{type(error).__name__}: {error}'


class GitStatusWorker(QThread):
    """后台线程获取 Git 状态，避免阻塞 UI"""
    completed = pyqtSignal(str, str, object)  # dir_path, repo_root, summary_or_None

    def __init__(self, dir_path, repo_root, git_exe, parent=None):
        super().__init__(parent)
        self.dir_path = dir_path
        self.repo_root = repo_root
        self.git_exe = git_exe
        self.branch = ''

    def run(self):
        try:
            result = subprocess.run(
                [self.git_exe, 'status', '--porcelain=v1', '-b', '--untracked-files=normal'],
                cwd=self.repo_root,
                capture_output=True, text=True, timeout=3,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000),
            )
            if result.returncode != 0:
                self.completed.emit(self.dir_path, self.repo_root, None)
                return
            lines = result.stdout.strip().splitlines()
            branch = ''
            staged = 0
            modified = 0
            untracked = 0
            for line in lines:
                if line.startswith('## '):
                    branch_info = line[3:]
                    branch = branch_info.split('...')[0].split()[0] if branch_info else ''
                    self.branch = branch
                    continue
                if len(line) < 2:
                    continue
                x, y = line[0], line[1]
                if x == '?' and y == '?':
                    untracked += 1
                else:
                    if x in ('M', 'A', 'D', 'R', 'C'):
                        staged += 1
                    if y in ('M', 'D'):
                        modified += 1
            is_clean = (staged == 0 and modified == 0 and untracked == 0)
            parts = []
            if branch:
                parts.append(tr("分支: {}").format(branch))
            if is_clean:
                parts.append(tr('<span style="color:#2e7d32;font-weight:bold">✔ 无更改</span>'))
            else:
                def _color(label, value, color_pos, color_zero='#2e7d32'):
                    color = color_pos if value > 0 else color_zero
                    return f'<span style="color:{color}">{label} {value}</span>'
                status_parts = [
                    _color(tr("暂存(Add)"), staged, "#d32f2f"),
                    _color(tr("修改(Commit)"), modified, "#d32f2f"),
                    _color(tr("未跟踪(待Add)"), untracked, "#f9a825"),
                ]
                parts.append('  '.join(status_parts))
            summary = ' | '.join(parts) if parts else None
            self.completed.emit(self.dir_path, self.repo_root, summary)
        except Exception:
            self.completed.emit(self.dir_path, self.repo_root, None)


class FolderSizeChecker(QThread):
    """后台线程检查文件夹大小，避免阻塞UI"""
    completed = pyqtSignal(str, int, bool)  # path, file_count, is_large
    
    def __init__(self, path, parent=None):
        super().__init__(parent)
        self.path = path
        self.should_stop = False
        self._mutex = QMutex()
    
    def run(self):
        """在后台线程中计算文件夹大小"""
        if not os.path.exists(self.path) or not os.path.isdir(self.path):
            self.completed.emit(self.path, 0, False)
            return
        
        try:
            file_count = 0
            start_time = time.time()
            
            # 快速计数，不递归子文件夹
            for entry in os.scandir(self.path):
                if self.should_stop:
                    debug_print(f"[FolderSizeChecker] Stopped checking {self.path}")
                    return
                
                file_count += 1
                
                # 超过阈值或超时则提前返回
                if file_count > LARGE_FOLDER_THRESHOLD:
                    debug_print(f"[FolderSizeChecker] Large folder detected: {self.path} (>{LARGE_FOLDER_THRESHOLD} files)")
                    self.completed.emit(self.path, file_count, True)
                    return
                
                # 超时检查
                if (time.time() - start_time) * 1000 > FOLDER_CHECK_TIMEOUT:
                    debug_print(f"[FolderSizeChecker] Timeout checking {self.path}")
                    self.completed.emit(self.path, file_count, file_count > LARGE_FOLDER_THRESHOLD)
                    return
            
            is_large = file_count > LARGE_FOLDER_THRESHOLD
            debug_print(f"[FolderSizeChecker] Folder {self.path}: {file_count} files (large={is_large})")
            self.completed.emit(self.path, file_count, is_large)
            
        except Exception as e:
            debug_print(f"[FolderSizeChecker] Error checking {self.path}: {e}")
            self.completed.emit(self.path, 0, False)
    
    def stop(self):
        """停止检查"""
        self._mutex.lock()
        self.should_stop = True
        self._mutex.unlock()

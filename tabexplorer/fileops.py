"""后台文件任务、同名冲突处理与目录比较。"""

import os
import queue
import shutil
import subprocess
import threading
import time

from PyQt5.QtCore import pyqtSignal, Qt, QThread, QTimer
from PyQt5.QtWidgets import (
    QDialog, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout,
)

from .i18n import tr
from .debuglog import _diagnostic_event
from .system import format_file_size
from .widgets import show_toast
from .search import _search_cache, _SearchCancelled, _SearchResultQueue


class FileBatchOpWorker(QThread):
    """后台执行批量复制/删除，避免系统 Shell 弹框阻塞 UI。"""
    completed = pyqtSignal(str, int, int, list)
    progress = pyqtSignal(str, int, int, str)  # op_type, done_count, total_count, current_name

    def __init__(self, op_type, src_paths, dst_dir=None, parent=None, max_workers=0):
        super().__init__(parent)
        self.op_type = str(op_type or '').lower()
        self.src_paths = [p for p in (src_paths or []) if isinstance(p, str) and p]
        self.dst_dir = dst_dir
        self.max_workers = int(max_workers or 0)
        self._cancel_requested = False
        self.cancelled = False
        self.ok_count = 0
        self.fail_count = 0
        self.done_units = 0
        self.total_units = 0
        self.started_at = 0.0
        self.bytes_done = 0
        self.bytes_total = 0
        self._progress_lock = threading.Lock()
        self._last_progress_at = 0.0
        self.failed_paths = []
        self.rename_plan = {}
        self.conflict_actions = {}  # 源路径 -> 'skip' | 'rename'；未列出的冲突项按 'rename' 处理
        self.skipped_count = 0

    def request_cancel(self):
        self._cancel_requested = True

    def _raise_if_cancelled(self):
        if self._cancel_requested:
            self.cancelled = True
            raise RuntimeError("FILE_OP_CANCELLED")

    def _emit_progress(self, current_name):
        now = time.monotonic()
        with self._progress_lock:
            if now - self._last_progress_at < 0.1:
                return
            self._last_progress_at = now
        self.progress.emit(self.op_type, self.done_units, max(1, self.total_units), current_name or "")

    def _get_io_workers(self, task_count):
        if not self.max_workers and str(self.dst_dir or '').startswith('\\\\'):
            return 1
        if task_count is None:
            return max(1, min(self.max_workers or 2, 16))
        if task_count <= 1:
            return 1
        if self.max_workers > 0:
            return max(1, min(self.max_workers, task_count))
        return min(2, task_count)

    @staticmethod
    def _is_reparse_path(path):
        return os.path.islink(path) or bool(getattr(os.lstat(path), 'st_file_attributes', 0) & 0x400)

    @staticmethod
    def _make_unique_path(target_path):
        """避免覆盖：若目标已存在，自动追加 " - copy" 后缀。"""
        if not os.path.exists(target_path):
            return target_path
        base, ext = os.path.splitext(target_path)
        index = 1
        while True:
            suffix = " - copy" if index == 1 else f" - copy{index}"
            candidate = f"{base}{suffix}{ext}"
            if not os.path.exists(candidate):
                return candidate
            index += 1

    @staticmethod
    def _clear_readonly(path):
        """Windows 下清除只读位，避免 rmtree 删除 .git 等只读文件失败。"""
        try:
            if os.name != 'nt' or not path:
                return

            # 优先纯 Python 方式改权限，避免在 windowed exe 中频繁拉起控制台进程。
            import stat
            try:
                mode = os.stat(path).st_mode
                os.chmod(path, mode | stat.S_IWRITE)
                return
            except Exception:
                pass

            # 兜底：仅在 chmod 失败时调用 attrib，并强制无窗口。
            try:
                subprocess.run(
                    ["attrib", "-R", path],
                    check=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000),
                )
            except Exception:
                pass
        except Exception:
            pass

    @classmethod
    def _retry_remove_once_cleared(cls, func, target, retries=6):
        """清只读后重试删除，覆盖 Windows 上短暂占用/权限刷新延迟场景。"""
        # 先直接删，常见场景下可避免无意义的权限修复与额外系统调用。
        try:
            func(target)
            return
        except Exception as ex:
            last_error = ex

        for attempt in range(retries):
            cls._clear_readonly(target)
            try:
                func(target)
                return
            except Exception as ex:
                last_error = ex
                time.sleep(0.05 * (attempt + 1))
        if last_error:
            raise last_error

    @classmethod
    def _rmtree_robust(cls, path):
        """递归删除目录，遇到只读文件自动修复后重试。"""
        import shutil

        def _onerror(func, target, exc_info):
            cls._retry_remove_once_cleared(func, target)

        shutil.rmtree(path, onerror=_onerror)

    @staticmethod
    def _estimate_path_units(path):
        """估算工作量：文件/目录总条目数（用于更细粒度进度显示）。"""
        try:
            if os.path.isfile(path):
                return 1
            if not os.path.isdir(path):
                return 1
            total = 1  # 根目录本身
            for root, dirs, files in os.walk(path):
                total += len(dirs) + len(files)
            return max(1, total)
        except Exception:
            return 1

    def _estimate_total_units(self):
        total = 0
        for p in self.src_paths:
            total += self._estimate_path_units(p)
        return max(1, total)

    @staticmethod
    def _collect_copy_tasks(src_dir, dst_dir):
        """收集目录复制任务：目录创建顺序执行，文件复制可并发。"""
        dirs_to_create = [dst_dir]
        file_tasks = []
        for root, dirs, files in os.walk(src_dir):
            rel = os.path.relpath(root, src_dir)
            dst_root = dst_dir if rel == '.' else os.path.join(dst_dir, rel)
            for dname in dirs:
                dirs_to_create.append(os.path.join(dst_root, dname))
            for fname in files:
                src_file = os.path.join(root, fname)
                dst_file = os.path.join(dst_root, fname)
                file_tasks.append((src_file, dst_file, fname))
        return dirs_to_create, file_tasks

    @classmethod
    def _collect_delete_tasks(cls, src_dir):
        """收集目录删除任务：文件可并发删除，目录按深度逆序删除。"""
        file_tasks = []
        dirs_to_remove = []
        def fail_walk(error):
            raise error

        for root, dirs, files in os.walk(src_dir, topdown=True, followlinks=False, onerror=fail_walk):
            for name in dirs + files:
                path = os.path.join(root, name)
                if cls._is_reparse_path(path):
                    raise ValueError(tr("永久删除拒绝包含链接的目录: ") + path)
            for fname in files:
                fpath = os.path.join(root, fname)
                file_tasks.append((fpath, fname))
            for dname in dirs:
                dpath = os.path.join(root, dname)
                dirs_to_remove.append((dpath, dname))
        dirs_to_remove.sort(key=lambda entry: entry[0].count(os.sep), reverse=True)
        dirs_to_remove.append((src_dir, os.path.basename(src_dir.rstrip('\\/')) or src_dir))
        return file_tasks, dirs_to_remove

    def _copy_file_task(self, src_file, dst_file):
        self._raise_if_cancelled()
        import tempfile
        parent = os.path.dirname(dst_file)
        if parent:
            os.makedirs(parent, exist_ok=True)
        descriptor, temporary = tempfile.mkstemp(prefix='.tabex-copy-', dir=parent or '.')
        try:
            with os.fdopen(descriptor, 'wb') as target, open(src_file, 'rb') as source:
                while True:
                    self._raise_if_cancelled()
                    block = source.read(1024 * 1024)
                    if not block:
                        break
                    target.write(block)
                    with self._progress_lock:
                        self.bytes_done += len(block)
                    self._emit_progress(os.path.basename(src_file))
            self._raise_if_cancelled()
            shutil.copystat(src_file, temporary)
            if os.path.lexists(dst_file):
                raise FileExistsError(dst_file)
            os.rename(temporary, dst_file)
        finally:
            if os.path.exists(temporary):
                os.remove(temporary)

    def _delete_file_task(self, src_file):
        self._raise_if_cancelled()
        self._retry_remove_once_cleared(os.remove, src_file)

    def _run_parallel_file_tasks(self, tasks, task_runner, task_name_getter):
        """并发执行文件级任务，按完成顺序回传逐项进度。"""
        if not tasks:
            return []

        errors = []
        max_workers = self._get_io_workers(len(tasks) if hasattr(tasks, '__len__') else None)
        if max_workers <= 1:
            for task in tasks:
                self._raise_if_cancelled()
                name = task_name_getter(task)
                try:
                    task_runner(task)
                    self.done_units += 1
                    self._emit_progress(name)
                except RuntimeError as e:
                    if str(e) == "FILE_OP_CANCELLED":
                        self.cancelled = True
                        break
                    errors.append(f"{name}: {e}")
                    self.done_units += 1
                    self._emit_progress(name)
                except Exception as e:
                    errors.append(f"{name}: {e}")
                    self.done_units += 1
                    self._emit_progress(name)
            return errors

        from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
        future_to_task = {}
        with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix='file-op') as executor:
            task_iter = iter(tasks)
            exhausted = False
            while future_to_task or not exhausted:
                if self._cancel_requested:
                    self.cancelled = True
                    for pending in future_to_task:
                        pending.cancel()
                    break
                while not exhausted and len(future_to_task) < max_workers * 2:
                    task = next(task_iter, None)
                    if task is None:
                        exhausted = True
                        break
                    future_to_task[executor.submit(task_runner, task)] = task
                if not future_to_task:
                    continue
                ready, _pending = wait(future_to_task, timeout=0.1, return_when=FIRST_COMPLETED)
                for future in ready:
                    task = future_to_task.pop(future)
                    name = task_name_getter(task)
                    try:
                        future.result()
                    except RuntimeError as error:
                        if str(error) == 'FILE_OP_CANCELLED':
                            self.cancelled = True
                        else:
                            errors.append(f"{name}: {error}")
                    except Exception as error:
                        errors.append(f"{name}: {error}")
                    self.done_units += 1
                    self._emit_progress(name)

        return errors

    def _copy_dir_cancelable(self, src_dir, dst_dir):
        def fail_walk(error):
            raise error

        def tasks():
            for root, dirs, files in os.walk(src_dir, followlinks=False, onerror=fail_walk):
                self._raise_if_cancelled()
                for dirname in dirs:
                    candidate = os.path.join(root, dirname)
                    if self._is_reparse_path(candidate):
                        raise ValueError(tr("暂不支持复制目录链接: ") + candidate)
                relative = os.path.relpath(root, src_dir)
                destination = dst_dir if relative == '.' else os.path.join(dst_dir, relative)
                os.makedirs(destination, exist_ok=True)
                self.total_units += len(files) + len(dirs)
                self.done_units += 1
                for filename in files:
                    self._raise_if_cancelled()
                    source = os.path.join(root, filename)
                    self.bytes_total += os.path.getsize(source)
                    yield source, os.path.join(destination, filename), filename

        return self._run_parallel_file_tasks(
            tasks(),
            task_runner=lambda t: self._copy_file_task(t[0], t[1]),
            task_name_getter=lambda t: t[2]
        )

    def _delete_dir_cancelable(self, src_dir):
        file_tasks, dirs_to_remove = self._collect_delete_tasks(src_dir)
        delete_errors = self._run_parallel_file_tasks(
            file_tasks,
            task_runner=lambda t: self._delete_file_task(t[0]),
            task_name_getter=lambda t: t[1]
        )

        for dpath, dname in dirs_to_remove:
            self._raise_if_cancelled()
            try:
                self._retry_remove_once_cleared(os.rmdir, dpath)
            except Exception as e:
                delete_errors.append(f"{dpath}: {e}")
            self.done_units += 1
            self._emit_progress(dname)

        return delete_errors

    def run(self):
        import shutil

        self.started_at = time.monotonic()
        self.ok_count = 0
        self.fail_count = 0
        errors = []
        self.total_units = max(1, len(self.src_paths))
        self.done_units = 0

        self._emit_progress("")

        try:
            if self.op_type == 'copy':
                if not self.dst_dir or not os.path.isdir(self.dst_dir):
                    self.failed_paths = list(self.src_paths)
                    self.completed.emit(self.op_type, 0, len(self.src_paths), [tr("复制失败：目标目录无效")])
                    return

                for src in self.src_paths:
                    try:
                        self._raise_if_cancelled()
                        src_norm = os.path.normpath(src)
                        name = os.path.basename(src_norm) or os.path.basename(os.path.dirname(src_norm)) or 'item'
                        target = os.path.join(self.dst_dir, name)
                        if self.conflict_actions.get(src) == 'skip' and os.path.lexists(target):
                            self.skipped_count += 1
                            self.done_units += 1
                            self._emit_progress(name)
                            continue
                        dst_path = self._make_unique_path(target)
                        if os.path.isdir(src_norm):
                            if self._is_reparse_path(src_norm):
                                raise ValueError(tr("暂不支持复制目录链接: ") + src_norm)
                            source_real = os.path.normcase(os.path.realpath(src_norm))
                            target_real = os.path.normcase(os.path.realpath(dst_path))
                            try:
                                nested = os.path.commonpath([source_real, target_real]) == source_real
                            except ValueError:
                                nested = False
                            if nested:
                                raise ValueError(tr("不能将目录复制到自身或其子目录"))
                            dir_errors = self._copy_dir_cancelable(src_norm, dst_path)
                            self._raise_if_cancelled()
                            if dir_errors:
                                raise RuntimeError("; ".join(dir_errors[:5]))
                        else:
                            self.bytes_total += os.path.getsize(src_norm)
                            self._copy_file_task(src_norm, dst_path)
                            self.done_units += 1
                            self._emit_progress(name)
                        self.ok_count += 1
                    except RuntimeError as e:
                        if str(e) == "FILE_OP_CANCELLED":
                            break
                        self.fail_count += 1
                        self.failed_paths.append(src)
                        errors.append(f"{src}: {e}")
                    except Exception as e:
                        self.fail_count += 1
                        self.failed_paths.append(src)
                        errors.append(f"{src}: {e}")

            elif self.op_type in ('delete', 'permanent_delete'):
                for src in self.src_paths:
                    try:
                        self._raise_if_cancelled()
                        src_norm = os.path.normpath(src)
                        name = os.path.basename(src_norm.rstrip('\\/')) or src_norm
                        if self.op_type == 'delete':
                            from send2trash import send2trash
                            send2trash(os.path.abspath(src_norm))
                            self.done_units += 1
                            self._emit_progress(name)
                        elif os.path.isdir(src_norm):
                            if self._is_reparse_path(src_norm):
                                raise ValueError(tr("永久删除拒绝目录链接: ") + src_norm)
                            dir_errors = self._delete_dir_cancelable(src_norm)
                            self._raise_if_cancelled()
                            if dir_errors:
                                raise RuntimeError("; ".join(dir_errors[:5]))
                        else:
                            self._retry_remove_once_cleared(os.remove, src_norm)
                            self.done_units += 1
                            self._emit_progress(name)
                        self.ok_count += 1
                    except RuntimeError as e:
                        if str(e) == "FILE_OP_CANCELLED":
                            break
                        self.fail_count += 1
                        self.failed_paths.append(src)
                        errors.append(f"{src}: {e}")
                    except Exception as e:
                        self.fail_count += 1
                        self.failed_paths.append(src)
                        errors.append(f"{src}: {e}")
            elif self.op_type == 'rename':
                destinations = set()
                for source in self.src_paths:
                    self._raise_if_cancelled()
                    target = self.rename_plan[source]
                    normalized = os.path.normcase(os.path.abspath(target))
                    if normalized in destinations or os.path.lexists(target):
                        raise FileExistsError(target)
                    if not os.path.lexists(source):
                        raise FileNotFoundError(source)
                    destinations.add(normalized)
                for source in self.src_paths:
                    self._raise_if_cancelled()
                    try:
                        os.rename(source, self.rename_plan[source])
                        self.ok_count += 1
                    except OSError as error:
                        self.fail_count += 1
                        self.failed_paths.append(source)
                        errors.append(f"{source}: {error}")
                    self.done_units += 1
                    self._emit_progress(os.path.basename(source))
            else:
                self.fail_count = len(self.src_paths)
                errors.append(tr("不支持的操作类型"))
        except Exception as e:
            if str(e) == 'FILE_OP_CANCELLED':
                self.cancelled = True
            else:
                self.fail_count = max(self.fail_count, len(self.src_paths) - self.ok_count)
                self.failed_paths = list(self.src_paths) if not self.ok_count else self.failed_paths
                errors.append(str(e))

        if self.cancelled and self.done_units < self.total_units:
            errors.append(tr("操作已取消"))

        self.completed.emit(self.op_type, self.ok_count, self.fail_count, errors)


class FileTaskPanel(QDialog):
    tasks_changed = pyqtSignal()

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(tr("文件任务"))
        self.resize(960, 400)
        self.records = []
        layout = QVBoxLayout(self)
        self.table = QTreeWidget(self)
        self.table.setHeaderLabels([tr("操作"), tr("目标 / 来源"), tr("状态"), tr("进度")])
        self.table.setRootIsDecorated(False)
        self.table.setSelectionMode(QTreeWidget.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setUniformRowHeights(True)
        self.table.setTextElideMode(Qt.ElideMiddle)
        from PyQt5.QtWidgets import QHeaderView, QStyle
        self.table.header().setStretchLastSection(False)
        self.table.header().setSectionResizeMode(1, QHeaderView.Stretch)
        for column in (0, 2, 3):
            self.table.header().setSectionResizeMode(column, QHeaderView.ResizeToContents)
        self.table.setStyleSheet(f'QTreeView::item {{ height: {max(28, self.fontMetrics().height() + 10)}px; }}')
        layout.addWidget(self.table)
        controls = QHBoxLayout()
        for name, label, icon, callback in (
                ('cancel_button', '取消任务', QStyle.SP_BrowserStop, self.cancel_selected),
                ('retry_button', '重试失败项', QStyle.SP_BrowserReload, self.retry_selected),
                ('errors_button', '查看错误', QStyle.SP_MessageBoxWarning, self.show_errors),
                ('open_button', '打开位置', QStyle.SP_DirOpenIcon, self.open_selected_location),
                ('clear_button', '清除已完成', QStyle.SP_DialogResetButton, self.clear_completed)):
            button = QPushButton(tr(label), self)
            button.setIcon(self.style().standardIcon(icon))
            button.setToolTip(tr(label))
            button.setMinimumHeight(28)
            button.clicked.connect(callback)
            setattr(self, name, button)
            controls.addWidget(button)
        layout.addLayout(controls)
        self.timer = QTimer(self)
        self.timer.setInterval(250)
        self.timer.timeout.connect(self.refresh_progress)
        self.table.currentItemChanged.connect(self._update_controls)
        self._update_controls()

    def _update_controls(self, *args):
        record = self.selected_record()
        worker = record['worker'] if record else None
        self.cancel_button.setEnabled(bool(worker and not record.get('completed')
                                           and not worker._cancel_requested))
        self.retry_button.setEnabled(bool(record and record['done'] and record['failed']))
        self.errors_button.setEnabled(bool(record and record['errors']))
        self.open_button.setEnabled(bool(record and (record['destination'] or record['paths'])))
        self.clear_button.setEnabled(any(entry['done'] for entry in self.records))

    def open_selected_location(self):
        record = self.selected_record()
        if not record:
            return
        path = record['destination'] or (os.path.dirname(record['paths'][0]) if record['paths'] else '')
        owner = self.parent()
        if path and owner is not None and hasattr(owner, 'add_new_tab'):
            owner.add_new_tab(path)

    def task_counts(self):
        running = sum(not record['done'] for record in self.records)
        failed = sum(bool(record['failed'] or record['errors']) and not record.get('cancelled', False)
                     for record in self.records)
        return running, failed

    def start_task(self, op_type, paths, destination=None, max_workers=0, tab=None, rename_plan=None,
                   conflict_actions=None):
        worker = FileBatchOpWorker(op_type, paths, destination, self, max_workers=max_workers)
        worker.rename_plan = dict(rename_plan or {})
        worker.conflict_actions = dict(conflict_actions or {})
        operation_name = {'copy': '复制', 'delete': '回收站', 'permanent_delete': '永久删除', 'rename': '重命名'}
        location = destination or (paths[0] if paths else '')
        row = QTreeWidgetItem([tr(operation_name.get(op_type, op_type)), location, tr("运行中"), ''])
        row.setToolTip(1, '\n'.join(paths) + ('\n' + destination if destination else ''))
        record = {'worker': worker, 'row': row, 'op': op_type, 'paths': list(paths),
                  'destination': destination, 'max_workers': max_workers, 'errors': [],
              'failed': [], 'done': False, 'completed': False, 'cancelled': False,
              'rename_plan': worker.rename_plan, 'conflict_actions': worker.conflict_actions}
        row.setData(0, Qt.UserRole, len(self.records))
        self.records.append(record)
        self.table.addTopLevelItem(row)
        self.table.setCurrentItem(row)
        _diagnostic_event('file_task_start', f'operation={op_type} items={len(paths)}')
        worker.completed.connect(self.task_completed)
        worker.finished.connect(self.thread_finished)
        if tab is not None:
            worker.progress.connect(tab._on_file_batch_op_progress)
            worker.completed.connect(tab._on_file_batch_op_finished)
            tab._file_op_worker = worker
        self.timer.start()
        worker.start()
        self._update_controls()
        self.tasks_changed.emit()
        return worker

    def task_completed(self, op_type, ok_count, fail_count, errors):
        from PyQt5.QtWidgets import QStyle
        _search_cache.clear()
        _diagnostic_event('file_task_complete', f'operation={op_type} ok={ok_count} failed={fail_count}')
        worker = self.sender()
        record = next((entry for entry in self.records if entry['worker'] is worker), None)
        if record is None:
            return
        record['errors'] = list(errors)
        record['failed'] = list(worker.failed_paths)
        record['completed'] = True
        record['cancelled'] = worker.cancelled
        if worker.cancelled:
            state, icon = tr('已取消'), QStyle.SP_BrowserStop
        elif fail_count or errors:
            state = tr('部分失败') if ok_count else tr('失败')
            icon = QStyle.SP_MessageBoxWarning
        else:
            state, icon = tr('全部成功'), QStyle.SP_DialogApplyButton
        summary = tr("{}：成功 {} 项，失败 {} 项").format(state, ok_count, fail_count)
        if getattr(worker, 'skipped_count', 0):
            summary += tr("，跳过 {} 项").format(worker.skipped_count)
        record['summary'] = summary
        record['row'].setText(2, summary)
        record['row'].setIcon(2, self.style().standardIcon(icon))
        record['row'].setText(3, format_file_size(worker.bytes_done) if op_type == 'copy' else str(worker.done_units))
        record['row'].setToolTip(2, summary + ('\n' + '\n'.join(errors) if errors else ''))
        self._update_controls()
        self.tasks_changed.emit()

    def thread_finished(self):
        worker = self.sender()
        finished_record = None
        for record in self.records:
            if record['worker'] is worker:
                record['done'] = True
                record['worker'] = None
                finished_record = record
                break
        worker.deleteLater()
        if not self.has_running_tasks():
            self.timer.stop()
        self._update_controls()
        self.tasks_changed.emit()
        if finished_record is not None:
            self._notify_task_result(finished_record)

    def _notify_task_result(self, record):
        owner = self.parent()
        if owner is None or self.isVisible():
            return

        def show_record():
            if record in self.records:
                self.table.setCurrentItem(record['row'])
                self.show()
                self.raise_()
                self.activateWindow()

        if record['errors'] or record['cancelled']:
            show_toast(owner, tr('文件任务'), record.get('summary', ''), level='warning',
                       action_text=tr('查看任务'), action=show_record)
        else:
            path = record['destination'] or (os.path.dirname(record['paths'][0]) if record['paths'] else '')
            action = (lambda: owner.add_new_tab(path)) if path and hasattr(owner, 'add_new_tab') else show_record
            show_toast(owner, tr('文件任务'), record.get('summary', ''), level='success',
                       action_text=tr('打开位置') if path and hasattr(owner, 'add_new_tab') else tr('查看任务'),
                       action=action)

    def has_running_tasks(self):
        return any(not entry['done'] for entry in self.records)

    def refresh_progress(self):
        for record in self.records:
            worker = record['worker']
            if worker is None or record['done']:
                continue
            if worker.op_type == 'copy':
                elapsed = max(0.1, time.monotonic() - worker.started_at)
                record['row'].setText(3, f"{format_file_size(worker.bytes_done)} / {format_file_size(worker.bytes_total)}  "
                                       f"{format_file_size(worker.bytes_done / elapsed)}/s")
            else:
                record['row'].setText(3, f"{worker.done_units} / {worker.total_units}")

    def selected_record(self):
        row = self.table.currentItem()
        return next((entry for entry in self.records if entry['row'] is row), None)

    def cancel_selected(self):
        record = self.selected_record()
        if record and record['worker'] and not record.get('completed'):
            record['worker'].request_cancel()
            record['row'].setText(2, tr("取消中"))
            self._update_controls()

    def retry_selected(self):
        record = self.selected_record()
        if not record or not record['done'] or not record['failed']:
            return
        from PyQt5.QtWidgets import QMessageBox
        if QMessageBox.question(self, tr("重试失败项"), '\n'.join(record['failed']),
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        self.start_task(record['op'], record['failed'], record['destination'], record['max_workers'],
                rename_plan=record['rename_plan'], conflict_actions=record.get('conflict_actions'))

    def show_errors(self):
        record = self.selected_record()
        if record:
            from PyQt5.QtWidgets import QMessageBox
            box = QMessageBox(self)
            box.setWindowTitle(tr("任务详情"))
            box.setText(record['row'].text(2))
            box.setDetailedText('\n'.join(record['errors']) or tr("无错误"))
            box.exec_()

    def clear_completed(self):
        for record in list(self.records):
            if record['done']:
                self.table.takeTopLevelItem(self.table.indexOfTopLevelItem(record['row']))
                self.records.remove(record)
        self._update_controls()
        self.tasks_changed.emit()


def _find_copy_conflicts(paths, dst_dir):
    """返回目标目录已有同名项的来源路径；复制到来源所在目录时按系统习惯自动重命名，不算冲突。"""
    dst_key = os.path.normcase(os.path.normpath(dst_dir))
    conflicts = []
    for source in paths:
        normalized = os.path.normpath(source)
        if os.path.normcase(os.path.dirname(normalized)) == dst_key:
            continue
        name = os.path.basename(normalized)
        if name and os.path.lexists(os.path.join(dst_dir, name)):
            conflicts.append(source)
    return conflicts


class CopyConflictDialog(QDialog):
    """后台粘贴前逐项选择同名冲突的处理方式。"""

    def __init__(self, conflicts, dst_dir, parent=None):
        super().__init__(parent)
        from PyQt5.QtWidgets import QComboBox, QDialogButtonBox
        self.setWindowTitle(tr("同名项目"))
        self.resize(640, 360)
        layout = QVBoxLayout(self)
        label = QLabel(tr("目标目录中已存在 {} 个同名项目：{}").format(len(conflicts), dst_dir), self)
        label.setWordWrap(True)
        layout.addWidget(label)
        self.table = QTreeWidget(self)
        self.table.setHeaderLabels([tr("名称"), tr("处理方式")])
        self.table.setRootIsDecorated(False)
        self.table.header().setStretchLastSection(False)
        self.table.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._combos = {}
        for source in conflicts:
            item = QTreeWidgetItem([os.path.basename(os.path.normpath(source)), ''])
            item.setToolTip(0, source)
            self.table.addTopLevelItem(item)
            combo = QComboBox(self.table)
            combo.addItem(tr("保留两者（自动重命名）"), 'rename')
            combo.addItem(tr("跳过"), 'skip')
            self.table.setItemWidget(item, 1, combo)
            self._combos[source] = combo
        layout.addWidget(self.table)
        row = QHBoxLayout()
        for text, value in ((tr("全部保留两者"), 'rename'), (tr("全部跳过"), 'skip')):
            button = QPushButton(text, self)
            button.clicked.connect(lambda _checked=False, choice=value: self.set_all(choice))
            row.addWidget(button)
        row.addStretch(1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        row.addWidget(buttons)
        layout.addLayout(row)

    def set_all(self, value):
        for combo in self._combos.values():
            combo.setCurrentIndex(combo.findData(value))

    def actions(self):
        return {source: combo.currentData() for source, combo in self._combos.items()}


def _plan_batch_rename(paths, find_text, replacement):
    import ntpath
    if not find_text:
        raise ValueError(tr("查找文本不能为空"))
    plan = {}
    targets = set()
    sources = {ntpath.normcase(ntpath.abspath(path)) for path in paths}
    reserved = {'CON', 'PRN', 'AUX', 'NUL'} | {f'{prefix}{number}' for prefix in ('COM', 'LPT') for number in range(1, 10)}
    for source in paths:
        old_name = ntpath.basename(source)
        new_name = old_name.replace(find_text, replacement)
        if old_name == new_name:
            continue
        if (not new_name or new_name in ('.', '..') or new_name.endswith((' ', '.'))
                or any(character in '<>:"/\\|?*' or ord(character) < 32 for character in new_name)
                or new_name.split('.')[0].upper() in reserved):
            raise ValueError(tr("非法文件名: ") + new_name)
        target = ntpath.join(ntpath.dirname(source), new_name)
        normalized = ntpath.normcase(ntpath.abspath(target))
        if normalized in targets or normalized in sources:
            raise ValueError(tr("重命名目标冲突: ") + target)
        targets.add(normalized)
        plan[source] = target
    return plan


def _compare_directories(left_root, right_root, cancel_event, emit):
    import filecmp

    def scan(root):
        if not os.path.isdir(root):
            raise ValueError(tr("目录不存在: ") + root)
        entries = {}

        def onerror(error):
            raise error

        for directory, directories, filenames in os.walk(root, followlinks=False, onerror=onerror):
            if cancel_event.is_set():
                return entries
            for name in directories + filenames:
                path = os.path.join(directory, name)
                relative = os.path.relpath(path, root)
                key = os.path.normcase(relative)
                is_link = os.path.islink(path) or bool(getattr(os.lstat(path), 'st_file_attributes', 0) & 0x400)
                kind = 'link' if is_link else ('directory' if name in directories else 'file')
                entries[key] = (relative, kind, path)
                if len(entries) > 100000:
                    raise ValueError(tr("目录比较超过 100000 项上限，请选择更小的目录"))
            directories[:] = [name for name in directories
                              if entries[os.path.normcase(os.path.relpath(os.path.join(directory, name), root))][1] != 'link']
        return entries

    left_entries = scan(left_root)
    if cancel_event.is_set():
        return
    right_entries = scan(right_root)
    for key in sorted(left_entries.keys() | right_entries.keys()):
        if cancel_event.is_set():
            return
        left = left_entries.get(key)
        right = right_entries.get(key)
        relative = (left or right)[0]
        if left is None:
            state = tr("仅右侧")
        elif right is None:
            state = tr("仅左侧")
        elif left[1] != right[1]:
            state = tr("类型不同")
        elif left[1] == 'link':
            state = tr("链接未比较")
        elif left[1] == 'directory':
            continue
        else:
            try:
                before_left, before_right = os.stat(left[2]), os.stat(right[2])
                same = filecmp.cmp(left[2], right[2], shallow=False)
                after_left, after_right = os.stat(left[2]), os.stat(right[2])
                signatures = lambda info: (info.st_size, info.st_mtime_ns)
                if signatures(before_left) != signatures(after_left) or signatures(before_right) != signatures(after_right):
                    state = tr("比较期间文件发生变化")
                elif same:
                    continue
                else:
                    state = tr("内容不同")
            except OSError as error:
                state = tr("读取失败: ") + str(error)
        emit((relative, state, left[2] if left else '', right[2] if right else ''))


class DirectoryCompareDialog(QDialog):
    def __init__(self, parent, left_path, right_path):
        super().__init__(parent)
        self.setWindowTitle(tr("目录差异比较"))
        self.resize(900, 500)
        self.cancel_event = threading.Event()
        self.results = None
        self.running = False
        layout = QVBoxLayout(self)
        self.paths = []
        for label, path in (("左侧", left_path), ("右侧", right_path)):
            row = QHBoxLayout()
            row.addWidget(QLabel(tr(label), self))
            edit = QLineEdit(path, self)
            self.paths.append(edit)
            row.addWidget(edit)
            browse = QPushButton(tr("浏览..."), self)
            browse.clicked.connect(lambda checked=False, target=edit: self.browse(target))
            row.addWidget(browse)
            layout.addLayout(row)
        self.table = QTreeWidget(self)
        self.table.setHeaderLabels([tr("相对路径"), tr("差异"), tr("左侧"), tr("右侧")])
        self.table.setRootIsDecorated(False)
        self.table.header().setStretchLastSection(False)
        self.table.header().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.header().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 200)
        self.table.setColumnWidth(1, 140)
        self.table.itemDoubleClicked.connect(self.open_item)
        layout.addWidget(self.table)
        controls = QHBoxLayout()
        self.start_button = QPushButton(tr("比较"), self)
        self.stop_button = QPushButton(tr("停止"), self)
        self.stop_button.setEnabled(False)
        self.status = QLabel('', self)
        self.start_button.clicked.connect(self.start_compare)
        self.stop_button.clicked.connect(self.stop_compare)
        controls.addWidget(self.start_button)
        controls.addWidget(self.stop_button)
        controls.addWidget(self.status, 1)
        layout.addLayout(controls)
        self.timer = QTimer(self)
        self.timer.setInterval(60)
        self.timer.timeout.connect(self.consume_results)

    def browse(self, target):
        from PyQt5.QtWidgets import QFileDialog
        path = QFileDialog.getExistingDirectory(self, tr("选择目录"), target.text())
        if path:
            target.setText(path)

    def start_compare(self):
        if self.running:
            return
        left, right = [edit.text().strip() for edit in self.paths]
        if not left or not right:
            return
        self.cancel_event = threading.Event()
        self.results = _SearchResultQueue(self.cancel_event)
        cancel_event, results = self.cancel_event, self.results
        self.table.clear()
        self.running = True
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.status.setText(tr("比较中"))

        def work():
            try:
                _compare_directories(left, right, cancel_event, lambda row: results.put(('row', row)))
                results.put(('done', ''))
            except _SearchCancelled:
                pass
            except Exception as error:
                try:
                    results.put(('error', str(error)))
                except _SearchCancelled:
                    pass

        threading.Thread(target=work, daemon=True, name='DirectoryCompare').start()
        self.timer.start()

    def consume_results(self):
        for _index in range(100):
            try:
                kind, payload = self.results.get_nowait()
            except queue.Empty:
                break
            if kind == 'row':
                row = QTreeWidgetItem(list(payload))
                for column, value in enumerate(payload):
                    row.setToolTip(column, value)
                self.table.addTopLevelItem(row)
            else:
                self.running = False
                self.timer.stop()
                self.start_button.setEnabled(True)
                self.stop_button.setEnabled(False)
                self.status.setText(payload if kind == 'error' else tr("比较完成，差异项: ") + str(self.table.topLevelItemCount()))
                break

    def stop_compare(self):
        self.cancel_event.set()
        self.running = False
        self.timer.stop()
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.status.setText(tr("已停止"))

    def open_item(self, item, column):
        path = item.text(3 if column == 3 else 2) or item.text(3)
        if path:
            self.parent().add_new_tab(os.path.dirname(path), select_file=os.path.basename(path))

    def closeEvent(self, event):
        self.stop_compare()
        super().closeEvent(event)


def _confirm_file_preview(parent, title, description, before, after):
    import difflib
    from PyQt5.QtWidgets import QPlainTextEdit, QDialogButtonBox
    dialog = QDialog(parent)
    dialog.setWindowTitle(title)
    dialog.resize(900, 550)
    layout = QVBoxLayout(dialog)
    label = QLabel(description, dialog)
    label.setWordWrap(True)
    layout.addWidget(label)
    preview = QPlainTextEdit(dialog)
    preview.setReadOnly(True)
    preview.setLineWrapMode(QPlainTextEdit.NoWrap)
    difference = ''.join(difflib.unified_diff(before.splitlines(keepends=True), after.splitlines(keepends=True),
                                            fromfile='Before', tofile='After'))
    preview.setPlainText(difference or tr("无变化"))
    layout.addWidget(preview)
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=dialog)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    buttons.button(QDialogButtonBox.Ok).setText(tr("确定"))
    buttons.button(QDialogButtonBox.Cancel).setText(tr("取消"))
    buttons.button(QDialogButtonBox.Cancel).setDefault(True)
    layout.addWidget(buttons)
    return dialog.exec_() == QDialog.Accepted


def _atomic_reviewed_write(path, expected, content):
    import tempfile
    parent = os.path.dirname(os.path.abspath(path))
    os.makedirs(parent, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix='.tabex-write-', dir=parent)
    try:
        with os.fdopen(descriptor, 'wb') as target:
            target.write(content)
            target.flush()
            os.fsync(target.fileno())
        if expected is None:
            if os.path.lexists(path):
                raise FileExistsError(path)
            os.rename(temporary, path)
        else:
            if os.path.islink(path):
                raise ValueError(tr("拒绝覆盖文件链接"))
            with open(path, 'rb') as source:
                if source.read() != expected:
                    raise ValueError(tr("文件在预览后已变化，请重新读取"))
            shutil.copymode(path, temporary)
            os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)

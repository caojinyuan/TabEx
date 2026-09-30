"""TabEx 功能、性能约束与模拟日常使用的回归测试。

阅读方法：每个 test_ 开头的方法就是一个测试用例，下面的中文说明描述场景和预期。
assertEqual / assertTrue / assertFalse：检查实际结果是否等于预期、是否为真或为假。
assertRaises：检查危险或无效操作是否按预期报错；这里报错反而表示保护生效。
patch / SimpleNamespace：用可控的模拟对象代替外部服务或界面对象，不代表真实环境测试。
TemporaryDirectory：在临时目录中准备文件，退出后自动清理，不使用用户的工作文件。
setUp：每个用例执行前准备环境；setUpClass：同一组用例共用的环境只准备一次。

运行：python -m unittest discover -s tests -v
覆盖边界：包含真实临时文件和 Qt 窗口测试，但不覆盖真实网络盘、Everything 服务、
系统回收站、完整 Explorer 交互及长时间性能压测；截图生成不等于自动验证视觉正确性。
"""

import ast
from collections import OrderedDict
from contextlib import ExitStack
import hashlib
import os
from pathlib import Path
import queue
import shutil
import tempfile
import threading
import time
import types
import unittest
from unittest.mock import Mock, patch
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QTimer
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem, QPushButton


SOURCE = Path(__file__).resolve().parents[1] / 'TabEx.py'
TREE = ast.parse(SOURCE.read_text(encoding='utf-8-sig'))


def load_definitions(names, **namespace):
    """从主程序提取指定类或函数，注入测试依赖，避免单元测试启动整个应用。"""
    selected = [node for node in TREE.body
                if isinstance(node, (ast.ClassDef, ast.FunctionDef)) and node.name in names]
    exec(compile(ast.Module(body=selected, type_ignores=[]), str(SOURCE), 'exec'), namespace)
    return namespace


class SearchLifecycleTests(unittest.TestCase):
    """第一组：搜索任务的结束、取消、结果隔离和缓存，共 5 个用例。"""

    def setUp(self):
        self.search = types.SimpleNamespace(do_search=lambda task, *args: None)
        self.namespace = load_definitions(
            {'_SearchCancelled', '_SearchResultQueue', '_SearchTask'},
            queue=queue, threading=threading, os=os, tr=lambda value: value,
            SEARCH_RESULT_QUEUE_MAXSIZE=2, SearchDialog=self.search)

    def task(self):
        return self.namespace['_SearchTask'](str(SOURCE.parent), '', 100)

    def test_finished_even_when_engine_returns_early(self):
        """搜索引擎提前返回后，任务仍正确结束。

        场景：用立即返回的模拟搜索引擎执行一次搜索。
        预期：搜索状态变为已结束，结果队列收到 finished 完成通知。
        """
        task = self.task()
        task.run()
        self.assertFalse(task.is_searching)
        self.assertEqual(task.result_queue.get_nowait(), {'type': 'finished'})

    def test_cancelled_task_cannot_write_into_next_task(self):
        """取消旧搜索后，旧结果不能污染新任务。

        场景：创建两个独立任务，取消旧任务后尝试向它写入结果。
        预期：旧任务写入被拒绝，新任务仍处于搜索状态且队列为空。
        """
        previous = self.task()
        current = self.task()
        previous.cancel_event.set()
        with self.assertRaises(self.namespace['_SearchCancelled']):
            previous.add_search_results_batch([{'path': 'old'}])
        self.assertTrue(current.is_searching)
        self.assertTrue(current.result_queue.empty())

    def test_cancel_releases_full_queue_producer(self):
        """队列已满且任务已取消时，写入会退出而不是继续等待。

        场景：填满容量为 2 的队列，先取消任务，再尝试写入第三条消息。
        预期：立即抛出取消异常；本用例不模拟已阻塞的生产线程被唤醒。
        """
        task = self.task()
        task.result_queue.put({})
        task.result_queue.put({})
        task.cancel_event.set()
        with self.assertRaises(self.namespace['_SearchCancelled']):
            task.result_queue.put({})

    def test_invalid_path_reports_error_and_finishes(self):
        """搜索目录无效时，报告错误并发送完成通知。

        场景：模拟 isdir 返回 False，表示目标不是有效目录。
        预期：先收到状态消息，再收到 finished；不访问真实离线目录。
        """
        task = self.task()
        with patch.object(os.path, 'isdir', return_value=False):
            task.run()
        self.assertEqual(task.result_queue.get_nowait()['type'], 'status')
        self.assertEqual(task.result_queue.get_nowait()['type'], 'finished')

    def test_cache_updates_expires_and_separates_engines(self):
        """缓存可更新、会过期，且不同搜索引擎不会共用缓存键。

        场景：同一个键先写旧值再写新值，将时间戳回拨 6 秒模拟过期。
        预期：更新后读到新值，过期后读不到值；Everything 与普通搜索的键不同。
        """
        namespace = load_definitions({'SearchCache'}, OrderedDict=OrderedDict,
                                     threading=threading, time=time, hashlib=hashlib)
        cache = namespace['SearchCache'](max_size=2)
        cache.put('key', ['old'])
        cache.put('key', ['new'])
        self.assertEqual(cache.get('key'), ['new'])
        cache._timestamps['key'] -= 6
        self.assertIsNone(cache.get('key'))
        self.assertNotEqual(cache.get_key('root', 'text', True, False, '', use_everything=True),
                            cache.get_key('root', 'text', True, False, '', use_everything=False))


class FileOperationTests(unittest.TestCase):
    """第二组：复制、删除和重命名的数据安全，共 8 个用例。"""

    def setUp(self):
        namespace = load_definitions(
            {'FileBatchOpWorker'}, QThread=QThread, pyqtSignal=pyqtSignal,
            os=os, threading=threading, time=time, shutil=shutil,
            tr=lambda value: value)
        self.worker_class = namespace['FileBatchOpWorker']

    def test_copy_cancel_removes_partial_file(self):
        """复制中途取消，不留下看似完整的目标文件或临时文件。

        场景：在临时目录创建约 2.1 MB 的文件，第一次进度回调时请求取消。
        预期：抛出取消异常，最终目标不存在，.tabex-copy- 临时文件被清理。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source'
            target = Path(directory) / 'target'
            source.write_bytes(b'content' * 300000)
            worker = self.worker_class('copy', [str(source)], directory)
            worker._emit_progress = lambda name: worker.request_cancel()
            with self.assertRaisesRegex(RuntimeError, 'FILE_OP_CANCELLED'):
                worker._copy_file_task(str(source), str(target))
            self.assertFalse(target.exists())
            self.assertEqual(list(Path(directory).glob('.tabex-copy-*')), [])

    def test_copy_never_overwrites_existing_file(self):
        """复制遇到已有目标文件时，保护目标原内容。

        场景：源文件内容为 new，目标文件已经存在且内容为 old。
        预期：复制报 FileExistsError，目标仍然是 old，不被覆盖。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source'
            target = Path(directory) / 'target'
            source.write_bytes(b'new')
            target.write_bytes(b'old')
            worker = self.worker_class('copy', [str(source)], directory)
            with self.assertRaises(FileExistsError):
                worker._copy_file_task(str(source), str(target))
            self.assertEqual(target.read_bytes(), b'old')

    def test_copy_into_descendant_rejected(self):
        """拒绝把目录复制到自己的子目录，避免递归自复制。

        场景：选择父目录作为源，把其中的 child 子目录作为目标位置。
        预期：任务记录一次失败，child 内没有生成复制内容。
        """
        with tempfile.TemporaryDirectory() as directory:
            child = Path(directory) / 'child'
            child.mkdir()
            worker = self.worker_class('copy', [directory], str(child))
            worker.run()
            self.assertEqual(worker.fail_count, 1)
            self.assertEqual(list(child.iterdir()), [])

    def test_delete_uses_recycle_bin_only(self):
        """普通删除不会绕过回收站接口直接删除文件。

        场景：把 send2trash 模拟成成功返回但不实际移动文件的函数。
        预期：任务记录成功且源文件仍存在；不验证 Windows 回收站的真实移动行为。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source'
            source.write_bytes(b'keep')
            trash = types.SimpleNamespace(send2trash=lambda path: None)
            with patch.dict('sys.modules', {'send2trash': trash}):
                worker = self.worker_class('delete', [str(source)])
                worker.run()
            self.assertTrue(source.exists())
            self.assertEqual(worker.ok_count, 1)

    def test_directory_copy_reports_actual_bytes(self):
        """复制目录时，字节统计正确且空目录被保留。

        场景：源目录包含一个空目录，以及大小分别为 3 和 4 字节的文件。
        预期：无失败，已复制与总字节数均为 7，目标中存在对应空目录。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source'
            destination = Path(directory) / 'destination'
            source.mkdir()
            destination.mkdir()
            (source / 'empty').mkdir()
            (source / 'one').write_bytes(b'123')
            (source / 'two').write_bytes(b'4567')
            worker = self.worker_class('copy', [str(source)], str(destination))
            worker.run()
            self.assertEqual(worker.fail_count, 0)
            self.assertEqual(worker.bytes_done, 7)
            self.assertEqual(worker.bytes_total, 7)
            self.assertTrue((destination / 'source' / 'empty').is_dir())

    def test_rename_preflight_preserves_all_sources_on_collision(self):
        """重命名前发现目标重名时，不修改任何一方的内容。

        场景：尝试把 old 重命名为已存在的 new，两者各有不同内容。
        预期：任务记录一次失败，old 和 new 的内容都保持不变。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'old'
            target = Path(directory) / 'new'
            source.write_bytes(b'original')
            target.write_bytes(b'existing')
            worker = self.worker_class('rename', [str(source)])
            worker.rename_plan = {str(source): str(target)}
            worker.run()
            self.assertEqual(worker.fail_count, 1)
            self.assertEqual(source.read_bytes(), b'original')
            self.assertEqual(target.read_bytes(), b'existing')

    def test_recycle_failure_never_falls_back_to_permanent_delete(self):
        """回收站操作失败时，绝不自动改为永久删除。

        场景：模拟 send2trash 抛出“回收站不可用”的错误。
        预期：任务记录失败，源文件仍然存在。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'keep'
            source.write_bytes(b'original')

            def fail_trash(path):
                raise OSError('recycle unavailable')

            with patch.dict('sys.modules', {'send2trash': types.SimpleNamespace(send2trash=fail_trash)}):
                worker = self.worker_class('delete', [str(source)])
                worker.run()
            self.assertEqual(worker.fail_count, 1)
            self.assertTrue(source.exists())

    def test_permanent_delete_refuses_reparse_directory(self):
        """永久删除拒绝目录链接，避免误删链接指向的数据。

        场景：模拟目录链接检测返回 True，再请求永久删除该目录。
        预期：任务记录失败，目录内文件保留；不创建或遍历真实系统链接。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'keep'
            source.write_bytes(b'original')
            worker = self.worker_class('permanent_delete', [directory])
            with patch.object(worker, '_is_reparse_path', return_value=True):
                worker.run()
            self.assertEqual(worker.fail_count, 1)
            self.assertTrue(source.exists())


class CompareAndWriteTests(unittest.TestCase):
    """第三组：目录比较、预览后的安全写入和重命名计划，共 4 个用例。"""

    def setUp(self):
        self.namespace = load_definitions({'_compare_directories', '_atomic_reviewed_write'},
                                          os=os, shutil=shutil, tr=lambda value: value)

    def test_compare_detects_content_and_missing_files(self):
        """目录比较能找到内容差异和单侧项目，排除相同文件。

        场景：两侧准备相同文件、同大小但内容不同的文件、单侧目录和单侧文件。
        预期：差异结果只包含 changed、only-left、only-right 三项。
        """
        with tempfile.TemporaryDirectory() as directory:
            left, right = Path(directory) / 'left', Path(directory) / 'right'
            left.mkdir()
            right.mkdir()
            (left / 'same').write_bytes(b'same')
            (right / 'same').write_bytes(b'same')
            (left / 'changed').write_bytes(b'123')
            (right / 'changed').write_bytes(b'456')
            (left / 'only-left').mkdir()
            (right / 'only-right').write_bytes(b'right')
            results = []
            self.namespace['_compare_directories'](str(left), str(right), threading.Event(), results.append)
            self.assertEqual({row[0] for row in results}, {'changed', 'only-left', 'only-right'})

    def test_reviewed_write_rejects_external_edit(self):
        """文件在预览后被外部修改时，拒绝用旧预览结果覆盖。

        场景：写入接口预期旧内容为 original，磁盘实际内容却是 external edit。
        预期：抛出错误，保留外部修改，且清理 .tabex-write- 临时文件。
        """
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'file'
            target.write_bytes(b'external edit')
            with self.assertRaises(ValueError):
                self.namespace['_atomic_reviewed_write'](str(target), b'original', b'AI edit')
            self.assertEqual(target.read_bytes(), b'external edit')
            self.assertEqual(list(Path(directory).glob('.tabex-write-*')), [])

    def test_reviewed_write_success_and_new_file_collision(self):
        """原子写入支持新建和更新，但新建模式不能覆盖已有文件。

        场景：新建文件，按匹配的旧内容更新，再尝试按新建模式覆盖。
        预期：前两步成功，第三步拒绝；最终内容及 CRLF 换行字节保持正确。
        """
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'file'
            write = self.namespace['_atomic_reviewed_write']
            write(str(target), None, b'original\r\n')
            write(str(target), b'original\r\n', b'changed\r\n')
            with self.assertRaises(FileExistsError):
                write(str(target), None, b'overwrite')
            self.assertEqual(target.read_bytes(), b'changed\r\n')

    def test_rename_plan_rejects_collision_and_invalid_names(self):
        """重命名预览计划拒绝重名和非法名称，正常替换则生成映射。

        场景：分别尝试改成已选中的 new.txt、Windows 保留名称 NUL.txt，以及合法新名。
        预期：前两种报错，合法情况返回 old.txt 到 new.txt 的映射；此处不实际重命名。
        """
        namespace = load_definitions({'_plan_batch_rename'}, tr=lambda value: value)
        plan = namespace['_plan_batch_rename']
        with self.assertRaises(ValueError):
            plan([r'C:\test\old.txt', r'C:\test\new.txt'], 'old', 'new')
        with self.assertRaises(ValueError):
            plan([r'C:\test\old.txt'], 'old', 'NUL')
        self.assertEqual(plan([r'C:\test\old.txt'], 'old', 'new'),
                         {r'C:\test\old.txt': r'C:\test\new.txt'})


class TaskPanelTests(unittest.TestCase):
    """第四组：全局文件任务面板与真实后台线程配合，共 1 个用例。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_panel_releases_worker_only_after_thread_finishes(self):
        """后台复制完成后，任务面板清除工作线程引用。

        场景：创建真实 Qt 任务面板和复制线程，处理界面事件等待完成，最多 5 秒。
        预期：面板没有运行中任务，worker 引用已清空，目标文件存在。
        """
        namespace = load_definitions(
            {'FileBatchOpWorker', 'FileTaskPanel'}, QThread=QThread, pyqtSignal=pyqtSignal,
            os=os, threading=threading, time=time, shutil=shutil, tr=lambda value: value,
            QDialog=QDialog, QVBoxLayout=QVBoxLayout, QHBoxLayout=QHBoxLayout,
            QTreeWidget=QTreeWidget, QTreeWidgetItem=QTreeWidgetItem, QPushButton=QPushButton,
            QTimer=QTimer, Qt=Qt, format_file_size=lambda value: str(value))
        namespace['_search_cache'] = types.SimpleNamespace(clear=lambda: None)
        namespace['_diagnostic_event'] = lambda *args: None
        panel = namespace['FileTaskPanel'](None)
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source'
            source.write_bytes(b'123')
            destination = Path(directory) / 'destination'
            destination.mkdir()
            panel.start_task('copy', [str(source)], str(destination))
            deadline = time.monotonic() + 5
            while panel.has_running_tasks() and time.monotonic() < deadline:
                self.app.processEvents()
            self.assertFalse(panel.has_running_tasks())
            self.assertIsNone(panel.records[0]['worker'])
            self.assertTrue((destination / 'source').exists())
            self.assertIn('成功 1 项，失败 0 项', panel.records[0]['row'].text(2))
            self.assertFalse(panel.cancel_button.isEnabled())
            self.assertFalse(panel.retry_button.isEnabled())
            self.assertFalse(panel.errors_button.isEnabled())
            self.assertTrue(panel.clear_button.isEnabled())
            self.assertEqual(panel.task_counts(), (0, 0))
            panel.clear_completed()
            self.assertFalse(panel.clear_button.isEnabled())
        panel.close()


class IntegrationTests(unittest.TestCase):
    """第五组：导入完整主程序，验证搜索、工作区及 Qt 对话框配合，共 9 个用例。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        import TabEx
        cls.module = TabEx

    def pump_until(self, predicate):
        """持续处理 Qt 界面事件，等待指定条件成立；超过 5 秒仍未成立则测试失败。"""
        deadline = time.monotonic() + 5
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
        self.assertTrue(predicate())

    def test_real_search_finishes_and_restores_buttons(self):
        """普通文件名搜索完成后，结果与按钮状态正确。

        场景：临时目录放入 needle.txt，在真实搜索窗口搜索 needle，关闭内容搜索。
        预期：显示一个结果，搜索按钮可用，停止按钮不可用；禁用 Everything 走本地搜索。
        """
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / 'needle.txt').write_text('content', encoding='utf-8')
            with patch.object(self.module, 'detect_everything', return_value=None):
                dialog = self.module.SearchDialog(directory)
            dialog.file_type_input.setText('txt')
            dialog.search_content_cb.setChecked(False)
            dialog.search_input.setEditText('needle')
            dialog.start_search()
            self.pump_until(lambda: not dialog.is_searching)
            self.assertEqual(dialog.result_model.rowCount(), 1)
            self.assertTrue(dialog.search_btn.isEnabled())
            self.assertFalse(dialog.stop_btn.isEnabled())
            dialog.close()

    def test_everything_early_return_restores_ui(self):
        """Everything 返回结果后，界面退出搜索状态并恢复排序。

        场景：真实搜索窗口使用模拟的 Everything 结果，返回一个临时文件路径。
        预期：显示一个结果，搜索按钮可用，停止按钮不可用，排序开启。
        模拟范围：不启动真实 es.exe 或 Everything 服务。
        """
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'needle.txt'
            path.write_text('content', encoding='utf-8')
            with patch.object(self.module, 'detect_everything', return_value='fake-es.exe'):
                dialog = self.module.SearchDialog(directory)
            dialog.search_input.setEditText('needle')
            with patch.object(self.module._SearchTask, 'search_with_everything', return_value=[str(path)]):
                dialog.start_search()
                self.pump_until(lambda: not dialog.is_searching)
            self.assertEqual(dialog.result_model.rowCount(), 1)
            self.assertTrue(dialog.search_btn.isEnabled())
            self.assertFalse(dialog.stop_btn.isEnabled())
            self.assertTrue(dialog.result_list.isSortingEnabled())
            dialog.close()

    def test_stop_then_restart_ignores_old_results(self):
        """停止搜索后立即重搜，只显示新任务的结果。

        场景：先搜索 old，立即停止，再搜索临时目录中实际存在的 new.txt。
        预期：旧任务取消标志已设置，最终只有 new.txt 一条结果。
        """
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'new.txt'
            source.write_text('new', encoding='utf-8')
            with patch.object(self.module, 'detect_everything', return_value=None):
                dialog = self.module.SearchDialog(directory)
            dialog.search_content_cb.setChecked(False)
            dialog.file_type_input.setText('txt')
            dialog.search_input.setEditText('old')
            dialog.start_search()
            previous = dialog._search_task
            dialog.stop_search()
            dialog.search_input.setEditText('new')
            dialog.start_search()
            self.pump_until(lambda: not dialog.is_searching)
            self.assertTrue(previous.cancel_event.is_set())
            rows = dialog.result_model.snapshot_rows()
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]['path'], str(source))
            dialog.close()

    def test_everything_failure_is_visible_and_finishes(self):
        """Everything 启动失败时，显示错误且允许再次搜索。

        场景：模拟启动子进程时抛出 ES unavailable 错误。
        预期：搜索结束，状态栏包含错误原因，搜索按钮重新可用。
        """
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(self.module, 'detect_everything', return_value='missing-es.exe'):
                dialog = self.module.SearchDialog(directory)
            with patch.object(self.module.subprocess, 'Popen', side_effect=OSError('ES unavailable')):
                dialog.start_search()
                self.pump_until(lambda: not dialog.is_searching)
            self.assertIn('ES unavailable', dialog.status_label.text())
            self.assertTrue(dialog.search_btn.isEnabled())
            dialog.close()

    def test_workspace_preserves_duplicates_groups_and_active_index(self):
        """保存工作区快照时，保留重复标签、固定状态和分组信息。

        场景：模拟左侧两个同路径标签，其中一个固定、另一个有颜色，右侧一个标签。
        预期：快照包含两组，左侧两个标签均保留，选中索引、固定标记和颜色正确。
        模拟范围：使用内存对象代表标签，不创建 Explorer，也不写配置文件。
        """
        tabs = [types.SimpleNamespace(current_path=r'C:\same', is_pinned=False, bookmark_group_color='red'),
                types.SimpleNamespace(current_path=r'C:\same', is_pinned=True),
                types.SimpleNamespace(current_path=r'D:\right', tab_group_separator_name='group')]
        left = types.SimpleNamespace(count=lambda: 2, widget=lambda index: tabs[index])
        right = types.SimpleNamespace(count=lambda: 1, widget=lambda index: tabs[2])
        owner = types.SimpleNamespace(_all_groups=lambda: [
            (types.SimpleNamespace(currentIndex=lambda: 1), left),
            (types.SimpleNamespace(currentIndex=lambda: 0), right)])
        result = self.module.MainWindow._capture_named_workspace(owner)
        self.assertEqual(len(result['groups']), 2)
        self.assertEqual(len(result['groups'][0]['tabs']), 2)
        self.assertEqual(result['groups'][0]['active_index'], 1)
        self.assertTrue(result['groups'][0]['tabs'][1]['is_pinned'])
        self.assertEqual(result['groups'][0]['tabs'][0]['bookmark_group_color'], 'red')

    def test_workspace_open_appends_and_restores_each_group(self):
        """打开工作区时追加标签，保留当前内容并恢复分屏。

        场景：模拟当前已有 existing 标签，再打开含两个左侧标签和一个右侧标签的工作区。
        预期：existing 不丢失，新标签按顺序追加，固定标记、左侧选中位置及分屏状态正确。
        模拟范围：标签创建、分屏和配置保存均用内存对象代替。
        """
        left = types.SimpleNamespace(items=[types.SimpleNamespace(current_path='existing')], selected=0)
        right = types.SimpleNamespace(items=[], selected=0)
        for group in (left, right):
            group.widget = lambda index, current=group: current.items[index]
            group.setCurrentIndex = lambda index, current=group: setattr(current, 'selected', index)
        state = {'groups': [
            {'tabs': [{'path': 'left-one', 'is_pinned': True}, {'path': 'left-two', 'bookmark_group_color': 'red'}], 'active_index': 1},
            {'tabs': [{'path': 'right-one'}], 'active_index': 0}]}
        owner = types.SimpleNamespace(
            config={'named_workspaces': {'test': state}}, _choose_named_workspace=lambda: 'test',
            _split_active=False, tab_widget=left, split_tab_widget=right,
            _content_stack_for=lambda target: target,
            _apply_tab_grouping_for_pane=lambda target: None,
            save_pinned_tabs=lambda: None, _schedule_session_snapshot=lambda: None)
        owner._activate_split_layout = lambda: setattr(owner, '_split_active', True)

        def add_tab(path, target_tabwidget, **kwargs):
            tab = types.SimpleNamespace(current_path=path, update_tab_title=lambda: None, **kwargs)
            target_tabwidget.items.append(tab)
            return len(target_tabwidget.items) - 1

        owner.add_new_tab = add_tab
        self.module.MainWindow.open_named_workspace(owner)
        self.assertEqual([tab.current_path for tab in left.items], ['existing', 'left-one', 'left-two'])
        self.assertTrue(left.items[1].is_pinned)
        self.assertEqual(left.selected, 2)
        self.assertEqual(right.items[0].current_path, 'right-one')
        self.assertTrue(owner._split_active)

    def test_slow_tab_activation_does_not_stat_or_scan(self):
        """激活慢路径标签时，不同步调用目录存在性检查。

        场景：模拟离线共享路径，并让任何 os.path.isdir 调用立即导致测试失败。
        预期：激活顺利完成，标签被标记为活动状态，没有调用 isdir。
        模拟范围：不连接真实网络盘，也不测响应时间或拦截所有文件系统接口。
        """
        tab = types.SimpleNamespace(current_path=r'\\offline\share', _is_slow_path=lambda path: True,
                                    _consume_pending_refresh=lambda **kwargs: None,
                                    _start_keepalive_sync=lambda: None)
        with patch.object(os.path, 'isdir', side_effect=AssertionError('UI filesystem call')):
            self.module.FileExplorerTab.set_refresh_active(tab, True)
        self.assertTrue(tab._refresh_active)

    def test_new_dialogs_construct_and_compare(self):
        """目录比较窗口能完成比较并生成非空截图。

        场景：创建真实 Qt 比较窗口，左侧临时目录有一个文件，右侧为空。
        预期：比较结束后显示一条差异，窗口截图对象非空，并保存到系统临时目录。
        边界：截图留作人工检查，此用例不自动判断布局或像素内容是否正确。
        """
        with tempfile.TemporaryDirectory() as directory:
            left, right = Path(directory) / 'left', Path(directory) / 'right'
            left.mkdir()
            right.mkdir()
            (left / 'only').write_text('left', encoding='utf-8')
            dialog = self.module.DirectoryCompareDialog(None, str(left), str(right))
            dialog.start_compare()
            self.pump_until(lambda: not dialog.running)
            self.assertEqual(dialog.table.topLevelItemCount(), 1)
            dialog.show()
            self.app.processEvents()
            self.assertFalse(dialog.grab().isNull())
            dialog.grab().save(os.path.join(tempfile.gettempdir(), 'TabEx-compare-check.png'))
            dialog.close()

    def test_preview_cancel_and_dialog_render(self):
        """差异预览显示变更文本，取消后不批准操作。

        场景：打开真实 Qt 预览窗口，展示 old 到 new 的变化，再自动执行取消。
        预期：预览包含删除行 -old，保存界面截图，确认函数返回 False。
        边界：本用例只验证预览与取消，不执行文件写入。
        """
        from PyQt5.QtWidgets import QPlainTextEdit

        def cancel_preview():
            dialog = self.app.activeModalWidget()
            self.assertIsNotNone(dialog)
            self.assertIn('-old', dialog.findChild(QPlainTextEdit).toPlainText())
            dialog.grab().save(os.path.join(tempfile.gettempdir(), 'TabEx-preview-check.png'))
            dialog.reject()

        QTimer.singleShot(50, cancel_preview)
        result = self.module._confirm_file_preview(None, 'Preview', 'sample.txt', 'old\n', 'new\n')
        self.assertFalse(result)


class FunctionalBoundaryTests(unittest.TestCase):
    """功能边界：真实临时文件配合可控错误，检查搜索和文件操作的分支。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        import TabEx
        cls.module = TabEx

    def search_rows(self, directory, keyword, filename=True, content=False, **options):
        """运行普通搜索的实际实现并收集结果；小数据集不会填满结果队列。"""
        task = self.module._SearchTask(directory, '', options.pop('limit', 100))
        self.module.SearchDialog.do_search(task, keyword, filename, content, **options)
        rows = []
        while not task.result_queue.empty():
            message = task.result_queue.get_nowait()
            if message['type'] == 'result_batch':
                rows.extend(message['items'])
        return rows

    def test_filename_case_word_and_extension_filters(self):
        """搜索大小写、整词和扩展名过滤可组合，排除大小写不同及部分词命中。"""
        with tempfile.TemporaryDirectory() as directory:
            for name in ('Needle.txt', 'needle.txt', 'Needles.txt', 'Needle.log'):
                (Path(directory) / name).write_bytes(b'')
            rows = self.search_rows(directory, 'Needle', file_types='txt',
                                    match_case=True, match_whole_word=True)
            self.assertEqual({Path(row['path']).name for row in rows}, {'Needle.txt'})

    def test_content_search_ascii_and_chinese_encodings(self):
        """内容搜索识别 ASCII 整词、UTF-8 中文和 GBK 中文，不把文件名当内容。"""
        with tempfile.TemporaryDirectory() as directory:
            contents = {'ascii.txt': b'Alpha alphabet', 'utf8.txt': '测试文字'.encode('utf-8'),
                        'gbk.txt': '测试文字'.encode('gbk'), 'Alpha.txt': b'unrelated'}
            for name, content in contents.items():
                (Path(directory) / name).write_bytes(content)
            for keyword, options, expected in (
                    ('Alpha', {'match_case': True, 'match_whole_word': True}, {'ascii.txt'}),
                    ('测试', {}, {'utf8.txt', 'gbk.txt'}),
                    ('absent', {}, set())):
                with self.subTest(keyword=keyword):
                    rows = self.search_rows(directory, keyword, False, True, file_types='txt', **options)
                    self.assertEqual({Path(row['path']).name for row in rows}, expected)

    def test_content_search_skips_known_binary_files(self):
        """二进制扩展名内即使包含关键字，也不会被当作文本内容搜索结果。"""
        with tempfile.TemporaryDirectory() as directory:
            (Path(directory) / 'data.exe').write_bytes(b'needle')
            (Path(directory) / 'data.txt').write_bytes(b'needle')
            rows = self.search_rows(directory, 'needle', False, True)
            self.assertEqual({Path(row['path']).name for row in rows}, {'data.txt'})

    def test_search_result_limit_and_cache_storage(self):
        """普通搜索遵守结果上限，完整的小结果集保存到缓存；界面复用另有流程测试。"""
        with tempfile.TemporaryDirectory() as directory:
            for index in range(30):
                (Path(directory) / f'needle-{index}.txt').write_bytes(b'text')
            self.assertEqual(len(self.search_rows(directory, 'needle', limit=7)), 7)
            cache = self.module.SearchCache()
            with patch.object(self.module, '_search_cache', cache):
                rows = self.search_rows(directory, 'needle', cache_key='test-key')
                self.assertEqual(len(rows), 30)
                self.assertEqual(cache.get('test-key'), rows)

    def test_everything_arguments_filter_and_result_limit(self):
        """Everything 参数保留目录范围及多个扩展名，结果再做名称过滤并限制数量。

        模拟子进程输出，不启动真实 Everything，也不对索引中的每个路径执行 exists。
        """
        task = self.module._SearchTask(r'C:\scope', 'fake-es.exe', 2)
        process = Mock(returncode=0)
        process.communicate.return_value = ('C:\\scope\\Needle.c\nC:\\scope\\other.h\n'
                                            'C:\\scope\\Needle.h\nC:\\scope\\Needle-more.c', '')
        process.poll.return_value = 0
        with patch.object(self.module.subprocess, 'Popen', return_value=process) as launch, \
                patch.object(os.path, 'exists', side_effect=AssertionError('逐条访问磁盘')):
            results = task.search_with_everything('Needle', task.search_path, '*.c,*.h', True, True)
        self.assertEqual(results, [r'C:\scope\Needle.c', r'C:\scope\Needle.h'])
        self.assertEqual(launch.call_args.args[0],
                         ['fake-es.exe', '-max-results', '2', 'path:"C:\\scope" Needle ext:c;h'])
        process.kill.assert_not_called()

    def test_everything_cancel_kills_and_reaps_process(self):
        """取消 Everything 搜索后终止并回收仍在运行的子进程，返回空结果。"""
        task = self.module._SearchTask('unused', 'fake-es.exe', 10)
        task.cancel_event.set()
        process = Mock()
        process.poll.return_value = None
        with patch.object(self.module.subprocess, 'Popen', return_value=process):
            self.assertEqual(task.search_with_everything('test', ''), [])
        process.kill.assert_called_once_with()
        process.communicate.assert_called_once_with()

    def test_everything_timeout_and_nonzero_exit_surface_errors(self):
        """Everything 超时与非零退出均上报错误，超时会清理子进程。

        用可控时钟直接越过截止时间，不实际等待 30 秒。
        """
        task = self.module._SearchTask('unused', 'fake-es.exe', 10)
        process = Mock(returncode=1)
        process.poll.return_value = None
        with patch.object(self.module.subprocess, 'Popen', return_value=process), \
                patch.object(self.module.time, 'monotonic', side_effect=[100, 131]):
            with self.assertRaises(RuntimeError):
                task.search_with_everything('test', '')
        process.kill.assert_called_once_with()
        process.poll.return_value = 1
        process.communicate.return_value = ('', 'index unavailable')
        with patch.object(self.module.subprocess, 'Popen', return_value=process):
            with self.assertRaisesRegex(RuntimeError, 'index unavailable'):
                task.search_with_everything('test', '')

    def test_copy_duplicate_names_keep_every_version(self):
        """反复复制同名文件时自动递增副本名，三个版本的字节内容都保留。"""
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'report.txt'
            destination = Path(directory) / 'destination'
            destination.mkdir()
            for value in (b'first', b'second', b'third'):
                source.write_bytes(value)
                worker = self.module.FileBatchOpWorker('copy', [str(source)], str(destination))
                worker.run()
                self.assertEqual(worker.ok_count, 1)
            self.assertEqual({path.name: path.read_bytes() for path in destination.iterdir()},
                             {'report.txt': b'first', 'report - copy.txt': b'second',
                              'report - copy2.txt': b'third'})

    def test_file_batch_continues_after_missing_source(self):
        """批量复制中一个源文件丢失时，记录失败项并继续复制其余有效文件。"""
        with tempfile.TemporaryDirectory() as directory:
            good = Path(directory) / 'good'
            missing = Path(directory) / 'missing'
            good.write_bytes(b'ok')
            destination = Path(directory) / 'destination'
            destination.mkdir()
            worker = self.module.FileBatchOpWorker('copy', [str(missing), str(good)], str(destination))
            worker.run()
            self.assertEqual((worker.ok_count, worker.fail_count), (1, 1))
            self.assertEqual(worker.failed_paths, [str(missing)])
            self.assertEqual((destination / 'good').read_bytes(), b'ok')

    def test_permanent_delete_nested_tree_preserves_neighbor(self):
        """显式永久删除仅移除选中的嵌套目录，旁边未选中的文件仍然存在。

        所有内容均在临时目录内，不删除用户文件或使用系统回收站。
        """
        with tempfile.TemporaryDirectory() as directory:
            selected = Path(directory) / 'selected'
            (selected / 'child' / 'empty').mkdir(parents=True)
            (selected / 'child' / 'file').write_bytes(b'delete')
            neighbor = Path(directory) / 'keep'
            neighbor.write_bytes(b'keep')
            worker = self.module.FileBatchOpWorker('permanent_delete', [str(selected)])
            worker.run()
            self.assertEqual((worker.ok_count, worker.fail_count), (1, 0))
            self.assertFalse(selected.exists())
            self.assertEqual(neighbor.read_bytes(), b'keep')

    def test_retry_remove_recovers_and_has_finite_attempts(self):
        """短暂删除失败允许重试，持续失败达到上限后抛错，不无限循环。

        删除、权限修复和等待均模拟，不修改实际文件权限。
        """
        worker = self.module.FileBatchOpWorker
        remove = Mock(side_effect=[PermissionError('busy'), None])
        with patch.object(worker, '_clear_readonly') as clear, patch.object(self.module.time, 'sleep'):
            worker._retry_remove_once_cleared(remove, 'unused', retries=2)
            self.assertEqual(remove.call_count, 2)
            clear.assert_called_once_with('unused')
            remove.reset_mock(side_effect=True)
            remove.side_effect = PermissionError('busy')
            with self.assertRaises(PermissionError):
                worker._retry_remove_once_cleared(remove, 'unused', retries=2)
            self.assertEqual(remove.call_count, 3)

    def test_parallel_task_errors_do_not_drop_other_tasks(self):
        """真实线程池中一个任务异常时，其余任务仍执行，错误携带失败任务名称。"""
        worker = self.module.FileBatchOpWorker('copy', [], max_workers=2)
        completed = []
        lock = threading.Lock()

        def execute(task):
            if task == 3:
                raise OSError('failed-task')
            with lock:
                completed.append(task)

        errors = worker._run_parallel_file_tasks(list(range(20)), execute, str)
        self.assertEqual(set(completed), set(range(20)) - {3})
        self.assertEqual(worker.done_units, 20)
        self.assertEqual(errors, ['3: failed-task'])

    def test_compare_type_mismatch_read_failure_and_cancel(self):
        """目录比较报告文件与目录类型冲突、读取错误；预先取消则不产生差异。"""
        with tempfile.TemporaryDirectory() as directory:
            left, right = Path(directory) / 'left', Path(directory) / 'right'
            left.mkdir()
            right.mkdir()
            (left / 'kind').mkdir()
            (right / 'kind').write_bytes(b'file')
            (left / 'locked').write_bytes(b'left')
            (right / 'locked').write_bytes(b'right')
            results = []
            with patch('filecmp.cmp', side_effect=PermissionError('test-locked')):
                self.module._compare_directories(str(left), str(right), threading.Event(), results.append)
            states = {row[0]: row[1] for row in results}
            self.assertEqual(states['kind'], self.module.tr('类型不同'))
            self.assertIn('test-locked', states['locked'])
            cancelled = threading.Event()
            cancelled.set()
            results.clear()
            self.module._compare_directories(str(left), str(right), cancelled, results.append)
            self.assertEqual(results, [])

    def test_atomic_write_commit_failure_preserves_original(self):
        """原子写入提交失败时保留原文件，并清除已经写好的临时文件。"""
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'file'
            target.write_bytes(b'original')
            with patch.object(os, 'replace', side_effect=PermissionError('busy')):
                with self.assertRaises(PermissionError):
                    self.module._atomic_reviewed_write(str(target), b'original', b'new')
            self.assertEqual(target.read_bytes(), b'original')
            self.assertEqual(list(Path(directory).glob('.tabex-write-*')), [])


class NormalUsageTests(unittest.TestCase):
    """模拟正常使用：真实 Qt 窗口、信号和后台线程，不启动原生 Explorer 主窗口。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        import TabEx
        cls.module = TabEx

    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        cache_patch = patch.object(self.module, '_search_cache', self.module.SearchCache())
        cache_patch.start()
        self.addCleanup(cache_patch.stop)

    def pump_until(self, predicate):
        """处理真实 Qt 事件；5 秒仅为防挂死超时，不作为性能达标阈值。"""
        deadline = time.monotonic() + 5
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
        self.assertTrue(predicate(), '等待 Qt 后台操作完成超时')

    def panel(self):
        """建立任务面板，失败时也先取消并回收线程，再清理临时文件。"""
        panel = self.module.FileTaskPanel(None)

        def cleanup():
            for record in panel.records:
                worker = record['worker']
                if worker is not None:
                    worker.request_cancel()
                    self.assertTrue(worker.wait(5000), '文件任务未退出')
            self.app.processEvents()
            panel.close()
            panel.deleteLater()

        self.addCleanup(cleanup)
        return panel

    def search_dialog(self, directory):
        """关闭 Everything，搜索结束或断言失败后都取消并回收本地搜索线程。"""
        with patch.object(self.module, 'detect_everything', return_value=None):
            dialog = self.module.SearchDialog(str(directory))
        dialog.search_content_cb.setChecked(False)
        dialog.file_type_input.setText('txt')

        def cleanup():
            thread = dialog.search_thread
            dialog.close()
            if thread is not None:
                thread.join(5)
                self.assertFalse(thread.is_alive(), '搜索线程未退出')
            dialog.deleteLater()

        self.addCleanup(cleanup)
        return dialog

    def compare_dialog(self, left, right):
        dialog = self.module.DirectoryCompareDialog(None, str(left), str(right))
        self.addCleanup(dialog.deleteLater)
        self.addCleanup(dialog.close)
        return dialog

    def test_copy_search_preview_rename_compare_workflow(self):
        """模拟工作流程：复制目录、搜索、确认编辑、批量改名、重新搜索和比较。

        预期：原目录不变，目标内容正确，重命名后旧名称消失，差异窗口显示两项。
        预览通过真实确定按钮接受；不调用 AI 服务或原生 Explorer。
        """
        from PyQt5.QtWidgets import QDialogButtonBox, QPlainTextEdit
        source = self.root / 'source'
        destination = self.root / 'destination'
        source.mkdir()
        destination.mkdir()
        (source / 'draft.txt').write_bytes(b'original\r\n')
        panel = self.panel()
        panel.start_task('copy', [str(source)], str(destination))
        self.pump_until(lambda: not panel.has_running_tasks())
        self.assertEqual(panel.records[0]['errors'], [])
        copied = destination / 'source'
        search = self.search_dialog(copied)
        search.search_input.setEditText('draft')
        search.search_btn.click()
        self.pump_until(lambda: not search.is_searching)
        self.assertEqual([row['path'] for row in search.result_model.snapshot_rows()],
                         [str(copied / 'draft.txt')])
        comparison = self.compare_dialog(source, copied)
        comparison.start_button.click()
        self.pump_until(lambda: not comparison.running)
        self.assertEqual(comparison.table.topLevelItemCount(), 0)
        preview = {}

        def accept_preview():
            dialog = self.app.activeModalWidget()
            if dialog is not None:
                preview['diff'] = dialog.findChild(QPlainTextEdit).toPlainText()
                buttons = dialog.findChild(QDialogButtonBox)
                preview['cancel_default'] = buttons.button(QDialogButtonBox.Cancel).isDefault()
                buttons.button(QDialogButtonBox.Ok).click()

        QTimer.singleShot(50, accept_preview)
        accepted = self.module._confirm_file_preview(None, 'Preview', 'draft.txt', 'original\r\n', 'edited\r\n')
        self.assertTrue(accepted)
        self.assertIn('+edited', preview['diff'])
        self.assertTrue(preview['cancel_default'])
        self.module._atomic_reviewed_write(str(copied / 'draft.txt'), b'original\r\n', b'edited\r\n')
        plan = self.module._plan_batch_rename([str(copied / 'draft.txt')], 'draft', 'final')
        panel.start_task('rename', list(plan), rename_plan=plan)
        self.pump_until(lambda: not panel.has_running_tasks())
        self.assertEqual(panel.records[-1]['errors'], [])
        search.search_input.setEditText('final')
        search.search_btn.click()
        self.pump_until(lambda: not search.is_searching)
        self.assertEqual([row['path'] for row in search.result_model.snapshot_rows()],
                         [str(copied / 'final.txt')])
        comparison.start_button.click()
        self.pump_until(lambda: not comparison.running)
        differences = {comparison.table.topLevelItem(index).text(0)
                       for index in range(comparison.table.topLevelItemCount())}
        self.assertEqual(differences, {'draft.txt', 'final.txt'})
        self.assertEqual((source / 'draft.txt').read_bytes(), b'original\r\n')
        self.assertEqual((copied / 'final.txt').read_bytes(), b'edited\r\n')
        self.assertFalse((copied / 'draft.txt').exists())
        self.assertFalse(panel.timer.isActive())
        panel.clear_completed()
        self.assertEqual(panel.table.topLevelItemCount(), 0)
        self.assertEqual(panel.records, [])

    def test_retry_only_failed_items_after_user_confirmation(self):
        """复制部分失败后先取消重试，再补齐源文件并确认重试，只处理失败项。

        预期：成功文件不生成重复副本，错误详情保留，任务结束后搜索缓存失效。
        确认框返回值模拟，其余任务及文件操作真实执行。
        """
        from PyQt5.QtWidgets import QMessageBox
        source = self.root / 'good.txt'
        missing = self.root / 'missing.txt'
        destination = self.root / 'destination'
        destination.mkdir()
        source.write_bytes(b'good')
        panel = self.panel()
        self.module._search_cache.put('before-copy', ['stale'])
        panel.start_task('copy', [str(source), str(missing)], str(destination))
        self.pump_until(lambda: not panel.has_running_tasks())
        record = panel.records[0]
        self.assertEqual(record['failed'], [str(missing)])
        self.assertTrue(record['errors'])
        self.assertIsNone(self.module._search_cache.get('before-copy'))
        panel.table.setCurrentItem(record['row'])
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.No):
            panel.retry_selected()
        self.assertEqual(len(panel.records), 1)
        missing.write_bytes(b'recovered')
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.Yes):
            panel.retry_selected()
        self.pump_until(lambda: not panel.has_running_tasks())
        self.assertEqual(panel.records[1]['paths'], [str(missing)])
        self.assertEqual(panel.records[1]['errors'], [])
        self.assertEqual({path.name for path in destination.iterdir()}, {'good.txt', 'missing.txt'})
        self.assertEqual((destination / 'missing.txt').read_bytes(), b'recovered')

    def test_two_tasks_keep_ui_responsive_across_panel_reopen(self):
        """两个后台任务运行时 Qt 仍处理心跳事件；关闭再打开面板不终止复制。

        用事件暂停实际文件复制，确定任务仍在运行，避免依赖电脑快慢。
        预期：两个任务各自完成，关闭面板不丢记录，全部结束才停止进度定时器。
        """
        source = self.root / 'source.txt'
        source.write_bytes(b'payload')
        destinations = [self.root / 'first', self.root / 'second']
        for destination in destinations:
            destination.mkdir()
        panel = self.panel()
        release = threading.Event()
        entered = [threading.Event(), threading.Event()]
        original_copy = self.module.FileBatchOpWorker._copy_file_task
        heartbeat = []
        timer = QTimer()
        timer.setInterval(0)
        timer.timeout.connect(lambda: heartbeat.append(True))

        def controlled_copy(worker, source_path, destination_path):
            index = 0 if Path(destination_path).parent == destinations[0] else 1
            entered[index].set()
            if not release.wait(5):
                raise RuntimeError('test gate timeout')
            return original_copy(worker, source_path, destination_path)

        with patch.object(self.module.FileBatchOpWorker, '_copy_file_task', controlled_copy):
            try:
                panel.show()
                for destination in destinations:
                    panel.start_task('copy', [str(source)], str(destination))
                timer.start()
                self.pump_until(lambda: all(event.is_set() for event in entered) and len(heartbeat) >= 3)
                self.assertTrue(panel.has_running_tasks())
                panel.refresh_progress()
                panel.clear_completed()
                self.assertEqual(len(panel.records), 2)
                panel.close()
                panel.show()
                self.assertTrue(panel.has_running_tasks())
            finally:
                timer.stop()
                release.set()
                self.pump_until(lambda: not panel.has_running_tasks())
        self.assertFalse(panel.timer.isActive())
        for destination in destinations:
            self.assertEqual((destination / source.name).read_bytes(), b'payload')
        self.assertTrue(all(record['worker'] is None and not record['errors'] for record in panel.records))

    def test_cancel_selected_task_then_start_another_copy(self):
        """用户取消选中复制任务后，可以立即进行下一次正常复制。

        用事件固定取消时机在文件提交之前；预期取消不留下目标或临时文件。
        """
        source = self.root / 'source.txt'
        source.write_bytes(b'payload')
        destination = self.root / 'destination'
        destination.mkdir()
        panel = self.panel()
        entered, release = threading.Event(), threading.Event()
        original_copy = self.module.FileBatchOpWorker._copy_file_task

        def controlled_copy(worker, source_path, destination_path):
            entered.set()
            if not release.wait(5):
                raise RuntimeError('test gate timeout')
            return original_copy(worker, source_path, destination_path)

        with patch.object(self.module.FileBatchOpWorker, '_copy_file_task', controlled_copy):
            try:
                panel.start_task('copy', [str(source)], str(destination))
                self.pump_until(entered.is_set)
                panel.table.setCurrentItem(panel.records[0]['row'])
                panel.cancel_selected()
                self.assertTrue(panel.records[0]['worker']._cancel_requested)
            finally:
                release.set()
                self.pump_until(lambda: not panel.has_running_tasks())
        self.assertEqual(list(destination.iterdir()), [])
        self.assertIn(self.module.tr('已取消'), panel.records[0]['row'].text(2))
        panel.start_task('copy', [str(source)], str(destination))
        self.pump_until(lambda: not panel.has_running_tasks())
        self.assertEqual((destination / source.name).read_bytes(), b'payload')
        self.assertEqual(panel.records[-1]['errors'], [])

    def test_repeated_search_window_open_search_close(self):
        """重复打开搜索窗口并切换关键字，首次扫描后复用缓存，不残留上次结果。

        五次真实本地搜索窗口，每个窗口先搜 alpha 再搜 beta，均校验结果路径。
        首次关闭后线程结束；后续窗口从缓存加载，不新建搜索线程。
        仅缓存使用固定时钟，避免慢机器上 5 秒 TTL 到期；Qt 和线程仍使用真实时间。
        """
        namespace = load_definitions(
            {'SearchCache'}, OrderedDict=OrderedDict, threading=threading, hashlib=hashlib,
            time=types.SimpleNamespace(monotonic=lambda: 100.0))
        self.module._search_cache = namespace['SearchCache']()
        for name in ('alpha.txt', 'beta.txt'):
            (self.root / name).write_bytes(b'content')
        for iteration in range(5):
            with self.subTest(iteration=iteration):
                dialog = self.search_dialog(self.root)
                dialog.show()
                for keyword in ('alpha', 'beta'):
                    dialog.search_input.setEditText(keyword)
                    dialog.search_btn.click()
                    self.pump_until(lambda: not dialog.is_searching)
                    self.assertEqual([row['path'] for row in dialog.result_model.snapshot_rows()],
                                     [str(self.root / f'{keyword}.txt')])
                thread = dialog.search_thread
                dialog.close()
                if iteration == 0:
                    self.assertIsNotNone(thread)
                    thread.join(5)
                    self.assertFalse(thread.is_alive())
                else:
                    self.assertIsNone(thread)

    def test_close_search_during_active_work_cancels_producer(self):
        """搜索进行中关闭窗口会通知后台退出，不必等待全部目录扫描完成。

        模拟扫描器等待取消事件，但窗口、任务对象和后台线程均真实运行。
        """
        dialog = self.search_dialog(self.root)
        entered = threading.Event()
        saw_cancel = threading.Event()

        def controlled_search(task, *args):
            entered.set()
            if task.cancel_event.wait(5):
                saw_cancel.set()

        with patch.object(self.module.SearchDialog, 'do_search', controlled_search):
            dialog.search_input.setEditText('test')
            dialog.search_btn.click()
            thread = dialog.search_thread
            try:
                self.pump_until(entered.is_set)
            finally:
                dialog.close()
                thread.join(5)
        self.assertTrue(saw_cancel.is_set())
        self.assertFalse(thread.is_alive())

    def test_compare_invalid_path_then_correct_and_retry(self):
        """比较目录输错路径后显示错误；修正路径重新比较，可正常显示差异。

        预期：失败和成功两次结束后都恢复比较按钮，停止按钮禁用。
        """
        left, right = self.root / 'left', self.root / 'right'
        left.mkdir()
        right.mkdir()
        (left / 'only.txt').write_bytes(b'left')
        dialog = self.compare_dialog(left, self.root / 'missing')
        dialog.start_button.click()
        self.pump_until(lambda: not dialog.running)
        self.assertIn('missing', dialog.status.text())
        self.assertTrue(dialog.start_button.isEnabled())
        self.assertFalse(dialog.stop_button.isEnabled())
        dialog.paths[1].setText(str(right))
        dialog.start_button.click()
        self.pump_until(lambda: not dialog.running)
        self.assertEqual(dialog.table.topLevelItemCount(), 1)
        self.assertEqual(dialog.table.topLevelItem(0).text(0), 'only.txt')
        self.assertTrue(dialog.start_button.isEnabled())
        self.assertFalse(dialog.stop_button.isEnabled())


class InterfaceTests(unittest.TestCase):
    """交互回归：真实 Qt 控件验证可用状态、活动侧和布局，不启动原生 Explorer。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        import TabEx
        cls.module = TabEx

    def pump_until(self, predicate):
        deadline = time.monotonic() + 5
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
        self.assertTrue(predicate())

    def pinning_host(self):
        """真实 Qt 标签及内容栈，隔离配置写入与原生 Shell，只保留固定和分组逻辑。"""
        from PyQt5.QtWidgets import QWidget, QTabWidget, QStackedWidget
        module = self.module

        class Host(QWidget):
            pin_tab = module.MainWindow.pin_tab
            unpin_tab = module.MainWindow.unpin_tab
            sort_tabs_by_pinned = module.MainWindow.sort_tabs_by_pinned
            _refresh_tab_labels = module.MainWindow._refresh_tab_labels
            _apply_tab_grouping_for_pane = module.MainWindow._apply_tab_grouping_for_pane
            _apply_tab_group_color = module.MainWindow._apply_tab_group_color

        owner = Host()
        self.addCleanup(owner.deleteLater)
        owner.config = {'show_tab_group_markers': True}
        groups = [(QTabWidget(owner), QStackedWidget(owner)) for _index in range(2)]
        owner.tab_widget, owner.content_stack = groups[0]
        owner.split_tab_widget, owner.split_content_stack = groups[1]
        owner._all_groups = lambda: groups
        owner._resolve_group = lambda target: (groups[1] + (True,) if target is groups[1][0] else groups[0] + (False,))
        owner._content_stack_for = lambda target: groups[1][1] if target is groups[1][0] else groups[0][1]
        owner.save_pinned_tabs = Mock()
        owner._schedule_session_snapshot = Mock()
        for side, (tabs, stack) in enumerate(groups):
            bar = module.CustomTabBar()
            bar.main_window = owner
            bar.owner_tabwidget = tabs
            tabs.setTabBar(bar)
            for name, pinned, color in (('first', True, ''), ('second', True, ''),
                                        ('group-one', False, '#64B5F6'), ('group-two', False, '#64B5F6'),
                                        ('plain', False, '')):
                pane = QWidget()
                pane.current_path = rf'C:\side-{side}\{name}'
                pane.is_pinned = pinned
                pane.bookmark_group_color = color
                pane.update_tab_title = owner._refresh_tab_labels
                stack.addWidget(pane)
                tabs.addTab(QWidget(), name)
            tabs.setCurrentIndex(4)
        owner._refresh_tab_labels()
        return owner, groups

    def test_pinning_preserves_existing_red_pin_after_group_refresh(self):
        """再固定一个标签后，旧固定标签仍显示原红色图钉；重复分组刷新和标题缓存不清除它。"""
        owner, groups = self.pinning_host()
        for tabs, stack in groups:
            with self.subTest(side=groups.index((tabs, stack))):
                original_pins = [stack.widget(0), stack.widget(1)]
                newly_pinned = stack.widget(4)
                owner.pin_tab(4, target_tabwidget=tabs)
                owner._apply_tab_grouping_for_pane(tabs)
                owner._refresh_tab_labels()
                for pane in original_pins + [newly_pinned]:
                    self.assertTrue(pane.is_pinned)
                    icon = tabs.tabIcon(stack.indexOf(pane))
                    self.assertEqual(icon.cacheKey(), self.module._pinned_tab_icon().cacheKey())
                self.assertIs(stack.widget(tabs.currentIndex()), newly_pinned)
                self.assertEqual(newly_pinned.bookmark_group_color, '')

    def test_unpin_moves_plain_tab_to_end_and_preserves_other_order(self):
        """取消固定且无分组的标签移到所属侧最右端，其他固定/分组/普通标签顺序和当前选中不变。"""
        owner, groups = self.pinning_host()
        for tabs, stack in groups:
            with self.subTest(side=groups.index((tabs, stack))):
                before = [stack.widget(index) for index in range(stack.count())]
                pages = [tabs.widget(index) for index in range(tabs.count())]
                tabs.setCurrentIndex(0)
                owner.unpin_tab(0, target_tabwidget=tabs)
                self.assertEqual([stack.widget(index) for index in range(stack.count())], before[1:] + before[:1])
                self.assertEqual([tabs.widget(index) for index in range(tabs.count())], pages[1:] + pages[:1])
                self.assertIs(stack.widget(tabs.currentIndex()), before[0])
                self.assertTrue(tabs.tabIcon(tabs.currentIndex()).isNull())
                self.assertEqual(tabs.tabIcon(0).cacheKey(), self.module._pinned_tab_icon().cacheKey())
                self.assertEqual([pane.bookmark_group_color for pane in before[2:4]], ['#64B5F6', '#64B5F6'])

    def test_long_pinned_tab_keeps_original_red_marker_visible(self):
        """长名称左侧省略时保留独立红图钉；固定和取消固定不会改变红钉外观。"""
        from PyQt5.QtWidgets import QStyle, QStyleOptionTab
        owner, groups = self.pinning_host()
        tabs, stack = groups[0]
        bar = tabs.tabBar()
        bar.setStyleSheet('QTabBar::tab { width: 120px; height: 28px; padding: 2px 4px; }')
        bar.setElideMode(Qt.ElideLeft)
        stack.widget(0).current_path = r'C:\some-very-long-parent\some-very-long-directory-name'
        owner._refresh_tab_labels()
        owner.resize(900, 100)
        tabs.resize(880, 90)
        tabs.show()
        groups[1][0].hide()
        owner.show()
        self.app.processEvents()
        option = QStyleOptionTab()
        bar.initStyleOption(option, 0)
        marker = bar.style().subElementRect(QStyle.SE_TabBarTabText, option, bar)
        self.assertFalse(option.icon.isNull())
        self.assertGreater(marker.left(), bar.tabRect(0).left())
        image = bar.grab(bar.tabRect(0)).toImage()
        red_pixels = sum(image.pixelColor(x_pos, y_pos).alpha() > 0
                         and image.pixelColor(x_pos, y_pos).red() > 150
                         and image.pixelColor(x_pos, y_pos).green() < 120
                         for y_pos in range(image.height()) for x_pos in range(image.width()))
        self.assertGreater(red_pixels, 5)
        self.assertTrue(bar.grab().save(os.path.join(tempfile.gettempdir(), 'TabEx-pin-check.png')))
        owner.close()

    def test_task_partial_failure_success_and_notification_actions(self):
        """部分失败可重试、查看记录，成功任务不能重试；通知动作打开真实目标位置。"""
        from PyQt5.QtWidgets import QWidget
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'one.txt'
            source.write_bytes(b'payload')
            missing = Path(directory) / 'missing.txt'
            destination = Path(directory) / 'destination'
            destination.mkdir()
            owner = QWidget()
            owner.add_new_tab = Mock()
            panel = self.module.FileTaskPanel(owner)
            try:
                with patch.object(self.module, 'show_toast') as toast:
                    panel.start_task('copy', [str(source), str(missing)], str(destination))
                    self.pump_until(lambda: not panel.has_running_tasks())
                    self.assertIn('部分失败：成功 1 项，失败 1 项', panel.records[0]['row'].text(2))
                    self.assertTrue(panel.retry_button.isEnabled())
                    self.assertTrue(panel.errors_button.isEnabled())
                    self.assertFalse(panel.cancel_button.isEnabled())
                    self.assertEqual(panel.task_counts(), (0, 1))
                    toast.assert_called_once()
                    toast.call_args.kwargs['action']()
                    self.assertTrue(panel.isVisible())
                    self.assertEqual(panel.selected_record()['failed'], [str(missing)])
                    panel.hide()
                    toast.reset_mock()
                    panel.start_task('copy', [str(source)], str(destination))
                    self.pump_until(lambda: not panel.has_running_tasks())
                    self.assertFalse(panel.retry_button.isEnabled())
                    self.assertFalse(panel.errors_button.isEnabled())
                    toast.assert_called_once()
                    self.assertEqual(toast.call_args.kwargs['action_text'], '打开位置')
                    toast.call_args.kwargs['action']()
                    owner.add_new_tab.assert_called_once_with(str(destination))
                    panel.table.setCurrentItem(panel.records[0]['row'])
                    self.assertTrue(panel.retry_button.isEnabled())
            finally:
                for record in panel.records:
                    if record['worker'] is not None:
                        record['worker'].request_cancel()
                        record['worker'].wait(5000)
                self.app.processEvents()
                panel.close()
                owner.deleteLater()

    def test_action_toast_is_nonmodal_clickable_and_pauses_on_hover(self):
        """各级提示使用明显的实色背景和白字；不抢激活，动作可点击，悬停暂停且关闭清理。"""
        from PyQt5.QtWidgets import QWidget, QLabel
        from PyQt5.QtGui import QPalette
        from PyQt5.QtCore import QEvent
        owner = QWidget()
        owner.resize(800, 600)
        self.addCleanup(owner.deleteLater)
        colors = {'info': '#1767b5', 'success': '#18733b', 'warning': '#946000',
                  'error': '#b42318', 'critical': '#b42318', 'unknown': '#1767b5'}
        for level, background in colors.items():
            with self.subTest(level=level):
                action = Mock()
                toast = self.module.show_toast(owner, '文件任务', '成功 3 项，失败 0 项', level=level,
                                               action_text='打开位置', action=action)
                try:
                    self.app.processEvents()
                    self.assertTrue(toast.testAttribute(Qt.WA_ShowWithoutActivating))
                    self.assertFalse(toast.testAttribute(Qt.WA_TransparentForMouseEvents))
                    self.assertLessEqual(toast.width(), 420)
                    self.assertFalse(toast.isModal())
                    pixmap = toast.grab()
                    image = pixmap.toImage()
                    self.assertEqual(image.pixelColor(round(8 * pixmap.devicePixelRatio()),
                                                      image.height() // 2).name(), background)
                    for label in toast.findChildren(QLabel):
                        self.assertEqual(label.palette().color(QPalette.WindowText).name(), '#ffffff')
                    self.assertEqual(toast.action_button.palette().color(QPalette.ButtonText).name(), '#202020')
                    self.assertTrue(pixmap.save(os.path.join(tempfile.gettempdir(), f'TabEx-toast-{level}.png')))
                    self.app.sendEvent(toast, QEvent(QEvent.Enter))
                    self.assertFalse(toast._timer.isActive())
                    self.app.sendEvent(toast, QEvent(QEvent.Leave))
                    self.assertTrue(toast._timer.isActive())
                    toast.action_button.click()
                    action.assert_called_once_with()
                    self.assertFalse(toast._timer.isActive())
                    self.assertNotIn(toast, self.module._active_toasts)
                finally:
                    toast.close()

    def test_toasts_with_different_heights_do_not_overlap_after_dismissal(self):
        """长短通知按实际高度叠放，关闭底部通知后重排；新通知不会覆盖剩余提示。"""
        from PyQt5.QtWidgets import QWidget
        owner = QWidget()
        owner.resize(900, 700)
        self.addCleanup(owner.deleteLater)
        with patch.object(self.module, '_active_toasts', []):
            toasts = []
            try:
                for message in ('长结果\n' * 4, '短结果'):
                    toasts.append(self.module.show_toast(owner, '文件任务', message,
                                                        action_text='打开位置', action=Mock()))
                self.app.processEvents()
                self.assertFalse(toasts[0].geometry().intersects(toasts[1].geometry()))
                old_position = toasts[1].pos()
                toasts[0].close()
                self.assertGreater(toasts[1].y(), old_position.y())
                toasts.append(self.module.show_toast(owner, '文件任务', '成功 3 项，失败 0 项',
                                                    action_text='打开位置', action=Mock()))
                self.assertFalse(toasts[1].geometry().intersects(toasts[2].geometry()))
                image = toasts[2].grab()
                self.assertFalse(image.isNull())
                self.assertTrue(image.save(os.path.join(tempfile.gettempdir(), 'TabEx-toast-check.png')))
            finally:
                for toast in list(self.module._active_toasts):
                    toast.close()

    def test_address_bar_mouse_and_edit_focus_activate_owning_pane(self):
        """鼠标点击路径栏、键盘进入编辑框都会激活所属侧，不改变当前目录。"""
        from PyQt5.QtTest import QTest
        from PyQt5.QtCore import QEvent
        from PyQt5.QtWidgets import QWidget, QStackedWidget
        stack = QStackedWidget()
        self.addCleanup(stack.deleteLater)
        pane = QWidget()
        stack.addWidget(pane)
        tabs = object()
        owner = types.SimpleNamespace(_all_groups=lambda: [(tabs, stack)], set_active_pane_to_group=Mock())
        pane.main_window = owner
        bar = self.module.SimplePathBar(pane)
        bar.set_path(r'C:\test')
        bar.activated.connect(lambda: self.module.FileExplorerTab._activate_path_bar(pane))
        QTest.mousePress(bar, Qt.LeftButton)
        owner.set_active_pane_to_group.assert_called_with(tabs)
        owner.set_active_pane_to_group.reset_mock()
        self.app.sendEvent(bar._edit, QEvent(QEvent.FocusIn))
        owner.set_active_pane_to_group.assert_called_once_with(tabs)
        self.assertEqual(bar._current_path, r'C:\test')

    def test_window_and_dialog_layouts_at_two_font_scales(self):
        """真实主窗口及搜索窗口在中英文、100%/150% 字号下布局不重叠并保存截图。

        主窗口使用空配置和内存书签，禁用延迟加载、Shell、全局快捷键轮询及持久化。
        截图位于系统临时目录 tabex-ui-checks；不等同于原生 Explorer 的端到端测试。
        """
        from PyQt5.QtGui import QFont
        from PyQt5.QtWidgets import QAbstractButton
        output = Path(tempfile.gettempdir()) / 'tabex-ui-checks'
        output.mkdir(exist_ok=True)
        original_font = self.app.font()
        original_language = self.module._app_language
        try:
            for language in ('zh', 'en'):
                for scale in (1.0, 1.5):
                    with self.subTest(language=language, scale=scale), tempfile.TemporaryDirectory() as directory, ExitStack() as patches:
                        self.module._set_app_language(language)
                        self.app.setFont(QFont('Segoe UI', round(9 * scale)))
                        config = {'enable_explorer_monitor': False, 'enable_title_shortcuts': False,
                                  'title_shortcuts': [], 'ai_chat': {'enabled': False, 'panel_visible': False}}
                        patches.enter_context(patch.object(self.module.MainWindow, 'load_config', return_value=config))
                        for method in ('ensure_default_bookmarks', '_delayed_initialization', '_check_shortcuts',
                                       '_start_shortcut_listener', '_maybe_auto_check_updates',
                                       'save_config', 'save_session_snapshot', 'save_pinned_tabs', '_setup_window_icon',
                                       'add_new_tab'):
                            patches.enter_context(patch.object(self.module.MainWindow, method))
                        manager = patches.enter_context(patch.object(self.module, 'BookmarkManager'))
                        manager.return_value.get_tree.return_value = {
                            'bookmark_bar': {'type': 'folder', 'children': [], 'id': '1'},
                            'other': {'type': 'folder', 'children': [], 'id': '2'}}
                        patches.enter_context(patch.object(self.module, 'get_app_data_path',
                                                          side_effect=lambda *parts: os.path.join(directory, *parts)))
                        patches.enter_context(patch.object(self.module, 'detect_everything', return_value=None))
                        window = self.module.MainWindow()
                        window.server_socket = Mock()
                        search = self.module.SearchDialog(directory)
                        try:
                            width = 1000 if scale == 1 else 1280
                            window.resize(width, 700)
                            window.show()
                            search.resize(680 if scale == 1 else 980, 580)
                            search.show()
                            search.advanced_button.click()
                            self.app.processEvents()
                            buttons = [button for button in window.titlebar_widget.findChildren(QAbstractButton)
                                       if button.isVisibleTo(window.titlebar_widget)]
                            previous_right = -1
                            for button in sorted(buttons, key=lambda item: item.mapTo(window.titlebar_widget, item.rect().topLeft()).x()):
                                position = button.mapTo(window.titlebar_widget, button.rect().topLeft())
                                self.assertGreaterEqual(position.x(), previous_right)
                                self.assertLessEqual(position.x() + button.width(), window.titlebar_widget.width())
                                self.assertFalse(button.icon().isNull(), button.toolTip())
                                self.assertTrue(button.accessibleName())
                                previous_right = position.x() + button.width()
                            self.assertEqual(search.width(), 680 if scale == 1 else 980)
                            for control in (search.search_input, search.search_btn, search.stop_btn, search.path_input,
                                            search.advanced_button, search.ai_summary_btn):
                                self.assertTrue(search.rect().contains(control.geometry()))
                            self.assertIn('Ctrl+F', window.search_button.toolTip())
                            window.config['hotkey_bindings'] = {'search': 'Ctrl+E'}
                            window._apply_hotkey_bindings()
                            self.assertIn('Ctrl+E', window.search_button.toolTip())
                            self.assertIn('Ctrl+E', window.search_button.accessibleName())
                            window.config['hotkey_bindings'] = {}
                            window._apply_hotkey_bindings()
                            for name, widget in (('window', window), ('search', search)):
                                image = widget.grab()
                                self.assertFalse(image.isNull())
                                self.assertTrue(image.save(str(output / f'{name}-{language}-{scale}.png')))
                        finally:
                            search.close()
                            window.close()
                            window.deleteLater()
                            self.app.processEvents()
        finally:
            self.app.setFont(original_font)
            self.module._set_app_language(original_language)

    def test_search_advanced_options_keep_filters_and_fit_narrow_window(self):
        """高级选项折叠不清除条件；窄窗口中搜索、停止和高级按钮互不重叠。"""
        with patch.object(self.module, 'detect_everything', return_value=None):
            dialog = self.module.SearchDialog(str(SOURCE.parent))
        try:
            dialog.resize(640, 520)
            dialog.show()
            self.app.processEvents()
            self.assertTrue(dialog.advanced_options.isHidden())
            dialog.advanced_button.click()
            self.assertFalse(dialog.advanced_options.isHidden())
            dialog.match_case_cb.setChecked(True)
            dialog.match_whole_word_cb.setChecked(True)
            dialog.advanced_button.click()
            self.assertTrue(dialog.advanced_options.isHidden())
            self.assertTrue(dialog.match_case_cb.isChecked())
            self.assertTrue(dialog.match_whole_word_cb.isChecked())
            self.assertEqual(dialog.advanced_button.text(), '高级 (2)')
            self.app.processEvents()
            self.assertLess(dialog.search_input.geometry().right(), dialog.search_btn.geometry().left())
            self.assertLess(dialog.search_btn.geometry().right(), dialog.stop_btn.geometry().left())
            for button in (dialog.search_btn, dialog.stop_btn, dialog.ai_summary_btn, dialog.advanced_button):
                self.assertTrue(dialog.rect().contains(button.geometry()))
            self.assertFalse(dialog.search_btn.icon().isNull())
            self.assertFalse(dialog.stop_btn.icon().isNull())
            self.assertEqual(dialog.file_type_input.text(), '*.c,*.h,*.xdm,*.arxml,*.xml')
        finally:
            dialog.close()

    def test_toolbar_icons_and_scaled_dialog_layouts(self):
        """中英文及放大字号下检查真实工具栏/搜索/任务控件不越界，截图供人工复核。

        不创建原生 Explorer、不保存用户配置；布局模拟倍率并非真实多显示器 DPI 切换。
        """
        from PyQt5.QtCore import QSize
        from PyQt5.QtGui import QFont
        from PyQt5.QtWidgets import QWidget, QToolButton, QStyle
        module = self.module

        class Host(QWidget):
            create_custom_titlebar = module.MainWindow.create_custom_titlebar
            _populate_workspace_tools_menu = module.MainWindow._populate_workspace_tools_menu
            _update_task_indicator = module.MainWindow._update_task_indicator

        screenshot_dir = Path(tempfile.gettempdir()) / 'TabEx-ui-check'
        screenshot_dir.mkdir(exist_ok=True)
        saved_font = self.app.font()
        saved_language = module._app_language
        try:
            for language, scale in (('zh', 1.0), ('en', 1.5)):
                with self.subTest(language=language, scale=scale):
                    module._set_app_language(language)
                    font = QFont('Segoe UI', round(9 * scale))
                    self.app.setFont(font)
                    host = Host()
                    host.config = {'ai_chat': {'enabled': True}, 'title_shortcuts': []}
                    host.dpi_scale = scale
                    for name in (
                            'on_title_shortcut_dropped', 'open_title_shortcut', 'on_title_shortcuts_changed',
                            'refresh_title_shortcuts_ui', 'open_tortoisegit_log_current_tab',
                            'open_tortoisegit_commit_current_tab', 'open_git_bash_current_tab',
                            'open_cmd_current_tab', 'open_powershell_current_tab', 'open_calculator',
                            'go_back_current_tab', 'go_forward_current_tab', 'add_new_tab', 'reopen_closed_tab',
                            'show_tab_list', 'show_search_dialog', 'show_file_tasks', 'show_directory_compare',
                            'preview_batch_rename', 'save_named_workspace', 'open_named_workspace',
                            'delete_named_workspace', 'permanently_delete_selected', 'export_diagnostics',
                            'insert_tab_group_marker', 'toggle_split_view', 'show_bookmark_manager_dialog',
                            'show_settings_menu', 'toggle_chat_panel'):
                        setattr(host, name, Mock())
                    layout = QVBoxLayout(host)
                    host.create_custom_titlebar(layout)
                    host.resize(int(1000 * scale), int(80 * scale))
                    host.show()
                    search = None
                    panel = None
                    try:
                        self.app.processEvents()
                        buttons = [getattr(host, name) for name in (
                            'git_log_button', 'git_commit_button', 'git_bash_button', 'cmd_button',
                            'powershell_button', 'calculator_button', 'back_button', 'forward_button',
                            'add_tab_button', 'reopen_tab_button', 'tab_list_button', 'search_button',
                            'workspace_tools_button', 'file_tasks_button', 'insert_group_btn', 'split_view_btn',
                            'bookmark_button', 'settings_button', 'ai_chat_btn')]
                        for button in buttons:
                            self.assertFalse(button.icon().isNull(), button.toolTip())
                            self.assertTrue(button.toolTip())
                            self.assertTrue(host.titlebar_widget.rect().contains(button.geometry()), button.toolTip())
                            if button is not host.file_tasks_button:
                                self.assertEqual(button.size(), QSize(int(32 * scale), int(32 * scale)))
                        for previous, following in zip(buttons, buttons[1:]):
                            self.assertLess(previous.geometry().right(), following.geometry().left())
                        self.assertTrue(host.grab().save(str(screenshot_dir / f'toolbar-{language}.png')))
                        with patch.object(module, 'detect_everything', return_value=None):
                            search = module.SearchDialog(str(screenshot_dir))
                        search.resize(int(700 * scale), int(520 * scale))
                        search.show()
                        search.advanced_button.click()
                        search.match_case_cb.setChecked(True)
                        self.app.processEvents()
                        label = search.layout().itemAt(0).layout().itemAt(0).widget()
                        self.assertGreaterEqual(label.width(), label.fontMetrics().horizontalAdvance(label.text()))
                        self.assertTrue(search.rect().contains(search.advanced_button.geometry()))
                        self.assertGreater(search.advanced_button.width(),
                                           search.advanced_button.fontMetrics().horizontalAdvance(search.advanced_button.text()))
                        self.assertTrue(search.grab().save(str(screenshot_dir / f'search-{language}.png')))
                        panel = module.FileTaskPanel(None)
                        panel.resize(int(960 * scale), int(400 * scale))
                        row = QTreeWidgetItem([module.tr('复制'), r'C:\work\destination',
                                              module.tr('{}：成功 {} 项，失败 {} 项').format(module.tr('部分失败'), 8, 2),
                                              '128 MB'])
                        row.setIcon(2, panel.style().standardIcon(QStyle.SP_MessageBoxWarning))
                        panel.table.addTopLevelItem(row)
                        panel.show()
                        self.app.processEvents()
                        self.assertGreaterEqual(panel.table.columnWidth(2),
                                                panel.table.fontMetrics().horizontalAdvance(row.text(2)) + 20)
                        for button in (panel.cancel_button, panel.retry_button, panel.errors_button,
                                       panel.open_button, panel.clear_button):
                            self.assertTrue(panel.rect().contains(button.geometry()))
                            self.assertGreaterEqual(button.width(), button.fontMetrics().horizontalAdvance(button.text()) + 20)
                        self.assertTrue(panel.grab().save(str(screenshot_dir / f'tasks-{language}.png')))
                    finally:
                        if search is not None:
                            search.close()
                        if panel is not None:
                            panel.close()
                            panel.deleteLater()
                        host.close()
                        host.deleteLater()
        finally:
            module._set_app_language(saved_language)
            self.app.setFont(saved_font)

    def test_tools_menu_groups_danger_and_preserves_actions(self):
        """工具菜单区分文件、工作区、危险操作和诊断，保留原操作连接。"""
        from PyQt5.QtWidgets import QMenu, QToolButton
        callbacks = {name: Mock() for name in (
            'show_file_tasks', 'show_directory_compare', 'preview_batch_rename', 'save_named_workspace',
            'open_named_workspace', 'delete_named_workspace', 'permanently_delete_selected', 'export_diagnostics',
            'check_for_updates')}
        owner = types.SimpleNamespace(style=self.app.style, **callbacks)
        menu = QMenu()
        self.addCleanup(menu.deleteLater)
        self.module.MainWindow._populate_workspace_tools_menu(owner, menu)
        sections = [action.text() for action in menu.actions() if action.isSeparator()]
        self.assertEqual(sections, ['文件操作', '危险操作', '诊断'])
        workspaces = next(action.menu() for action in menu.actions() if action.menu())
        self.assertEqual(len(workspaces.actions()), 3)
        workspaces.actions()[1].trigger()
        callbacks['open_named_workspace'].assert_called_once()
        owner.export_diagnostics_action.trigger()
        callbacks['export_diagnostics'].assert_called_once()
        next(action for action in menu.actions() if action.text() == '检查更新...').trigger()
        callbacks['check_for_updates'].assert_called_once_with(manual=True)
        owner.file_tasks_button = QToolButton()
        self.addCleanup(owner.file_tasks_button.deleteLater)
        owner._file_task_panel = types.SimpleNamespace(task_counts=lambda: (2, 1))
        self.module.MainWindow._update_task_indicator(owner)
        self.assertEqual(owner.file_tasks_button.text(), '2')
        self.assertIn('运行 2，失败 1', owner.file_tasks_button.toolTip())
        self.assertIn('运行 2，失败 1', owner.file_tasks_action.text())

    def test_tab_labels_disambiguate_paths_duplicates_and_roots_without_io(self):
        """同名目录使用最短父路径区分，重复路径编号，根目录和特殊目录可读；不访问磁盘。"""
        paths = [r'C:\alpha\src', r'C:\beta\src', r'C:\alpha\src', 'D:\\', 'shell:RecycleBinFolder']
        with patch.object(os, 'scandir', side_effect=AssertionError('UI disk scan')), \
                patch.object(os.path, 'exists', side_effect=AssertionError('UI disk stat')):
            labels = self.module._tab_display_labels(paths)
        self.assertEqual(labels, [r'alpha\src [1]', r'beta\src', r'alpha\src [2]', 'D:\\', '回收站'])

    def test_toolbar_fallback_icons_are_distinct_and_render_without_native_apps(self):
        """程序未安装时使用本地SVG，所有工具图标可渲染且不同操作不再共用通用图标。"""
        from PyQt5.QtWidgets import QToolButton, QStyle
        button = QToolButton()
        self.addCleanup(button.deleteLater)
        images = {}
        with patch.object(self.module, '_native_tool_executable', return_value=None), \
                patch.dict(self.module._TOOL_NATIVE_CACHE, {}, clear=True):
            for name in self.module._TOOL_ICON_FILES:
                self.module._set_tool_icon(button, name, QStyle.SP_FileIcon, 24)
                image = button.icon().pixmap(24, 24).toImage()
                self.assertFalse(image.isNull(), name)
                pixels = tuple(image.pixel(x_pos, y_pos) for y_pos in range(image.height())
                               for x_pos in range(image.width()))
                self.assertTrue(any(image.pixelColor(x_pos, y_pos).alpha() for y_pos in range(image.height())
                                    for x_pos in range(image.width())), name)
                self.assertNotIn(pixels, images.values(), name)
                images[name] = pixels

    def test_native_program_icons_are_cached_and_tortoise_commit_is_distinct(self):
        """读取本机EXE图标但不启动程序；缺少安装时使用后备图标，提取成功则缓存并区分提交入口。"""
        from PyQt5.QtGui import QIcon, QPixmap, QColor
        from PyQt5.QtWidgets import QToolButton, QStyle
        pixmap = QPixmap(32, 32)
        pixmap.fill(QColor('#a72942'))
        native = QIcon(pixmap)
        button = QToolButton()
        self.addCleanup(button.deleteLater)
        with patch.dict(self.module._TOOL_NATIVE_CACHE, {}, clear=True), \
                patch.object(self.module, '_native_tool_executable', return_value='local-program.exe'), \
                patch.object(self.module.TitleShortcutBar, '_extract_icon_fast', return_value=native) as extract:
            self.module._set_tool_icon(button, 'app-tortoisegit-log', QStyle.SP_FileIcon, 32)
            plain = button.icon().pixmap(32, 32).toImage()
            self.module._set_tool_icon(button, 'app-tortoisegit-commit', QStyle.SP_FileIcon, 32)
            commit = button.icon().pixmap(32, 32).toImage()
            extract.assert_called_once_with('local-program.exe')
            self.assertEqual(plain.pixelColor(0, 0), commit.pixelColor(0, 0))
            self.assertNotEqual(plain, commit)
        for program in ('cmd', 'powershell', 'git-bash', 'tortoisegit'):
            with self.subTest(program=program):
                path = self.module._native_tool_executable(program)
                if path:
                    icon = self.module.TitleShortcutBar._extract_icon_fast(path)
                    self.assertIsNotNone(icon)
                    self.assertFalse(icon.isNull())

    def test_icon_assets_resolve_from_onefile_bundle_directory(self):
        """模拟单文件EXE资源目录，图标从包内icons加载而不是依赖源码目录；缺失时仍有后备。"""
        from PyQt5.QtWidgets import QToolButton, QStyle
        with tempfile.TemporaryDirectory() as directory:
            asset_dir = Path(directory) / 'icons'
            asset_dir.mkdir()
            shutil.copyfile(SOURCE.parent / 'icons' / 'bot.svg', asset_dir / 'bot.svg')
            with patch.object(self.module.sys, '_MEIPASS', directory, create=True), \
                    patch.dict(self.module._TOOL_ASSET_CACHE, {}, clear=True):
                icon = self.module._tool_asset_icon('bot')
                self.assertFalse(icon.pixmap(24, 24).isNull())
                self.assertIn(str(asset_dir / 'bot.svg'), self.module._TOOL_ASSET_CACHE)
                button = QToolButton()
                try:
                    with patch.object(self.module, '_tool_asset_icon', return_value=self.module.QIcon()):
                        self.module._set_tool_icon(button, 'preferences-system', QStyle.SP_FileDialogInfoView)
                        self.assertFalse(button.icon().isNull())
                finally:
                    button.deleteLater()

    def test_tab_list_filters_and_resolves_moved_tab_by_identity(self):
        """标签列表按路径筛选；标签移到另一组后仍按对象身份跳转，Enter 可执行。"""
        from PyQt5.QtTest import QTest
        from PyQt5.QtWidgets import QWidget, QTabWidget, QStackedWidget
        module = self.module

        class Host(QWidget):
            _refresh_tab_labels = module.MainWindow._refresh_tab_labels

        owner = Host()
        self.addCleanup(owner.deleteLater)
        groups = [(QTabWidget(owner), QStackedWidget(owner)) for _index in range(2)]
        owner.content_stack = groups[0][1]
        owner._all_groups = lambda: groups
        owner.set_active_pane_to_group = Mock()
        panes = []
        for path in (r'C:\alpha\src', r'C:\beta\src'):
            pane = QWidget()
            pane.current_path = path
            groups[0][1].addWidget(pane)
            groups[0][0].addTab(QWidget(), 'src')
            panes.append(pane)
        dialog = module.TabListDialog(owner)
        self.addCleanup(dialog.deleteLater)
        dialog.filter_input.setText('beta src')
        self.assertTrue(dialog.table.topLevelItem(0).isHidden())
        self.assertFalse(dialog.table.topLevelItem(1).isHidden())
        groups[0][1].removeWidget(panes[1])
        groups[0][0].removeTab(1)
        groups[1][1].addWidget(panes[1])
        groups[1][0].addTab(QWidget(), 'src')
        QTest.keyClick(dialog.filter_input, Qt.Key_Return)
        owner.set_active_pane_to_group.assert_called_once_with(groups[1][0])
        self.assertEqual(dialog.result(), QDialog.Accepted)

    def test_tab_close_button_keeps_geometry_on_hover(self):
        """关闭按钮使用图标并固定预留区域，悬停不改变标签矩形或按钮对象。"""
        from PyQt5.QtTest import QTest
        from PyQt5.QtCore import QPoint
        bar = self.module.CustomTabBar()
        self.addCleanup(bar.deleteLater)
        bar.addTab('alpha')
        bar.addTab('beta')
        bar.resize(400, 34)
        bar.show()
        self.app.processEvents()
        before = bar.tabRect(0)
        button = bar.tabButton(0, bar.RightSide)
        self.assertFalse(button.icon().isNull())
        QTest.mouseMove(bar, before.center())
        QTest.mouseMove(bar, QPoint(390, 15))
        self.assertEqual(bar.tabRect(0), before)
        self.assertIs(bar.tabButton(0, bar.RightSide), button)
        bar.close()

    def test_split_indicator_tracks_active_side_without_resizing(self):
        """切换左右操作侧只切换标签与地址栏强调；关闭分屏后恢复单侧样式，尺寸不变。"""
        from PyQt5.QtWidgets import QWidget, QTabWidget, QStackedWidget
        groups = []
        panes = []
        for _index in range(2):
            tabs, stack = QTabWidget(), QStackedWidget()
            self.addCleanup(tabs.deleteLater)
            self.addCleanup(stack.deleteLater)
            pane = QWidget()
            pane.path_bar = self.module.SimplePathBar(pane)
            pane.path_bar.resize(500, 30)
            stack.addWidget(pane)
            tabs.addTab(QWidget(), 'test')
            groups.append((tabs, stack))
            panes.append(pane)
        owner = types.SimpleNamespace(content_stack=groups[0][1], _split_active=True,
                                      _all_groups=lambda: groups, get_active_pane=lambda: panes[1])
        before = [pane.path_bar.size() for pane in panes]
        update = self.module.MainWindow._update_active_pane_indicator
        update(owner)
        self.assertFalse(groups[0][0].tabBar().property('activePane'))
        self.assertTrue(groups[1][0].tabBar().property('activePane'))
        self.assertEqual([pane.path_bar._pane_active for pane in panes], [False, True])
        for pane, color in zip(panes, ('#b7bcc4', '#2f6fdb')):
            image = pane.path_bar.grab().toImage()
            self.assertEqual(image.pixelColor(5, image.height() - 1).name(), color)
        owner.get_active_pane = lambda: panes[0]
        update(owner)
        self.assertEqual([pane.path_bar._pane_active for pane in panes], [True, False])
        self.assertEqual([pane.path_bar.size() for pane in panes], before)
        owner._split_active = False
        update(owner)
        self.assertFalse(panes[0].path_bar._split_indicator)


class PerformanceContractTests(unittest.TestCase):
    """性能约束：检查资源上限和调度行为，不用机器相关的毫秒耗时判断快慢。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        import TabEx
        cls.module = TabEx

    def test_progress_notifications_are_throttled(self):
        """密集文件进度通知被节流，跨过时间间隔后允许下一次通知。

        场景：模拟 1000 次通知发生在同一时刻，然后将时钟推进 0.2 秒。
        预期：总共只有两条 Qt 进度信号；用模拟时钟，不依赖实际执行速度。
        """
        worker = self.module.FileBatchOpWorker('copy', [])
        observed = []
        worker.progress.connect(lambda *values: observed.append(values))
        with patch.object(self.module.time, 'monotonic', return_value=100.0):
            for index in range(1000):
                worker._emit_progress(str(index))
        with patch.object(self.module.time, 'monotonic', return_value=100.2):
            worker._emit_progress('next')
        self.assertEqual(len(observed), 2)
        self.assertEqual(observed[-1][-1], 'next')

    def test_folder_checker_close_defers_running_thread_deletion(self):
        """关闭标签时目录检查仍阻塞：超时后不能删除仍运行的 QThread。

        模拟 wait 超时；预期请求停止，但不立即 deleteLater，也不保留标签父对象。
        避免直接销毁真实运行线程导致整个测试进程被 Qt 中止。
        """
        checker = Mock()
        checker.isRunning.return_value = True
        checker.wait.return_value = False
        self.addCleanup(self.module._retained_background_threads.discard, checker)
        tab = types.SimpleNamespace(folder_checker=checker)
        self.module.FileExplorerTab._release_folder_checker(tab, wait_ms=0)
        self.assertIsNone(tab.folder_checker)
        checker.stop.assert_called_once_with()
        checker.deleteLater.assert_not_called()
        checker.setParent.assert_called_once_with(None)

    def test_result_signal_does_not_end_folder_or_git_thread_lifetime(self):
        """目录/Git/PIDL 结果返回后线程仍可能运行，删除标签不能提前销毁工作线程。

        真实 QThread 发出结果后等待事件；删除父控件并处理延迟删除消息，
        再放行线程。预期结果与退出信号分离，线程最终退出并从保留集合移除。
        """
        from PyQt5 import sip
        from PyQt5.QtCore import QCoreApplication, QEvent
        from PyQt5.QtWidgets import QWidget
        for worker_class in (self.module.FolderSizeChecker, self.module.GitStatusWorker,
                     self.module._PidlResolver):
            with self.subTest(worker=worker_class.__name__):
                release = threading.Event()
                entered = threading.Event()
                observed = []
                owner = QWidget()
                if worker_class is self.module.FolderSizeChecker:
                    worker = worker_class('unused', owner)
                    payload = ('unused', 1, False)
                    result_signal = worker.completed
                elif worker_class is self.module.GitStatusWorker:
                    worker = worker_class('unused', 'unused', 'git', owner)
                    payload = ('unused', 'unused', 'clean')
                    result_signal = worker.completed
                else:
                    worker = worker_class('unused', 1, owner)
                    payload = ('unused', None, -1, 1)
                    result_signal = worker.resolved

                def run():
                    result_signal.emit(*payload)
                    entered.set()
                    release.wait(5)

                worker.run = run
                result_signal.connect(lambda *values: observed.append(values))
                worker.start()
                try:
                    self.assertTrue(entered.wait(3))
                    self.module._retain_thread_until_finished(worker)
                    owner.deleteLater()
                    QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
                    self.app.processEvents()
                    self.assertTrue(sip.isdeleted(owner))
                    self.assertFalse(sip.isdeleted(worker))
                    self.assertTrue(worker.isRunning())
                    self.assertEqual(observed, [payload])
                    self.assertIn(worker, self.module._retained_background_threads)
                finally:
                    release.set()
                    if not sip.isdeleted(worker):
                        self.assertTrue(worker.wait(5000))
                    self.app.processEvents()
                    if not sip.isdeleted(owner):
                        owner.deleteLater()
                self.assertNotIn(worker, self.module._retained_background_threads)
                QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
                self.assertTrue(sip.isdeleted(worker))

    def test_parallel_submission_has_bounded_backlog(self):
        """批量复制不会一次性提交全部文件，待处理任务量受并发上限约束。

        场景：用模拟执行器接收 1000 个任务；只有调用 wait 时才让任务完成。
        预期：并发数为 2 时，待完成任务最多 4 个，所有任务最终处理一次。
        """
        from concurrent.futures import Future
        pending = []
        submitted = []
        peak = []

        class Executor:
            def __init__(self, max_workers, **kwargs):
                self.max_workers = max_workers

            def __enter__(self):
                return self

            def __exit__(self, *args):
                return False

            def submit(self, callback, task):
                future = Future()
                pending.append((future, callback, task))
                submitted.append(task)
                peak.append(len(pending))
                return future

        def complete_one(futures, **kwargs):
            future, callback, task = pending.pop(0)
            future.set_result(callback(task))
            return {future}, set(futures) - {future}

        worker = self.module.FileBatchOpWorker('copy', [], max_workers=2)
        with patch('concurrent.futures.ThreadPoolExecutor', Executor), \
                patch('concurrent.futures.wait', complete_one):
            errors = worker._run_parallel_file_tasks(iter(range(1000)), lambda task: task, str)
        self.assertEqual(errors, [])
        self.assertEqual(submitted, list(range(1000)))
        self.assertLessEqual(max(peak), 4)
        self.assertEqual(worker.done_units, 1000)

    def test_blocked_search_producer_exits_after_cancel(self):
        """真正等待满队列的生产线程，在收到取消后能够退出。

        场景：填满队列，用事件确认后台线程已经进入满队列等待，再取消。
        预期：线程在防挂死超时内退出并报告取消，旧消息没有被替换。
        """
        cancelled = threading.Event()
        result_queue = self.module._SearchResultQueue(cancelled)
        result_queue.maxsize = 1
        result_queue.put('first')
        waiting = threading.Event()
        outcomes = []
        original_wait = result_queue.not_full.wait

        def observe_wait(timeout=None):
            waiting.set()
            return original_wait(timeout)

        def produce():
            try:
                result_queue.put('second')
            except self.module._SearchCancelled:
                outcomes.append('cancelled')

        producer = threading.Thread(target=produce, daemon=True)
        with patch.object(result_queue.not_full, 'wait', side_effect=observe_wait):
            producer.start()
            try:
                self.assertTrue(waiting.wait(3), '生产者没有进入满队列等待')
            finally:
                cancelled.set()
                producer.join(3)
        self.assertFalse(producer.is_alive())
        self.assertEqual(outcomes, ['cancelled'])
        self.assertEqual(result_queue.get_nowait(), 'first')

    def test_search_cache_lru_bounds_memory_entries(self):
        """缓存容量达到上限时，只淘汰最久没有访问的记录。

        场景：容量设为 2，写入 first、second，访问 first 后再写 third。
        预期：保留 first 和 third；缓存值与时间戳都不残留 second。
        """
        cache = self.module.SearchCache(max_size=2)
        cache.put('first', [1])
        cache.put('second', [2])
        self.assertEqual(cache.get('first'), [1])
        cache.put('third', [3])
        self.assertIsNone(cache.get('second'))
        self.assertEqual(set(cache.cache), {'first', 'third'})
        self.assertEqual(set(cache._timestamps), {'first', 'third'})
        cache.clear()
        self.assertEqual(len(cache.cache), 0)
        self.assertEqual(len(cache._timestamps), 0)

    def test_network_and_local_worker_limits(self):
        """文件任务并发遵循默认值、网络盘降并发和显式设置。

        场景：分别创建本地默认、UNC 默认及手动指定并发数的工作线程对象。
        预期：本地默认 2，UNC 默认 1，显式并发不超过任务数；不访问网络盘。
        """
        local = self.module.FileBatchOpWorker('copy', [], r'C:\target')
        network = self.module.FileBatchOpWorker('copy', [], r'\\server\share')
        configured = self.module.FileBatchOpWorker('copy', [], max_workers=4)
        for count, expected in ((1, 1), (50, 2), (None, 2)):
            with self.subTest(count=count):
                self.assertEqual(local._get_io_workers(count), expected)
        self.assertEqual(network._get_io_workers(50), 1)
        self.assertEqual(configured._get_io_workers(2), 2)
        self.assertEqual(configured._get_io_workers(None), 4)

    def test_inflight_snapshot_does_not_schedule_duplicate_scan(self):
        """上一次目录扫描未返回时，连续轮询不重复提交后台任务。

        场景：模拟可见本地标签已有扫描任务，连续调用 50 次轮询。
        预期：线程池未被获取，也没有同步调用 isdir 或 scandir。
        """
        tab = types.SimpleNamespace(
            current_path=r'C:\test', _refresh_active=True, _snapshot_inflight=True,
            _selection_guard_active=lambda: False, _is_slow_path=lambda path: False)
        with patch.object(self.module.QThreadPool, 'globalInstance') as pool, \
                patch.object(os, 'scandir') as scan, patch.object(os.path, 'isdir') as isdir:
            for _index in range(50):
                self.module.FileExplorerTab._poll_directory_changes(tab)
        pool.assert_not_called()
        scan.assert_not_called()
        isdir.assert_not_called()

    def test_background_tab_stops_refresh_timers(self):
        """切到后台的标签停止目录轮询与路径同步定时器。

        场景：模拟四个正在运行的定时器，将标签设置为不活跃。
        预期：四个定时器均被停止，文件监听更新为 None，不执行目录扫描。
        """
        timer_names = ('dir_poll_timer', '_keepalive_sync_timer', '_path_sync_timer', '_path_sync_stop_timer')
        timers = {name: Mock(isActive=Mock(return_value=True)) for name in timer_names}
        tab = types.SimpleNamespace(current_path=r'C:\test', _refresh_file_watch_paths=Mock(), **timers)
        self.module.FileExplorerTab.set_refresh_active(tab, False)
        self.assertFalse(tab._refresh_active)
        for timer in timers.values():
            timer.stop.assert_called_once_with()
        tab._refresh_file_watch_paths.assert_called_once_with(None)

    def test_stale_snapshot_never_refreshes_new_directory(self):
        """快速切换目录后，旧目录的后台结果不会刷新新目录。

        场景：当前目录已改为 new，收到 old 的快照，再收到 new 的相同和变化快照。
        预期：旧结果被丢弃，相同结果不刷新，只有当前目录变化时发起一次刷新。
        """
        tab = types.SimpleNamespace(current_path='new', _snapshot_inflight=True,
                                    _last_dir_snapshot=(1, 2), _request_refresh=Mock())
        callback = self.module.FileExplorerTab._on_dir_snapshot_ready
        callback(tab, 'old', (3, 4))
        self.assertEqual(tab._last_dir_snapshot, (1, 2))
        callback(tab, 'new', (1, 2))
        tab._request_refresh.assert_not_called()
        callback(tab, 'new', (3, 4))
        tab._request_refresh.assert_called_once_with(reason='poll')


class OptimizationTests(unittest.TestCase):
    """性能与交互优化：快捷键钩子、标签休眠、配置写盘、Explorer 监听和粘贴冲突，共 12 个用例。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        import TabEx
        cls.module = TabEx

    def shortcut_host(self, hotkeys=None, count=5):
        tabs = Mock()
        tabs.count.return_value = count
        host = Mock()
        host.config = {'hotkeys': hotkeys or {}}
        host.get_active_group_tabwidget.return_value = tabs
        host._select_tab_by_number = lambda number: self.module.MainWindow._select_tab_by_number(host, number)
        return host, tabs

    def run_shortcut(self, host, vk, ctrl=False, shift=False, alt=False):
        return self.module.MainWindow._run_shortcut(host, vk, ctrl, shift, alt)

    def test_ctrl_number_switches_tab_in_active_group(self):
        """Ctrl+数字切换活动标签组的标签，Ctrl+9 到最后一个。

        场景：活动组有 5 个标签，依次按 Ctrl+3、Ctrl+9、Ctrl+7。
        预期：切到索引 2 和 4；超出标签数的 Ctrl+7 不切换。
        """
        host, tabs = self.shortcut_host()
        self.assertTrue(self.run_shortcut(host, 0x33, ctrl=True))
        self.assertTrue(self.run_shortcut(host, 0x39, ctrl=True))
        self.run_shortcut(host, 0x37, ctrl=True)
        self.assertEqual([call.args[0] for call in tabs.setCurrentIndex.call_args_list], [2, 4])

    def test_ctrl_number_respects_explorer_view_keys_and_settings(self):
        """Ctrl+Shift+数字留给 Explorer 视图切换，关闭设置后 Ctrl+数字不生效。"""
        host, tabs = self.shortcut_host()
        self.assertFalse(self.run_shortcut(host, 0x33, ctrl=True, shift=True))
        disabled, disabled_tabs = self.shortcut_host({'switch_tab_number': False})
        self.assertFalse(self.run_shortcut(disabled, 0x33, ctrl=True))
        tabs.setCurrentIndex.assert_not_called()
        disabled_tabs.setCurrentIndex.assert_not_called()

    def test_shortcut_modifiers_must_match(self):
        """Alt 组合不被 AltGr(Ctrl+Alt) 触发，F5 带 Ctrl 时交给 Explorer。"""
        host, _tabs = self.shortcut_host()
        self.assertTrue(self.run_shortcut(host, 0x56, alt=True))
        host.quick_paste_to_current_directory.assert_called_once_with()
        self.assertFalse(self.run_shortcut(host, 0x51, ctrl=True, alt=True))
        host.quick_cancel_background_file_operation.assert_not_called()
        self.assertTrue(self.run_shortcut(host, 0x74))
        self.assertFalse(self.run_shortcut(host, 0x74, ctrl=True))
        host.refresh_current_tab.assert_called_once_with()

    def test_keyboard_hook_thread_installs_and_stops(self):
        """键盘钩子在独立线程安装，停止后线程退出。不模拟真实按键。"""
        hook = self.module._ShortcutKeyHook()
        self.assertTrue(hook.start())
        thread = hook._thread
        hook.stop()
        self.assertFalse(thread.is_alive())

    def test_config_flush_compares_with_last_written_content(self):
        """配置未变化时不再读盘或写盘，变化后原子写入。

        场景：首次写入后在外部改动文件，再用相同配置写盘，最后修改配置。
        预期：相同配置不覆盖外部内容；配置变化后文件更新。
        """
        with tempfile.TemporaryDirectory() as directory:
            host = types.SimpleNamespace(config={'a': 1})
            path = Path(directory) / 'config.json'
            with patch.object(self.module, 'get_app_data_path', side_effect=lambda *parts: os.path.join(directory, *parts)):
                flush = self.module.MainWindow._flush_config_to_disk
                flush(host)
                self.assertIn('"a": 1', path.read_text(encoding='utf-8'))
                path.write_text('external', encoding='utf-8')
                flush(host)
                self.assertEqual(path.read_text(encoding='utf-8'), 'external')
                host.config['a'] = 2
                flush(host)
                self.assertIn('"a": 2', path.read_text(encoding='utf-8'))

    def test_browser_hibernate_keeps_path_for_rebuild(self):
        """休眠销毁 Shell 视图并记录待导航路径；未初始化的视图不处理。"""
        browser = types.SimpleNamespace(_init_ok=True, cleanup=Mock(), _pending_path=None)
        self.assertTrue(self.module.IExplorerBrowserWidget.hibernate(browser, 'C:\\work'))
        browser.cleanup.assert_called_once_with()
        self.assertEqual(browser._pending_path, 'C:\\work')
        idle = types.SimpleNamespace(_init_ok=False, cleanup=Mock())
        self.assertFalse(self.module.IExplorerBrowserWidget.hibernate(idle, 'C:\\work'))
        idle.cleanup.assert_not_called()

    def test_tab_hibernate_skips_busy_tabs(self):
        """有后台文件任务的标签不休眠，空闲后台标签休眠并标记状态。"""
        explorer = Mock(spec=self.module.IExplorerBrowserWidget)
        explorer._init_ok = True
        explorer.hibernate.return_value = True
        worker = Mock()
        worker.isRunning.return_value = True
        tab = types.SimpleNamespace(explorer=explorer, current_path='C:\\work', _refresh_active=False,
                                    isVisible=lambda: False, _file_op_worker=worker)
        hibernate = self.module.FileExplorerTab.hibernate_shell_view
        self.assertFalse(hibernate(tab))
        worker.isRunning.return_value = False
        self.assertTrue(hibernate(tab))
        self.assertTrue(tab._hibernated)
        explorer.hibernate.assert_called_once_with('C:\\work')

    def test_only_idle_background_tabs_hibernate(self):
        """只释放超过闲置时间的非当前标签；设置为 0 时关闭。

        场景：模拟时钟 10000 秒，三个标签分别闲置 31 分钟、5 分钟，以及闲置但为当前标签。
        """
        def tab(idle_minutes):
            return types.SimpleNamespace(_last_active_at=10000 - idle_minutes * 60,
                                         hibernate_shell_view=Mock(return_value=True))
        idle, recent, current = tab(31), tab(5), tab(60)
        stack = Mock()
        stack.currentWidget.return_value = current
        stack.count.return_value = 3
        stack.widget.side_effect = [idle, recent, current].__getitem__
        host = Mock()
        host.config = {'tab_hibernate_minutes': 30}
        host._all_groups.return_value = [(None, stack)]
        with patch.object(self.module.time, 'monotonic', return_value=10000):
            self.assertEqual(self.module.MainWindow._hibernate_idle_tabs(host), 1)
            host.config['tab_hibernate_minutes'] = 0
            self.assertEqual(self.module.MainWindow._hibernate_idle_tabs(host), 0)
        idle.hibernate_shell_view.assert_called_once_with()
        recent.hibernate_shell_view.assert_not_called()
        current.hibernate_shell_view.assert_not_called()
        host._schedule_tab_labels.assert_called_once_with()

    def test_tab_tooltip_branch_matches_current_path(self):
        """标签提示只显示当前路径对应的 Git 分支，导航后旧分支不会残留。"""
        pane = types.SimpleNamespace(current_path='C:\\repo', _git_branch_info=('C:\\repo', 'main'))
        self.assertEqual(self.module._tab_git_branch(pane), 'main')
        pane.current_path = 'C:\\other'
        self.assertEqual(self.module._tab_git_branch(pane), '')

    def test_paste_conflict_skip_and_keep_both(self):
        """同名冲突可按项跳过或保留两者，跳过不覆盖且计入跳过数。

        场景：目标目录已有 a.txt 和 b.txt，来源 a 选跳过、b 选保留两者；同目录复制不算冲突。
        """
        with tempfile.TemporaryDirectory() as directory:
            source_dir, target_dir = Path(directory) / 'src', Path(directory) / 'dst'
            source_dir.mkdir()
            target_dir.mkdir()
            for name in ('a.txt', 'b.txt'):
                (source_dir / name).write_text('new', encoding='utf-8')
                (target_dir / name).write_text('old', encoding='utf-8')
            sources = [str(source_dir / 'a.txt'), str(source_dir / 'b.txt')]
            self.assertEqual(self.module._find_copy_conflicts(sources, str(target_dir)), sources)
            self.assertEqual(self.module._find_copy_conflicts(sources, str(source_dir)), [])
            worker = self.module.FileBatchOpWorker('copy', sources, str(target_dir))
            worker.conflict_actions = {sources[0]: 'skip', sources[1]: 'rename'}
            worker.run()
            self.assertEqual((worker.ok_count, worker.fail_count, worker.skipped_count), (1, 0, 1))
            self.assertEqual((target_dir / 'a.txt').read_text(encoding='utf-8'), 'old')
            self.assertEqual((target_dir / 'b.txt').read_text(encoding='utf-8'), 'old')
            self.assertEqual((target_dir / 'b - copy.txt').read_text(encoding='utf-8'), 'new')

    def test_conflict_dialog_applies_choice_to_all(self):
        """冲突对话框默认保留两者，“全部跳过”后所有项改为跳过。"""
        dialog = self.module.CopyConflictDialog(['C:\\a.txt', 'C:\\b.txt'], 'D:\\dst')
        self.addCleanup(dialog.deleteLater)
        self.assertEqual(set(dialog.actions().values()), {'rename'})
        dialog.set_all('skip')
        self.assertEqual(dialog.actions(), {'C:\\a.txt': 'skip', 'C:\\b.txt': 'skip'})

    def test_explorer_show_event_wait(self):
        """窗口显示事件到达时立即唤醒扫描，无事件时按超时返回。不打开真实 Explorer。"""
        host = types.SimpleNamespace(explorer_monitoring=True)
        window = self.module.MainWindow
        hook = window._create_explorer_show_hook(host)
        self.assertIsNotNone(hook)
        try:
            self.assertFalse(window._wait_explorer_show_event(host, hook, 0.05))
            hook['pending'].set()
            self.assertTrue(window._wait_explorer_show_event(host, hook, 5))
        finally:
            hook['user32'].UnhookWinEvent(hook['handle'])


class HotkeyAndUpdateTests(unittest.TestCase):
    """自定义快捷键与检查更新，共 10 个用例。网络请求全部模拟，不访问 GitHub。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        import TabEx
        cls.module = TabEx

    def test_hotkey_text_parsing_accepts_qt_names(self):
        """按键文本兼容 Qt 录制结果与手写配置，统一为规范写法；不支持的写法返回 None。"""
        module = self.module
        cases = {'Ctrl+Shift+Backtab': 'Ctrl+Shift+Tab', 'alt+delete': 'Alt+Del', 'Ctrl+Return': 'Ctrl+Enter',
                 'Shift+Ctrl+t': 'Ctrl+Shift+T', 'F12': 'F12', 'Ctrl+Alt+PageDown': 'Ctrl+Alt+PgDown'}
        for text, expected in cases.items():
            self.assertEqual(module._format_hotkey(*module._parse_hotkey(text)), expected, text)
        for text in ('', 'Ctrl++', 'Meta+T', 'Ctrl+Ctrl+T', 'Ctrl+ä', 'Num+5'):
            self.assertIsNone(module._parse_hotkey(text), text)

    def test_default_bindings_are_valid(self):
        """默认快捷键全部可解析且互不冲突，与改造前的按键一致。"""
        bindings = self.module._hotkey_bindings({})
        self.assertEqual(self.module._validate_hotkey_bindings(bindings), [])
        self.assertEqual(bindings['reopen_tab'], 'Ctrl+Shift+T')
        self.assertEqual(bindings['quick_delete'], 'Alt+Del')

    def test_validation_reports_unsafe_and_duplicate_bindings(self):
        """缺少修饰键、占用 Explorer 文件快捷键、与 Ctrl+数字冲突或重复分配时报错。"""
        validate = self.module._validate_hotkey_bindings
        bindings = self.module._hotkey_bindings({})
        cases = (('search', 'T'), ('search', 'Ctrl+C'), ('search', 'Ctrl+3'), ('search', 'Ctrl+W'), ('search', 'Ctrl+ä'))
        for command, text in cases:
            with self.subTest(text=text):
                self.assertEqual(len(validate(dict(bindings, **{command: text}))), 1)
        self.assertEqual(validate(dict(bindings, search='Ctrl+3'), number_switch_enabled=False), [])
        self.assertEqual(validate(dict(bindings, search='', close_tab='')), [])

    def test_remapped_hotkey_dispatches_new_combination_only(self):
        """改键后新组合触发原命令，旧组合失效；清空按键即停用。"""
        host = Mock()
        host.config = {'hotkeys': {}, 'hotkey_bindings': {'search': 'Ctrl+E', 'refresh': ''}}
        run = self.module.MainWindow._run_shortcut
        self.assertTrue(run(host, 0x45, True, False, False))
        self.assertFalse(run(host, 0x46, True, False, False))
        self.assertFalse(run(host, 0x74, False, False, False))
        host.show_search_dialog.assert_called_once_with()
        host.refresh_current_tab.assert_not_called()

    def test_shell_key_filter_follows_bindings(self):
        """交给 TabEx 的组合随绑定变化；关闭 Ctrl+数字后数字键交还 Explorer。"""
        combos = self.module._hotkey_reserved_combos(
            {'hotkey_bindings': {'search': 'Ctrl+E'}, 'hotkeys': {'switch_tab_number': False}})
        self.assertIn((True, False, 0x45), combos)
        self.assertNotIn((True, False, 0x46), combos)
        self.assertNotIn((True, False, 0x31), combos)
        self.assertIn((False, False, 0x74), combos)

    def test_settings_dialog_blocks_conflicts_and_restores_defaults(self):
        """设置窗口录入重复按键时拒绝保存，恢复默认后可保存；不写入实际配置。"""
        from PyQt5.QtGui import QKeySequence
        from PyQt5.QtWidgets import QCheckBox
        dialog = self.module.SettingsDialog({'hotkey_bindings': {'go_back': 'Alt+B'}})
        self.addCleanup(dialog.deleteLater)
        self.assertEqual(dialog.hotkey_edits['go_back'].keySequence().toString(QKeySequence.PortableText), 'Alt+B')
        forward_switch = next(box for box in dialog.findChildren(QCheckBox) if box.text() == self.module.tr('前进'))
        dialog.hotkey_navigate.setChecked(False)
        self.assertFalse(forward_switch.isChecked())
        dialog.hotkey_edits['search'].setKeySequence(QKeySequence('Ctrl+W'))
        bindings, errors = dialog._collect_hotkey_bindings()
        self.assertEqual(bindings['search'], 'Ctrl+W')
        self.assertEqual(len(errors), 1)
        with patch('PyQt5.QtWidgets.QMessageBox.warning') as warning:
            dialog.accept()
        warning.assert_called_once()
        self.assertNotEqual(dialog.result(), QDialog.Accepted)
        dialog._reset_hotkey_bindings()
        bindings, errors = dialog._collect_hotkey_bindings()
        self.assertEqual(errors, [])
        self.assertEqual(bindings, self.module._hotkey_bindings({}))

    def fake_urlopen(self, release, head):
        import io
        import json
        def opener(request, timeout=None):
            url = request.full_url
            if url == self.module.UPDATE_RELEASE_API:
                return io.BytesIO(json.dumps(release).encode('utf-8'))
            self.assertEqual(url, self.module.UPDATE_SOURCE_URL.format(tag=release['tag_name']))
            return io.BytesIO(head.encode('utf-8'))
        return opener

    def test_release_version_read_from_tagged_source(self):
        """发布标签与软件版本号不同，版本号从该标签源码开头读取；仓库外链接被替换。"""
        release = {'tag_name': 'v1.14.0', 'html_url': 'https://example.com/evil'}
        head = '# header\nAPP_VERSION = "3.80"\n'
        with patch('urllib.request.urlopen', side_effect=self.fake_urlopen(release, head)):
            result = self.module._fetch_latest_release()
        self.assertEqual(result, {'tag': 'v1.14.0', 'version': '3.80', 'url': self.module.UPDATE_RELEASES_PAGE})
        with patch('urllib.request.urlopen', side_effect=self.fake_urlopen(dict(release, tag_name='../x'), head)):
            with self.assertRaises(ValueError):
                self.module._fetch_latest_release()
        self.assertGreater(self.module._version_tuple('3.100'), self.module._version_tuple('v3.99'))

    def test_update_errors_are_readable(self):
        """GitHub 限流和超时转换为可读提示。"""
        import socket
        import urllib.error
        limited = urllib.error.HTTPError(self.module.UPDATE_RELEASE_API, 403, 'rate limited', {}, None)
        self.assertIn('频繁', self.module._update_error_text(limited))
        self.assertIn('超时', self.module._update_error_text(socket.timeout()))

    def update_host(self):
        host = types.SimpleNamespace(config={}, save_config=Mock(), _update_worker=object())
        return host

    def test_update_notice_once_per_version_for_auto_check(self):
        """自动检查发现新版本只提醒一次；手动检查总是提示结果。"""
        finished = self.module.MainWindow._on_update_check_finished
        host = self.update_host()
        newer = {'tag': 'v9', 'version': '99.0', 'url': self.module.UPDATE_RELEASES_PAGE}
        with patch.object(self.module, 'show_toast') as toast:
            finished(host, dict(newer, manual=False))
            finished(host, dict(newer, manual=False))
            self.assertEqual(toast.call_count, 1)
            self.assertEqual(toast.call_args.kwargs['action_text'], '查看更新')
            finished(host, dict(newer, manual=True))
            self.assertEqual(toast.call_count, 2)
            finished(host, {'tag': 'v1', 'version': '1.0', 'url': '', 'manual': True})
            self.assertIn('最新版本', toast.call_args.args[2])
        self.assertIsNone(host._update_worker)
        self.assertGreater(host.config['last_update_check'], 0)
        self.assertEqual(host.config['update_notified_version'], '99.0')

    def test_update_errors_silent_for_auto_check(self):
        """自动检查失败不打扰用户且不记录检查时间；手动检查失败给出警告。"""
        finished = self.module.MainWindow._on_update_check_finished
        host = self.update_host()
        with patch.object(self.module, 'show_toast') as toast:
            finished(host, {'error': 'offline', 'manual': False})
            toast.assert_not_called()
            finished(host, {'error': 'offline', 'manual': True})
            self.assertEqual(toast.call_args.kwargs['level'], 'warning')
        self.assertNotIn('last_update_check', host.config)


if __name__ == '__main__':
    unittest.main()
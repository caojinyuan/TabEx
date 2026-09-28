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


if __name__ == '__main__':
    unittest.main()
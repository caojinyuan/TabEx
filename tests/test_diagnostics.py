"""崩溃诊断包：临时目录内验证隐私、保留、容量和导出，不触发真实崩溃。"""

import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
import zipfile

from tabex_diagnostics import Diagnostics, MAX_LOG_BYTES, redact


class DiagnosticsTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.diagnostics = Diagnostics(self.root / 'diagnostics', 'test-version')
        self.addCleanup(self.diagnostics.close)

    def test_redacts_paths_urls_credentials_and_email(self):
        """导出脱敏覆盖本地路径、UNC、URL、常见凭据字段及邮箱。"""
        for value, secret in (
                (r'path="C:\Users\someone\private.txt"', 'private.txt'),
                (r'path="\\server\private\file"', 'server'),
                ('https://user:pass@example.com/?token=hidden', 'hidden'),
                ('api_key="my secret"', 'my secret'),
                ('Authorization: Bearer abc123', 'abc123'),
                ('mail=someone@example.com', 'someone@example.com')):
            with self.subTest(value=value):
                self.assertNotIn(secret, redact(value))

    def test_bundle_allowlist_excludes_config_chat_and_unrelated_logs(self):
        """ZIP 仅含白名单诊断文件，排除配置、聊天、任意文件和非诊断日志行。"""
        (self.diagnostics.session / 'config.json').write_text('SECRET-CONFIG')
        (self.diagnostics.session / 'chat_history.json').write_text('SECRET-CHAT')
        debug = self.root / 'debug.log'
        debug.write_text('[AI] SECRET-PROMPT [ClosedTabs]\n[ClosedTabs] path="C:\\private\\file"\n', encoding='utf-8')
        self.diagnostics.record('qt_fatal', 'QThread: Destroyed while thread is still running')
        self.diagnostics.snapshot({'tabs': 3, 'tasks': [{'operation': 'copy', 'done': False}]})
        bundle = self.root / 'bundle.zip'
        self.diagnostics.export(bundle, debug)
        with zipfile.ZipFile(bundle) as archive:
            combined = '\n'.join(archive.read(name).decode('utf-8') for name in archive.namelist())
            for secret in ('SECRET-CONFIG', 'SECRET-CHAT', 'SECRET-PROMPT', 'private'):
                self.assertNotIn(secret, combined)
            self.assertIn('QThread: Destroyed', combined)
            self.assertIn('manifest.json', archive.namelist())
            self.assertIsNone(archive.testzip())

    def test_previous_session_survives_restart_and_export(self):
        """新会话不会覆盖旧会话错误，导出仍包含上次异常退出记录。"""
        self.diagnostics.record('qt_fatal', 'test-crash')
        self.diagnostics.mark('fatal')
        later = Diagnostics(self.root / 'diagnostics', 'next-version')
        self.addCleanup(later.close)
        bundle = self.root / 'restart.zip'
        later.export(bundle)
        with zipfile.ZipFile(bundle) as archive:
            state = json.loads(archive.read(f'sessions/{self.diagnostics.session.name}/session.json'))
            self.assertEqual(state['status'], 'fatal')

    def test_exception_records_stack_without_message_or_source(self):
        """异常保留类型、函数和行号，但不记录可能含密钥的异常消息和源码行。"""
        try:
            raise ValueError('PRIVATE-TOKEN-123')
        except ValueError as error:
            self.diagnostics.exception('python_exception', type(error), error, error.__traceback__)
        content = (self.diagnostics.session / 'events.jsonl').read_text()
        self.assertIn('ValueError', content)
        self.assertNotIn('PRIVATE-TOKEN-123', content)
        self.assertEqual(self.diagnostics.state['status'], 'error')

    def test_log_rotation_is_bounded_and_export_missing_log_is_valid(self):
        """事件日志达到上限即轮转；调试日志缺失不阻止导出，清单记录缺失原因。"""
        events = self.diagnostics.session / 'events.jsonl'
        events.write_text('x' * (MAX_LOG_BYTES + 1))
        self.diagnostics.record('test', 'new')
        self.assertLess(events.stat().st_size, 1024)
        self.assertTrue((self.diagnostics.session / 'events.previous.jsonl').exists())
        bundle = self.root / 'missing.zip'
        self.diagnostics.export(bundle, self.root / 'absent.log')
        with zipfile.ZipFile(bundle) as archive:
            manifest = json.loads(archive.read('manifest.json'))
            self.assertIn('Debug log unavailable', manifest['warnings'])

    def test_retention_prunes_old_inactive_sessions_but_keeps_live_process(self):
        """保留最近五次运行，不删除仍活跃的会话；旧异常退出记录也受容量限制。"""
        for index in range(8):
            session = self.diagnostics.directory / f'20200101-00000{index}-12345678'
            session.mkdir()
            (session / 'session.json').write_text(json.dumps({'status': 'running', 'pid': 123}))
        with patch('tabex_diagnostics._process_alive', return_value=True):
            self.diagnostics._prune()
        self.assertEqual(len(list(self.diagnostics.directory.iterdir())), 9)
        with patch('tabex_diagnostics._process_alive', return_value=False):
            self.diagnostics._prune()
        self.assertEqual(len(list(self.diagnostics.directory.iterdir())), 5)
        self.assertTrue(self.diagnostics.session.exists())

    def test_export_failure_preserves_existing_zip_and_removes_temporary(self):
        """ZIP 提交失败时不损坏已有诊断包，临时文件会清理。"""
        target = self.root / 'bundle.zip'
        target.write_bytes(b'old-bundle')
        with patch('tabex_diagnostics.os.replace', side_effect=PermissionError('busy')):
            with self.assertRaises(PermissionError):
                self.diagnostics.export(target)
        self.assertEqual(target.read_bytes(), b'old-bundle')
        self.assertEqual(list(self.root.glob('.tabex-diagnostics-*')), [])

    def test_previous_debug_is_captured_before_original_is_overwritten(self):
        """启动前筛选旧调试日志，原文件被下一次启动覆盖后仍可导出上次致命错误。"""
        debug = self.root / 'debug.log'
        debug.write_text('Qt Message: QThread: Destroyed while thread is still running\n')
        self.diagnostics.capture_debug(debug)
        debug.write_text('new session')
        target = self.root / 'bundle.zip'
        self.diagnostics.export(target)
        with zipfile.ZipFile(target) as archive:
            name = f'sessions/{self.diagnostics.session.name}/previous-debug-filtered.log'
            self.assertIn(b'QThread: Destroyed', archive.read(name))

    def test_exception_hooks_chain_and_restore_without_recording_message(self):
        """主线程及 Python 线程异常钩子保留原处理器，退出时还原，不收集异常消息。"""
        import sys
        import threading
        from types import SimpleNamespace
        main_hook, thread_hook = Mock(), Mock()
        with patch.object(sys, 'excepthook', main_hook), \
                patch.object(threading, 'excepthook', thread_hook), \
                patch('tabex_diagnostics.faulthandler.is_enabled', return_value=True):
            self.diagnostics.install()
            try:
                error = ValueError('HIDDEN-EXCEPTION-MESSAGE')
                sys.excepthook(ValueError, error, None)
                threading.excepthook(SimpleNamespace(exc_type=ValueError, exc_value=error, exc_traceback=None))
                main_hook.assert_called_once()
                thread_hook.assert_called_once()
            finally:
                self.diagnostics.close()
            self.assertIs(sys.excepthook, main_hook)
            self.assertIs(threading.excepthook, thread_hook)
        content = (self.diagnostics.session / 'events.jsonl').read_text()
        self.assertNotIn('HIDDEN-EXCEPTION-MESSAGE', content)

    def test_normal_exit_has_clean_marker(self):
        """正常退出有 closed 及结束时间，区分没有清理标记的异常或强制退出。"""
        self.diagnostics.close()
        state = json.loads((self.diagnostics.session / 'session.json').read_text())
        self.assertEqual(state['status'], 'closed')
        self.assertIn('ended', state)

    def test_real_thread_dump_is_exported_with_paths_redacted(self):
        """实际写入当前 Python 线程栈并导出，不触发崩溃且不改变全局故障处理器。"""
        path = self.diagnostics.session / 'fault.log'
        with path.open('wb') as target:
            self.diagnostics.fault_file = target
            try:
                self.diagnostics.dump_threads()
            finally:
                self.diagnostics.fault_file = None
        self.assertGreater(path.stat().st_size, 0)
        bundle = self.root / 'stack.zip'
        self.diagnostics.export(bundle)
        with zipfile.ZipFile(bundle) as archive:
            content = archive.read(f'sessions/{self.diagnostics.session.name}/fault.log').decode('utf-8')
        self.assertIn('test_real_thread_dump', content)
        self.assertNotIn(str(Path(__file__).resolve().parent), content)

    def test_process_probe_recognizes_current_process(self):
        """进程存活检查识别当前进程及无效 PID，不终止任何进程。"""
        import os
        from tabex_diagnostics import _process_alive
        self.assertTrue(_process_alive(os.getpid()))
        self.assertFalse(_process_alive(-1))

    def test_external_fault_handler_is_preserved(self):
        """外部调试器已启用 faulthandler 时不覆盖或关闭它，仍提供 Qt 显式栈记录。"""
        with patch('tabex_diagnostics.faulthandler.is_enabled', return_value=True), \
                patch('tabex_diagnostics.faulthandler.enable') as enable, \
                patch('tabex_diagnostics.faulthandler.disable') as disable, \
                patch('tabex_diagnostics.faulthandler.dump_traceback') as dump:
            self.diagnostics.install()
            try:
                self.assertEqual(self.diagnostics.state['native_fault_capture'], 'external-handler')
                self.diagnostics.dump_threads()
                dump.assert_called_once_with(file=self.diagnostics.fault_file, all_threads=True)
            finally:
                self.diagnostics.close()
            enable.assert_not_called()
            disable.assert_not_called()


class QtDiagnosticsTests(unittest.TestCase):
    """真实 Qt 线程与导出动作，使用精简窗口代替原生 Explorer，不改用户配置。"""

    @classmethod
    def setUpClass(cls):
        from PyQt5.QtWidgets import QApplication
        import TabEx
        cls.app = QApplication.instance() or QApplication([])
        cls.module = TabEx

    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.diagnostics = Diagnostics(self.root / 'diagnostics', self.module.APP_VERSION)
        self.addCleanup(self.diagnostics.close)
        service_patch = patch.object(self.module, '_diagnostics', self.diagnostics)
        service_patch.start()
        self.addCleanup(service_patch.stop)

    def host(self):
        from PyQt5.QtWidgets import QWidget, QMenu
        module = self.module

        class ExportHost(QWidget):
            export_diagnostics = module.MainWindow.export_diagnostics
            _diagnostic_export_finished = module.MainWindow._diagnostic_export_finished
            _capture_diagnostic_snapshot = module.MainWindow._capture_diagnostic_snapshot
            _all_groups = lambda self: []

        host = ExportHost()
        menu = QMenu(host)
        host.export_diagnostics_action = menu.addAction('Export', host.export_diagnostics)

        def cleanup():
            worker = getattr(host, '_diagnostic_export_worker', None)
            if worker is not None:
                self.assertTrue(worker.wait(5000))
                self.app.processEvents()
            host.close()
            host.deleteLater()

        self.addCleanup(cleanup)
        return host

    def pump_until(self, predicate):
        import time
        deadline = time.monotonic() + 5
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
        self.assertTrue(predicate())

    def test_qt_fatal_is_recorded_with_debug_disabled(self):
        """关闭调试模式仍记录 Qt 致命错误和调用栈；不向 Qt 发送真正 abort。"""
        from PyQt5.QtCore import QtFatalMsg
        with patch.object(self.module, '_DEBUG_MODE', False), \
                patch.object(self.diagnostics, 'dump_threads') as dump:
            self.module.qt_message_handler(QtFatalMsg, None, 'QThread: Destroyed while thread is still running')
        dump.assert_called_once()
        self.assertEqual(self.diagnostics.state['status'], 'fatal')
        self.assertIn('QThread: Destroyed', (self.diagnostics.session / 'events.jsonl').read_text())

    def test_cancel_export_creates_no_zip(self):
        """用户拒绝隐私确认或取消保存对话框时，不启动导出也不创建 ZIP。"""
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        host = self.host()
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.No), \
                patch.object(QFileDialog, 'getSaveFileName') as choose:
            host.export_diagnostics_action.trigger()
        choose.assert_not_called()
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.Yes), \
                patch.object(QFileDialog, 'getSaveFileName', return_value=('', '')):
            host.export_diagnostics_action.trigger()
        self.assertIsNone(getattr(host, '_diagnostic_export_worker', None))
        self.assertEqual(list(self.root.glob('*.zip')), [])

    def test_confirm_export_uses_background_thread_and_restores_action(self):
        """确认菜单导出后，真实后台线程生成 ZIP，界面保持响应且完成后恢复按钮。

        用事件暂停后台导出，确定关闭保护生效；保存路径及确认对话框由测试提供。
        """
        import threading
        from PyQt5.QtGui import QCloseEvent
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        host = self.host()
        target = self.root / 'support.zip'
        host.add_new_tab = Mock()
        entered, release = threading.Event(), threading.Event()
        original = self.diagnostics.export

        def export(*args):
            entered.set()
            if not release.wait(5):
                raise RuntimeError('test gate timeout')
            return original(*args)

        with patch.object(QMessageBox, 'question', return_value=QMessageBox.Yes), \
                patch.object(QFileDialog, 'getSaveFileName', return_value=(str(target), '')), \
                patch.object(QMessageBox, 'information') as information, \
                patch.object(self.module, 'show_toast') as toast, \
                patch.object(QMessageBox, 'warning') as warning, \
                patch.object(self.diagnostics, 'export', side_effect=export), \
                patch.object(self.module, '_DEBUG_LOG_PATH', str(self.root / 'missing.log')):
            try:
                host.export_diagnostics_action.trigger()
                self.pump_until(entered.is_set)
                self.assertFalse(host.export_diagnostics_action.isEnabled())
                event = QCloseEvent()
                with patch.object(self.module, 'show_toast'):
                    self.module.MainWindow.closeEvent(host, event)
                self.assertFalse(event.isAccepted())
            finally:
                release.set()
                self.pump_until(lambda: getattr(host, '_diagnostic_export_worker', None) is None)
            self.assertTrue(host.export_diagnostics_action.isEnabled())
            information.assert_not_called()
            toast.assert_called_once()
            self.assertEqual(toast.call_args.kwargs['action_text'], self.module.tr('打开位置'))
            toast.call_args.kwargs['action']()
            host.add_new_tab.assert_called_once_with(str(self.root))
            warning.assert_not_called()
        with zipfile.ZipFile(target) as archive:
            self.assertIsNone(archive.testzip())
            snapshot = json.loads(archive.read(f'sessions/{self.diagnostics.session.name}/snapshot.json'))
            self.assertIn('python_threads', snapshot)
            self.assertIn('qt', snapshot)

    def test_export_failure_is_reported_and_can_retry(self):
        """后台写入失败向用户报错，恢复导出动作，不遗留运行中的线程引用。"""
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        host = self.host()
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.Yes), \
                patch.object(QFileDialog, 'getSaveFileName', return_value=(str(self.root / 'out.zip'), '')), \
                patch.object(QMessageBox, 'warning') as warning, \
                patch.object(self.diagnostics, 'export', side_effect=PermissionError('test denied')):
            host.export_diagnostics_action.trigger()
            self.pump_until(lambda: host._diagnostic_export_worker is None)
            self.assertTrue(host.export_diagnostics_action.isEnabled())
            warning.assert_called_once()
            self.assertIn('PermissionError', warning.call_args.args[2])

    def test_snapshot_excludes_paths_and_error_details(self):
        """任务快照只记录计数、类别、进度，不收集源路径、目标路径和错误消息。"""
        from types import SimpleNamespace
        host = self.host()
        host._file_task_panel = SimpleNamespace(records=[{
            'op': 'copy', 'paths': [r'C:\private\secret.txt'], 'destination': r'C:\private',
            'errors': ['SECRET-CONTENT'], 'failed': [r'C:\private\secret.txt'],
            'done': True, 'worker': None,
        }])
        host._capture_diagnostic_snapshot()
        text = (self.diagnostics.session / 'snapshot.json').read_text()
        self.assertNotIn('private', text)
        self.assertNotIn('SECRET-CONTENT', text)
        self.assertEqual(json.loads(text)['tasks'][0]['failures'], 1)

    def test_startup_falls_back_when_primary_directory_is_read_only(self):
        """首选目录无权限时回退临时目录；两处都失败时不阻止软件启动。"""
        service = Mock()
        with patch('tabex_diagnostics.Diagnostics', side_effect=[PermissionError('denied'), service]) as factory, \
                patch.dict('os.environ', {'LOCALAPPDATA': str(self.root / 'readonly')}), \
                patch('tempfile.gettempdir', return_value=str(self.root / 'fallback')):
            self.module._start_diagnostics()
        self.assertEqual(factory.call_count, 2)
        self.assertEqual(Path(factory.call_args.args[0]), self.root / 'fallback' / 'TabEx' / 'diagnostics')
        service.install.assert_called_once()
        service.capture_debug.assert_called_once()
        with patch.object(self.module, '_diagnostics', None), \
                patch('tabex_diagnostics.Diagnostics', side_effect=PermissionError('denied')):
            self.module._start_diagnostics()
            self.assertIsNone(self.module._diagnostics)

    def test_packaged_mode_uses_executable_directory_and_records_mode(self):
        """模拟 frozen 标记确认 EXE 使用稳定日志目录，诊断注明打包模式；不代替实机打包测试。"""
        with patch.object(self.module.sys, 'frozen', True, create=True), \
                patch.object(self.module.sys, 'executable', str(self.root / 'TabExplorer.exe')):
            self.assertEqual(Path(self.module.get_app_data_path('TabEx_debug_latest.log')),
                             self.root / 'TabEx_debug_latest.log')
            packaged = Diagnostics(self.root / 'packaged', self.module.APP_VERSION)
        self.addCleanup(packaged.close)
        self.assertTrue(packaged.state['frozen'])
        target = self.root / 'packaged.zip'
        packaged.export(target)
        with zipfile.ZipFile(target) as archive:
            state = json.loads(archive.read(f'sessions/{packaged.session.name}/session.json'))
            self.assertTrue(state['frozen'])


if __name__ == '__main__':
    unittest.main()
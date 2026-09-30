"""Service contracts, failure recovery and isolated startup behavior."""

from contextlib import ExitStack
import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
import types
import unittest
from unittest.mock import Mock, patch

from PyQt5.QtWidgets import QApplication
from tabexplorer.i18n import _set_app_language
from tabexplorer.workers import QuickFindWorker


class StartupTests(unittest.TestCase):
    def test_legacy_console_encoding_cannot_fail_on_status_symbols(self):
        """GBK 标准流在启动时配置为 UTF-8，状态符号不会产生未捕获异常。"""
        import io
        from tabexplorer.app import _configure_standard_streams
        output = io.BytesIO()
        stream = io.TextIOWrapper(output, encoding='gbk')
        try:
            with patch.object(sys, 'stdout', stream), patch.object(sys, 'stderr', None):
                _configure_standard_streams()
                print('\u2713 \u2717')
                stream.flush()
            self.assertIn('\u2713 \u2717', output.getvalue().decode('utf-8'))
        finally:
            stream.close()


class InstanceTests(unittest.TestCase):
    def setUp(self):
        from tabexplorer.instance import InstanceCoordinator
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.owner = InstanceCoordinator(temporary.name, temporary.name)
        self.client = InstanceCoordinator(temporary.name, temporary.name)
        self.addCleanup(self.owner.close)

    def test_no_argument_activation_and_duplicate_request_are_acknowledged_once(self):
        """无参数启动发送激活消息，同一请求重发不会重复打开标签。"""
        self.assertTrue(self.owner.acquire())
        self.assertFalse(self.client.acquire())
        received = []
        self.owner.set_receiver(received.append)
        self.assertTrue(self.client.send('', request_id='a' * 32))
        self.assertTrue(self.client.send('', request_id='a' * 32))
        self.assertEqual(received, [''])

    def test_pending_requests_and_long_unicode_paths_are_preserved(self):
        """启动窗口前收到的长中文路径会完整缓存，绑定界面后交付。"""
        self.assertTrue(self.owner.acquire())
        path = 'C:\\' + '\u6587\u4ef6\\' * 1800
        self.assertTrue(self.client.send(path))
        received = []
        self.owner.set_receiver(received.append)
        self.assertEqual(received, [path])

    def test_fragmented_messages_and_incomplete_connections(self):
        """TCP 分段仍完整解码，异常客户端不阻塞后续请求。"""
        import socket
        import struct
        from tabexplorer.instance import _receive_message
        payload = json.dumps({'path': '\u6587\u4ef6'}).encode('utf-8')
        fragments = iter(bytes([value]) for value in struct.pack('!I', len(payload)) + payload)
        connection = Mock(recv=Mock(side_effect=lambda size: next(fragments)))
        self.assertEqual(_receive_message(connection), {'path': '\u6587\u4ef6'})
        self.assertTrue(self.owner.acquire())
        endpoint = json.loads(self.owner.endpoint.read_text(encoding='utf-8'))
        with socket.create_connection(('127.0.0.1', endpoint['port'])) as stalled:
            stalled.sendall(b'\x00')
            self.assertTrue(self.client.send('next', timeout=2))

    def test_lock_release_and_cross_process_exclusion(self):
        """真实第二个进程不能取得写入权；持有者退出后锁可重新取得。"""
        import subprocess
        self.assertTrue(self.owner.acquire())
        script = ('import sys; from tabexplorer.instance import InstanceCoordinator; '
                  'owner=InstanceCoordinator(sys.argv[1], sys.argv[1]); '
                  'sys.exit(1 if owner.acquire() else 0)')
        result = subprocess.run([sys.executable, '-c', script, str(self.owner.endpoint.parent)],
                                timeout=10, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.owner.close()
        self.assertTrue(self.client.acquire())
        self.client.close()

    def test_crashed_owner_lock_is_recovered(self):
        """持锁进程不执行清理就退出时，下次启动能回收残留锁与端点。"""
        import subprocess
        script = ('import os, sys; from tabexplorer.instance import InstanceCoordinator; '
                  'owner=InstanceCoordinator(sys.argv[1], sys.argv[1]); '
                  'os._exit(0 if owner.acquire() else 1)')
        result = subprocess.run([sys.executable, '-c', script, str(self.owner.endpoint.parent)],
                                timeout=10, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(self.owner.acquire())

    def test_invalid_token_and_oversized_frame_do_not_reach_window(self):
        """无效认证和超长消息不产生窗口动作，后续有效请求仍可处理。"""
        import secrets
        import socket
        import struct
        from tabexplorer.instance import MAX_MESSAGE_BYTES, _send_message
        self.assertTrue(self.owner.acquire())
        received = []
        self.owner.set_receiver(received.append)
        endpoint = json.loads(self.owner.endpoint.read_text(encoding='utf-8'))
        with socket.create_connection(('127.0.0.1', endpoint['port'])) as connection:
            _send_message(connection, {'version': 1, 'id': secrets.token_hex(16), 'path': '', 'token': 'invalid'})
        with socket.create_connection(('127.0.0.1', endpoint['port'])) as connection:
            connection.sendall(struct.pack('!I', MAX_MESSAGE_BYTES + 1))
        self.assertTrue(self.client.send('valid'))
        self.assertEqual(received, ['valid'])


class AsyncWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_quick_find_stops_enumerating_after_result_limit(self):
        """200 个匹配后停止枚举，不读取或排序整个大目录。"""
        worker = QuickFindWorker('unused', 'match')
        consumed, results = [], []

        def entries():
            for index in range(10000):
                consumed.append(index)
                yield types.SimpleNamespace(name=f'match-{index}', path=f'match-{index}')

        iterator = Mock(__enter__=Mock(return_value=entries()), __exit__=Mock(return_value=False))
        worker.completed.connect(lambda *values: results.append(values))
        with patch('os.scandir', return_value=iterator):
            worker.run()
        self.assertEqual(len(consumed), 200)
        self.assertEqual(len(results[0][1]), 200)

    def test_quick_find_cancellation_discards_results(self):
        """目录扫描取消后不把部分结果交给界面。"""
        worker = QuickFindWorker('unused', 'match')
        results = []
        worker.completed.connect(lambda *values: results.append(values))
        with patch('os.scandir', side_effect=OSError('offline')):
            worker.run()
        self.assertEqual(results[0][2], 'offline')
        entered, release = threading.Event(), threading.Event()

        def blocked_scan(path):
            entered.set()
            release.wait(3)
            raise OSError('offline')

        results.clear()
        with patch('os.scandir', side_effect=blocked_scan):
            worker.start()
            try:
                self.assertTrue(entered.wait(2))
                worker.requestInterruption()
            finally:
                release.set()
                self.assertTrue(worker.wait(3000))
        self.app.processEvents()
        self.assertEqual(results, [])

    def test_opening_a_file_resolves_parent_in_worker(self):
        """命令行文件路径在后台解析为所在目录，保留原启动行为。"""
        from tabexplorer.workers import OpenPathWorker
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'note.txt'
            path.write_text('data', encoding='utf-8')
            worker = OpenPathWorker(str(path))
            results = []
            worker.completed.connect(lambda *values: results.append(values))
            worker.run()
            self.assertEqual(results, [(directory, '')])

    def test_data_directory_override_does_not_change_asset_directory(self):
        """验证数据目录与程序资源目录分离，不修改默认配置或影响图标定位。"""
        from tabexplorer.paths import get_app_base_dir, get_app_data_path
        base = get_app_base_dir()
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, TABEX_DATA_DIR=directory):
            self.assertEqual(get_app_data_path('config.json'), os.path.join(directory, 'config.json'))
            self.assertEqual(get_app_base_dir(), base)


class BuildPipelineTests(unittest.TestCase):
    def test_failed_build_or_smoke_never_replaces_working_executable(self):
        """打包、原生验证或原子替换失败都保留已有可用 EXE。"""
        from tools.build_release import build
        import subprocess
        for failure in ('build', 'validation', 'replace'):
            with self.subTest(failure=failure), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                executable = root / 'TabExplorer.exe'
                executable.write_bytes(b'old working executable')

                def run(command, **keywords):
                    if 'PyInstaller' in command:
                        if failure == 'build':
                            raise subprocess.CalledProcessError(1, command)
                        destination = Path(command[command.index('--distpath') + 1])
                        destination.mkdir()
                        (destination / 'TabExplorer.exe').write_bytes(b'candidate')

                verify = Mock(return_value={'ok': True, 'frozen': True})
                if failure == 'validation':
                    verify.side_effect = RuntimeError('native failed')
                with ExitStack() as stack:
                    if failure == 'replace':
                        stack.enter_context(patch('os.replace', side_effect=PermissionError('in use')))
                    with self.assertRaises((RuntimeError, OSError, subprocess.CalledProcessError)):
                        build(root, install=False, run=run, verify=verify)
                self.assertEqual(executable.read_bytes(), b'old working executable')
                self.assertEqual(list(root.glob('.tabex-build-*')), [])

    def test_only_validated_frozen_build_is_promoted(self):
        """测试与打包先执行，原生 EXE 验证成功后才替换正式文件。"""
        from tools.build_release import build
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            executable = root / 'TabExplorer.exe'
            executable.write_bytes(b'old')
            commands = []

            def run(command, **keywords):
                commands.append(command)
                if 'PyInstaller' in command:
                    destination = Path(command[command.index('--distpath') + 1])
                    destination.mkdir()
                    (destination / 'TabExplorer.exe').write_bytes(b'new')

            def verify(*arguments, **keywords):
                self.assertEqual(executable.read_bytes(), b'old')
                return {'ok': True, 'frozen': True}

            self.assertEqual(build(root, install=False, run=run, verify=verify), executable)
            self.assertIn('unittest', commands[0])
            self.assertEqual(executable.read_bytes(), b'new')


class ConfigStoreTests(unittest.TestCase):
    def test_invalid_field_types_do_not_discard_unrelated_settings(self):
        """配置字段类型异常只回退该字段，不丢失其他有效设置。"""
        from tabexplorer.persistence import ConfigStore
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.json'
            path.write_text(json.dumps({'hotkeys': [], 'cached_tabs': None, 'language': 'en',
                                        'ai_chat': {'panel_width': 'bad', 'model': 'local'}}), encoding='utf-8')
            config = ConfigStore(path).load()
            self.assertIsInstance(config['hotkeys'], dict)
            self.assertEqual(config['cached_tabs'], [])
            self.assertEqual(config['ai_chat']['panel_width'], 360)
            self.assertEqual(config['ai_chat']['model'], 'local')
            self.assertEqual(config['language'], 'en')
        _set_app_language('zh')

    def test_replace_failure_keeps_old_config_and_can_be_retried(self):
        """替换失败后旧文件完整、临时文件清理、内存缓存未误标为成功。"""
        from tabexplorer.persistence import ConfigStore
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.json'
            path.write_text('{"old": true}', encoding='utf-8')
            store = ConfigStore(path)
            with patch('os.replace', side_effect=PermissionError('busy')):
                self.assertFalse(store.save({'new': True}))
            self.assertEqual(json.loads(path.read_text(encoding='utf-8')), {'old': True})
            self.assertIsNone(store.last_written)
            self.assertFalse(Path(str(path) + '.tmp').exists())
            self.assertTrue(store.save({'new': True}))
            self.assertEqual(json.loads(path.read_text(encoding='utf-8')), {'new': True})

    def test_unpreserved_corrupt_config_is_never_overwritten(self):
        """损坏配置不能备份时禁止默认值写回覆盖原文件。"""
        from tabexplorer.persistence import ConfigStore
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.json'
            path.write_text('{broken', encoding='utf-8')
            store = ConfigStore(path)
            with patch('tabexplorer.persistence._preserve_unreadable_file', return_value=None):
                config = store.load()
            self.assertFalse(store.save(config))
            self.assertEqual(path.read_text(encoding='utf-8'), '{broken')

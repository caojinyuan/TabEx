"""Service contracts, failure recovery and isolated startup behavior."""

from contextlib import ExitStack
import json
import os
from pathlib import Path
import shutil
import subprocess
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


@unittest.skipUnless(os.name == 'nt', 'Project Python bootstrap requires Windows')
class ProjectPythonTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).resolve().parents[1]
        self.script = self.root / 'tools' / 'project_python.ps1'
        self.powershell = str(Path(os.environ['SystemRoot']) / 'System32' /
                              'WindowsPowerShell' / 'v1.0' / 'powershell.exe')

    def ensure(self, script, environment=None):
        return subprocess.run(
            [self.powershell, '-NoProfile', '-ExecutionPolicy', 'Bypass',
             '-File', str(script), '-Action', 'Ensure'],
            env=environment, capture_output=True, text=True, timeout=60)

    def test_batch_entrypoints_use_the_shared_project_runtime(self):
        for filename, action in (('0_install_requirements.bat', 'Install'),
                                 ('1_run_TabEx.bat', 'Run'), ('2_build_exe.bat', 'Build')):
            with self.subTest(filename=filename):
                content = (self.root / filename).read_text(encoding='utf-8')
                self.assertIn('tools\\project_python.ps1" -Action ' + action, content)
                self.assertNotIn('.venv\\Scripts', content)
                self.assertNotIn('set "PY=python"', content)

    def test_project_python_ignores_system_python_and_poisoned_environment(self):
        python = self.root / '.venv' / 'python-3.13.13' / 'python.exe'
        if not python.is_file():
            self.skipTest('Project runtime has not been bootstrapped')
        environment = dict(os.environ, PATH=os.environ['SystemRoot'], PYTHONHOME='Z:\\missing-python',
                           PYTHONPATH='Z:\\missing-packages', PYTHONUSERBASE='Z:\\missing-user',
                           PIP_TARGET='Z:\\missing-target', PIP_REQUIRE_VIRTUALENV='true')
        result = self.ensure(self.script, environment)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn(str(python), result.stdout)
        self.assertIn('Version: 3.13.13', result.stdout)
        self.assertFalse((python.parent / 'tools').exists())
        self.assertTrue((python.parent / 'pythonw.exe').is_file())
        self.assertTrue((python.parent / 'Lib').is_dir())

    def test_corrupt_cached_package_is_rejected_without_system_fallback(self):
        with tempfile.TemporaryDirectory(prefix='TabEx setup ') as directory:
            root = Path(directory)
            tools = root / 'tools'
            tools.mkdir()
            script = tools / self.script.name
            shutil.copy2(self.script, script)
            cache = root / '.venv' / 'cache'
            cache.mkdir(parents=True)
            (cache / 'python-3.13.13.zip').write_bytes(b'not a trusted Python package')
            result = self.ensure(script)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('Cached Python package is invalid', result.stdout)
            self.assertFalse((root / '.venv' / 'python-3.13.13').exists())

    def test_legacy_virtual_environment_is_preserved_outside_runtime_root(self):
        with tempfile.TemporaryDirectory(prefix='TabEx legacy setup ') as directory:
            root = Path(directory)
            tools = root / 'tools'
            tools.mkdir()
            script = tools / self.script.name
            shutil.copy2(self.script, script)
            legacy = root / '.venv'
            cache = legacy / 'cache'
            cache.mkdir(parents=True)
            configuration = 'home = Z:\\old-python\n'
            (legacy / 'pyvenv.cfg').write_text(configuration, encoding='utf-8')
            (cache / 'python-3.13.13.zip').write_bytes(b'not a trusted Python package')
            result = self.ensure(script)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn('Cached Python package is invalid', result.stdout)
            backups = list(root.glob('.venv-legacy-*'))
            self.assertEqual(len(backups), 1)
            self.assertEqual((backups[0] / 'pyvenv.cfg').read_text(encoding='utf-8'), configuration)
            self.assertFalse((legacy / 'pyvenv.cfg').exists())


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
                for tool in ('es.exe', 'rg.exe', '7zr.exe'):
                    (root / tool).write_bytes(b'tool')

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
            for tool in ('es.exe', 'rg.exe', '7zr.exe'):
                (root / tool).write_bytes(b'tool')
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


class QuickLookTests(unittest.TestCase):
    """QuickLook 集成：焦点判断、断点续传、校验与解压；用 7zr 现场生成的小压缩包代替真实下载。"""

    def setUp(self):
        from tabexplorer import quicklook
        self.quicklook = quicklook
        self.root = tempfile.TemporaryDirectory()
        self.addCleanup(self.root.cleanup)
        environment = patch.dict(os.environ, {'TABEX_DATA_DIR': self.root.name})
        environment.start()
        self.addCleanup(environment.stop)
        source = Path(self.root.name) / 'source'
        source.mkdir()
        (source / 'QuickLook.exe').write_bytes(b'fake quicklook')
        archive = Path(self.root.name) / 'fixture.7z'
        seven_zip = Path(__file__).resolve().parents[1] / '7zr.exe'
        subprocess.run([str(seven_zip), 'a', str(archive), str(source / '*'), '-bso0', '-bsp0'], check=True)
        self.payload = archive.read_bytes()
        import hashlib
        self.digest = hashlib.sha256(self.payload).hexdigest()

    def opener(self, payload=None, record=None):
        payload = self.payload if payload is None else payload

        def open_url(request, timeout=None):
            start = int(request.get_header('Range', 'bytes=0-').split('=')[1].rstrip('-'))
            if record is not None:
                record.append(start)
            import io
            response = io.BytesIO(payload[start:])
            response.status = 206 if start else 200
            return response
        return open_url

    def install(self, **keywords):
        return self.quicklook.install_quicklook(
            url='https://example.invalid/QuickLook.7z', sha256=self.digest, size=len(self.payload),
            opener=keywords.pop('opener', self.opener()), **keywords)

    def test_default_is_off_and_not_found_until_downloaded(self):
        """没有下载时找不到 QuickLook；配置默认关闭该功能。"""
        from tabexplorer.persistence import ConfigStore
        with patch.dict(os.environ, {'LOCALAPPDATA': '', 'ProgramFiles': '', 'ProgramFiles(x86)': ''}):
            self.assertIsNone(self.quicklook.find_quicklook())
        store = ConfigStore(Path(self.root.name) / 'config.json')
        self.assertIs(store.load()['quicklook_enabled'], False)

    def test_install_extracts_verified_archive(self):
        """校验通过后解压到数据目录，并清理压缩包和临时目录。"""
        executable = self.install()
        self.assertEqual(Path(executable).read_bytes(), b'fake quicklook')
        self.assertEqual(self.quicklook.find_quicklook(), executable)
        self.assertEqual(sorted(path.name for path in Path(self.root.name).glob('QuickLook*')), ['QuickLook'])

    def test_checksum_mismatch_discards_download(self):
        """内容与固定哈希不符时拒绝安装并删除已下载文件。"""
        with self.assertRaises(self.quicklook.QuickLookInstallError):
            self.install(opener=self.opener(self.payload + b'tampered'))
        self.assertFalse((Path(self.root.name) / 'QuickLook.7z.part').exists())
        self.assertFalse((Path(self.root.name) / 'QuickLook').exists())

    def test_cancelled_download_keeps_partial_file(self):
        """取消下载抛出取消异常，保留 .part 供下次继续，不产生安装目录。"""
        with self.assertRaises(self.quicklook.QuickLookCancelled):
            self.install(cancelled=lambda: True)
        self.assertFalse((Path(self.root.name) / 'QuickLook').exists())
        Path(self.root.name, 'QuickLook.7z.part').write_bytes(self.payload[:10])
        with self.assertRaises(self.quicklook.QuickLookCancelled):
            self.install(cancelled=lambda: True)
        self.assertEqual(Path(self.root.name, 'QuickLook.7z.part').read_bytes(), self.payload[:10])

    def test_download_resumes_from_partial_file(self):
        """已有部分下载时发送 Range 请求，只补齐剩余字节，最终哈希校验通过。"""
        half = len(self.payload) // 2
        Path(self.root.name, 'QuickLook.7z.part').write_bytes(self.payload[:half])
        calls = []
        self.install(opener=self.opener(record=calls))
        self.assertEqual(calls, [half])
        self.assertTrue((Path(self.root.name) / 'QuickLook' / 'QuickLook.exe').is_file())

    def test_space_only_previews_in_file_list(self):
        """只有文件列表控件持有焦点才预览；重命名框、目录树、Qt 控件和刚输入过字符时不触发。"""
        for name in ('DirectUIHWND', 'SysListView32', 'SHELLDLL_DefView'):
            self.assertTrue(self.quicklook.is_file_view_focus_class(name), name)
        for name in ('', 'Edit', 'ATL:Edit', 'SysTreeView32', 'Qt5152QWindowIcon', 'ComboBox'):
            self.assertFalse(self.quicklook.is_file_view_focus_class(name), name)
        self.assertTrue(self.quicklook.is_typing_key(0x41, False, False))
        self.assertTrue(self.quicklook.is_typing_key(0x31, False, False))
        for vk, ctrl, alt in ((0x41, True, False), (0x41, False, True), (0x28, False, False), (0x20, False, False)):
            self.assertFalse(self.quicklook.is_typing_key(vk, ctrl, alt))

    def test_only_first_launched_instance_is_stopped_on_exit(self):
        """本程序启动的 QuickLook 退出时结束；之后只转交路径的进程不会替换它。"""
        first, forwarder = Mock(), Mock()
        first.poll.return_value = None
        forwarder.poll.return_value = 0
        self.quicklook._owned_process = None
        with patch('subprocess.Popen', side_effect=[first, forwarder]):
            self.assertTrue(self.quicklook.preview(r'C:\a.txt', r'C:\ql\QuickLook.exe'))
            self.assertTrue(self.quicklook.preview(r'C:\b.txt', r'C:\ql\QuickLook.exe'))
        self.quicklook.stop_owned_process()
        first.terminate.assert_called_once()
        forwarder.terminate.assert_not_called()


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

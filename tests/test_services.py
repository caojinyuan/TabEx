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


class GitBashLaunchTests(unittest.TestCase):
    def test_launch_uses_current_directory_with_spaces_and_unicode(self):
        from tabexplorer.system import launch_shell_tool
        with tempfile.TemporaryDirectory(prefix='TabEx Git Bash ') as temporary:
            current_dir = Path(temporary) / '\u5f53\u524d\u76ee\u5f55 with spaces'
            current_dir.mkdir()
            git_root = Path(temporary) / 'Git'
            git_root.mkdir()
            executable = git_root / 'git-bash.exe'
            executable.touch()
            with patch('tabexplorer.system.find_git_install_root', return_value=str(git_root)), \
                    patch('tabexplorer.system.launch_detached_async') as launch:
                launch_shell_tool('git-bash', str(current_dir))
            launch.assert_called_once_with(
                [str(executable), f'--cd={current_dir}'], cwd=str(current_dir))


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
        from tabexplorer.paths import get_app_base_dir, get_app_config_path, get_app_data_path
        base = get_app_base_dir()
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, TABEX_DATA_DIR=directory):
            self.assertEqual(get_app_config_path('config.json'), os.path.join(directory, 'config', 'config.json'))
            self.assertEqual(get_app_data_path('chat_history.json'), os.path.join(directory, 'chat_history.json'))
            self.assertEqual(get_app_base_dir(), base)


class ConfigPathTests(unittest.TestCase):
    names = ('config.json', 'bookmarks.json', 'frequent_dirs.json', 'ai_workflows.json')

    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        environment = patch.dict(os.environ, TABEX_DATA_DIR=temporary.name)
        environment.start()
        self.addCleanup(environment.stop)

    def test_legacy_files_move_without_changing_content(self):
        from tabexplorer.paths import get_app_config_path, get_app_data_path, get_app_log_path
        for name in self.names:
            with self.subTest(name=name):
                original = self.root / name
                original.write_bytes(b'\xef\xbb\xbf{"value": "unchanged"}\r\n')
                target = Path(get_app_config_path(name))
                self.assertEqual(target, self.root / 'config' / name)
                self.assertEqual(target.read_bytes(), b'\xef\xbb\xbf{"value": "unchanged"}\r\n')
                self.assertFalse(original.exists())
                self.assertEqual(get_app_config_path(name), str(target))
        self.assertEqual(get_app_data_path('chat_history.json'), str(self.root / 'chat_history.json'))
        self.assertEqual(get_app_log_path('TabEx_debug_latest.log'), str(self.root / 'logs' / 'TabEx_debug_latest.log'))
        self.assertEqual(get_app_data_path('example'), str(self.root / 'example'))

    def test_existing_config_takes_priority_without_overwriting_either_file(self):
        from tabexplorer.paths import get_app_config_path
        directory = Path(get_app_config_path())
        for name in self.names:
            (self.root / name).write_bytes(b'legacy')
            (directory / name).write_bytes(b'current')
            self.assertEqual(get_app_config_path(name), str(directory / name))
            self.assertEqual((self.root / name).read_bytes(), b'legacy')
            self.assertEqual((directory / name).read_bytes(), b'current')

    def test_migration_failure_is_reported_and_preserves_original(self):
        from tabexplorer import paths
        original = self.root / 'config.json'
        original.write_bytes(b'original')
        with patch.object(paths.os, 'rename', side_effect=PermissionError('locked')):
            with self.assertRaises(PermissionError):
                paths.get_app_config_path('config.json')
        self.assertEqual(original.read_bytes(), b'original')
        self.assertFalse((self.root / 'config' / 'config.json').exists())

    def test_directory_creation_failure_is_reported(self):
        from tabexplorer.paths import get_app_config_path
        (self.root / 'config').write_bytes(b'not a directory')
        with self.assertRaises(FileExistsError):
            get_app_config_path('config.json')

    def test_source_and_frozen_configs_ignore_bundle_extraction_directory(self):
        from tabexplorer import paths
        environment = dict(os.environ)
        environment.pop('TABEX_DATA_DIR')
        source = self.root / 'source'
        (source / 'tabexplorer').mkdir(parents=True)
        executable = self.root / 'release' / 'TabExplorer.exe'
        with patch.dict(os.environ, environment, clear=True), \
                patch.object(paths, '__file__', str(source / 'tabexplorer' / 'paths.py')), \
                patch.object(sys, 'executable', str(executable)), \
                patch.object(sys, '_MEIPASS', str(self.root / 'bundle'), create=True):
            for frozen, base in ((False, source), (True, executable.parent)):
                with self.subTest(frozen=frozen), patch.object(sys, 'frozen', frozen, create=True):
                    self.assertEqual(paths.get_app_config_path('config.json'),
                                     str(base / 'config' / 'config.json'))
                    self.assertTrue((base / 'config').is_dir())
        self.assertFalse((self.root / 'bundle').exists())

    def test_default_consumers_read_and_write_in_config_directory(self):
        from tabexplorer.bookmarks import BookmarkManager
        from tabexplorer.chat import ChatPanel
        from tabexplorer.frequent import FrequentDirs
        from tabexplorer.mainwindow import MainWindow
        qt = QApplication.instance() or QApplication([])
        host = types.SimpleNamespace()
        host.config = MainWindow.load_config(host)
        self.assertTrue(MainWindow._flush_config_to_disk(host))
        bookmarks = BookmarkManager()
        bookmarks.bookmark_tree = {'bookmark_bar': {'children': [], 'type': 'folder'}}
        bookmarks.save_bookmarks(immediate=True)
        self.assertEqual(BookmarkManager().get_tree(), bookmarks.get_tree())
        frequent = FrequentDirs()
        self.assertTrue(frequent.record(r'C:\work', now=1))
        self.assertTrue(frequent.save())
        self.assertEqual(FrequentDirs().entries, frequent.entries)
        with patch.object(ChatPanel, '_setup_ui'), patch.object(ChatPanel, '_load_history'):
            panel = ChatPanel(host)
            self.assertEqual(panel.workflow_file, str(self.root / 'config' / 'ai_workflows.json'))
            self.assertEqual(panel.history_file, str(self.root / 'chat_history.json'))
            self.assertTrue(panel.workflow_templates)
            panel.deleteLater()
        qt.processEvents()
        self.assertEqual({path.name for path in (self.root / 'config').iterdir()}, set(self.names))
        for name in self.names:
            self.assertFalse((self.root / name).exists())
        host._app_started_at = time.time()
        self.assertIn(str(self.root / 'config'), MainWindow.runtime_information(host))

    def test_startup_language_reads_migrated_configuration(self):
        (self.root / 'config.json').write_text('{"language": "en"}', encoding='utf-8')
        result = subprocess.run(
            [sys.executable, '-c', 'from tabexplorer.i18n import _app_language; print(_app_language)'],
            cwd=Path(__file__).resolve().parents[1], capture_output=True, text=True, timeout=10)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), 'en')
        self.assertTrue((self.root / 'config' / 'config.json').is_file())
        self.assertFalse((self.root / 'config.json').exists())


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

    def test_installer_uses_requirements_directory(self):
        content = self.script.read_text(encoding='utf-8')
        self.assertIn("-r (Join-Path $root 'requirements\\requirements.txt')", content)
        self.assertNotIn("-r (Join-Path $root 'requirements.txt')", content)
        self.assertTrue((self.root / 'requirements' / 'requirements.txt').is_file())
        self.assertFalse((self.root / 'requirements.txt').exists())

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


class BundledToolTests(unittest.TestCase):
    def assert_tools_found(self, root):
        from tabexplorer import paths, search
        self.assertEqual(search.detect_everything(), str(root / 'tools' / 'es.exe'))
        self.assertEqual(search.detect_ripgrep(), str(root / 'tools' / 'rg.exe'))
        self.assertEqual(paths.get_bundled_path(os.path.join('tools', 'rg.exe')),
                         str(root / 'tools' / 'rg.exe'))

    def test_source_tools_are_independent_of_working_and_data_directories(self):
        root = Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(sys, '_MEIPASS', None, create=True), \
                patch.dict(os.environ, {'TABEX_DATA_DIR': directory}), \
                patch('os.getcwd', return_value=directory), \
                patch('shutil.which', side_effect=AssertionError('Bundled tools must take priority')):
            self.assert_tools_found(root)
            for name in ('es.exe', 'rg.exe'):
                self.assertFalse((root / name).exists())
                self.assertTrue((root / 'tools' / name).is_file())

    def test_frozen_tools_take_priority_over_source_tools(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'tools').mkdir()
            for name in ('es.exe', 'rg.exe'):
                (root / 'tools' / name).write_bytes(b'tool')
            with patch.object(sys, '_MEIPASS', directory, create=True), \
                    patch('shutil.which', side_effect=AssertionError('Bundled tools must take priority')):
                self.assert_tools_found(root)

    def test_missing_bundled_search_tools_still_fall_back_to_path(self):
        from tabexplorer import paths, search
        with tempfile.TemporaryDirectory() as directory, \
                patch.object(sys, '_MEIPASS', None, create=True), \
                patch.object(paths, 'get_app_base_dir', return_value=directory), \
                patch('shutil.which', side_effect=lambda name: str(Path(directory) / 'path' / name)):
            self.assertEqual(search.detect_everything(), str(Path(directory) / 'path' / 'es.exe'))
            self.assertEqual(search.detect_ripgrep(), str(Path(directory) / 'path' / 'rg.exe'))
            self.assertIsNone(paths.get_bundled_path(os.path.join('tools', 'rg.exe')))


class BuildPipelineTests(unittest.TestCase):
    def test_runtime_collection_overrides_both_root_and_qt_dlls(self):
        from tools.build_release import collect_msvc_runtime, MSVC_DLLS
        image = Mock()
        image.FILE_HEADER.Machine = 0x8664
        image.VS_FIXEDFILEINFO = [types.SimpleNamespace(FileVersionMS=(14 << 16) | 42)]
        image.__enter__ = Mock(return_value=image)
        image.__exit__ = Mock(return_value=False)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in MSVC_DLLS:
                (root / name).touch()
            with patch('pefile.PE', return_value=image):
                binaries = collect_msvc_runtime(root)
                self.assertEqual(binaries, [(root / name, destination) for name in MSVC_DLLS
                                            for destination in ('.', r'PyQt5\Qt5\bin')])
                for machine, version in ((0x14c, 42), (0x8664, 26)):
                    image.FILE_HEADER.Machine = machine
                    image.VS_FIXEDFILEINFO[0].FileVersionMS = (14 << 16) | version
                    with self.assertRaisesRegex(RuntimeError, '14.42'):
                        collect_msvc_runtime(root)
            (root / MSVC_DLLS[0]).unlink()
            with self.assertRaisesRegex(RuntimeError, 'Missing'):
                collect_msvc_runtime(root)

    def test_build_installs_from_requirements_directory_before_building(self):
        from tools.build_release import build
        root = Path(__file__).resolve().parents[1]
        requirements = root / 'requirements' / 'requirements-build.txt'
        self.assertTrue(requirements.is_file())
        self.assertFalse((root / 'requirements-build.txt').exists())
        run = Mock(side_effect=subprocess.CalledProcessError(1, 'pip'))
        with self.assertRaises(subprocess.CalledProcessError):
            build(root, run=run)
        run.assert_called_once_with(
            [sys.executable, '-X', 'utf8', '-m', 'pip', 'install', '-r', str(requirements)],
            cwd=root, check=True)

    def test_failed_build_or_smoke_never_replaces_working_executable(self):
        """打包、原生验证或原子替换失败都保留已有可用 EXE。"""
        from tools.build_release import build
        import subprocess
        for failure in ('build', 'validation', 'replace'):
            with self.subTest(failure=failure), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                executable = root / 'TabExplorer.exe'
                executable.write_bytes(b'old working executable')
                (root / 'tools').mkdir()
                for tool in ('es.exe', 'rg.exe'):
                    (root / 'tools' / tool).write_bytes(b'tool')

                def run(command, **keywords):
                    if 'PyInstaller' in command:
                        if failure == 'build':
                            raise subprocess.CalledProcessError(1, command)
                        destination = Path(command[command.index('--distpath') + 1])
                        destination.mkdir()
                        (destination / 'TabExplorer.exe').write_bytes(b'candidate')

                verify = Mock(return_value={'ok': True, 'frozen': True, 'ocr': {'ok': True}})
                if failure == 'validation':
                    verify.side_effect = RuntimeError('native failed')
                with ExitStack() as stack:
                    stack.enter_context(patch('tools.build_release.collect_msvc_runtime', return_value=[]))
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
            (root / 'tools').mkdir()
            for tool in ('es.exe', 'rg.exe'):
                (root / 'tools' / tool).write_bytes(b'tool')
            commands = []

            def run(command, **keywords):
                commands.append(command)
                if 'PyInstaller' in command:
                    destination = Path(command[command.index('--distpath') + 1])
                    destination.mkdir()
                    (destination / 'TabExplorer.exe').write_bytes(b'new')

            def verify(*arguments, **keywords):
                self.assertEqual(executable.read_bytes(), b'old')
                return {'ok': True, 'frozen': True, 'ocr': {'ok': True}}

            runtime = root / 'msvcp140.dll'
            with patch('tools.build_release.collect_msvc_runtime',
                       return_value=[(runtime, '.'), (runtime, r'PyQt5\Qt5\bin')]):
                self.assertEqual(build(root, install=False, run=run, verify=verify), executable)
            self.assertIn('unittest', commands[0])
            command = commands[1]
            self.assertEqual(command[command.index('--collect-data') + 1], 'rapidocr')
            self.assertEqual(command[command.index('--collect-submodules') + 1],
                             'rapidocr.inference_engine.onnxruntime')
            self.assertIn('pygments.lexers',
                          [command[index + 1] for index, value in enumerate(command) if value == '--collect-submodules'])
            self.assertIn('pypdfium2',
                          [command[index + 1] for index, value in enumerate(command) if value == '--copy-metadata'])
            self.assertEqual(command[command.index('--runtime-hook') + 1],
                             str(root / 'tools' / 'pyi_rth_ocr.py'))
            bundled_data = [command[index + 1] for index, value in enumerate(command) if value == '--add-data']
            self.assertEqual(bundled_data, [str(root / 'icons') + ';icons'] +
                             [str(root / 'tools' / tool) + ';tools' for tool in ('es.exe', 'rg.exe')])
            bundled_binaries = [command[index + 1] for index, value in enumerate(command)
                                if value == '--add-binary']
            self.assertEqual(bundled_binaries, [str(runtime) + ';.', str(runtime) + r';PyQt5\Qt5\bin'])
            self.assertEqual(executable.read_bytes(), b'new')

    def test_missing_or_failed_packaged_ocr_never_promotes_executable(self):
        from tools.build_release import build
        for result in ({'ok': True, 'frozen': True},
                       {'ok': True, 'frozen': True, 'ocr': {'ok': False}}):
            with self.subTest(result=result), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                (root / 'tools').mkdir()
                for name in ('es.exe', 'rg.exe'):
                    (root / 'tools' / name).touch()
                executable = root / 'TabExplorer.exe'
                executable.write_bytes(b'old')

                def run(command, **keywords):
                    if 'PyInstaller' in command:
                        destination = Path(command[command.index('--distpath') + 1])
                        destination.mkdir()
                        (destination / 'TabExplorer.exe').write_bytes(b'candidate')

                with patch('tools.build_release.collect_msvc_runtime', return_value=[]):
                    with self.assertRaisesRegex(RuntimeError, 'OCR validation'):
                        build(root, install=False, run=run, verify=Mock(return_value=result))
                self.assertEqual(executable.read_bytes(), b'old')


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
            with patch('tabexplorer.json_store._preserve_unreadable_file', return_value=None):
                config = store.load()
            self.assertFalse(store.save(config))
            self.assertEqual(path.read_text(encoding='utf-8'), '{broken')

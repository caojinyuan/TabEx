"""Data preservation, navigation ordering and bounded logging contracts."""

import json
import io
import os
from pathlib import Path
import tempfile
import types
import unittest
import zipfile
from contextlib import ExitStack
from unittest.mock import Mock, patch

from PyQt5.QtCore import QEvent, Qt, QTimer
from PyQt5.QtWidgets import QApplication, QWidget

from tabexplorer.bookmarks import BookmarkManager
from tabexplorer.chat import ChatPanel
from tabexplorer.frequent import FrequentDirs
from tabexplorer.json_store import JsonFileStore, _preserve_unreadable_file
from tabexplorer.persistence import ConfigStore
from tabexplorer.explorer_tab import FileExplorerTab
from tabexplorer.mainwindow import MainWindow
from tabexplorer.paths import translate_common_path
from tabexplorer.shellview import IExplorerBrowserWidget
from tabexplorer import debuglog
from tabexplorer.diagnostics import Diagnostics


class DataProtectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def test_transient_bookmark_read_failure_blocks_all_automatic_saves(self):
        path = self.root / 'bookmarks.json'
        path.write_text('{"roots": {"saved": {}}}', encoding='utf-8')
        with patch('builtins.open', side_effect=PermissionError('read denied')):
            manager = BookmarkManager(str(path))
        notify = manager.error_handler = Mock()
        self.assertFalse(manager.save_bookmarks())
        self.assertFalse(manager.save_bookmarks(immediate=True))
        self.assertTrue(manager.write_blocked)
        self.assertIsNone(manager._save_timer)
        self.assertEqual(path.read_text(encoding='utf-8'), '{"roots": {"saved": {}}}')
        notify.assert_called_once()

    def test_invalid_bookmark_shape_is_backed_up_and_backup_failure_blocks_writes(self):
        path = self.root / 'bookmarks.json'
        for allowed in (True, False):
            with self.subTest(backup_allowed=allowed):
                path.write_text('{"roots": []}', encoding='utf-8')
                with patch('tabexplorer.json_store._preserve_unreadable_file',
                           wraps=_preserve_unreadable_file if allowed else None,
                           **({} if allowed else {'return_value': ''})):
                    manager = BookmarkManager(str(path))
                manager.bookmark_tree = {'new': {}}
                self.assertEqual(manager.save_bookmarks(immediate=True), allowed)
                original = Path(manager.recovered_backup) if allowed else path
                self.assertEqual(original.read_text(encoding='utf-8'), '{"roots": []}')

    def test_backups_created_in_same_second_do_not_overwrite_each_other(self):
        path = self.root / 'data.json'
        backups = []
        with patch('time.strftime', return_value='20261030-183000'):
            for value in ('first', 'second'):
                path.write_text(value, encoding='utf-8')
                backups.append(Path(_preserve_unreadable_file(str(path))))
        self.assertNotEqual(backups[0], backups[1])
        self.assertEqual([path.read_text(encoding='utf-8') for path in backups], ['first', 'second'])

    def history_host(self, path):
        host = types.SimpleNamespace(history_file=str(path), messages=[], main_window=None,
                                     _trim_chat_history=Mock(), _rebuild_chat_display=Mock())
        host._save_history = types.MethodType(ChatPanel._save_history, host)
        return host

    def test_invalid_chat_shape_is_preserved_before_new_history_is_saved(self):
        path = self.root / 'chat_history.json'
        path.write_text('{"unexpected": "preserve me"}', encoding='utf-8')
        host = self.history_host(path)
        with patch('tabexplorer.chat.show_toast') as toast:
            ChatPanel._load_history(host)
        toast.assert_called_once()
        self.assertFalse(path.exists())
        backup = Path(host._history_store.recovered_backup)
        host.messages = [{'role': 'user', 'content': 'new'}]
        self.assertTrue(host._save_history())
        self.assertEqual(backup.read_text(encoding='utf-8'), '{"unexpected": "preserve me"}')
        self.assertEqual(json.loads(path.read_text(encoding='utf-8')), host.messages)

    def test_chat_read_failure_does_not_allow_later_messages_to_overwrite_history(self):
        path = self.root / 'chat_history.json'
        path.write_text('[{"role": "user", "content": "old"}]', encoding='utf-8')
        host = self.history_host(path)
        with patch('builtins.open', side_effect=PermissionError('locked')), patch('tabexplorer.chat.show_toast'):
            ChatPanel._load_history(host)
        host.messages = [{'role': 'user', 'content': 'new'}]
        with patch('tabexplorer.chat.show_toast') as toast:
            self.assertFalse(host._save_history())
            self.assertFalse(host._save_history())
        toast.assert_called_once()
        self.assertIn('old', path.read_text(encoding='utf-8'))

    def test_frequent_directory_read_failure_blocks_exit_save(self):
        path = self.root / 'frequent_dirs.json'
        path.write_text('[]', encoding='utf-8')
        with patch('builtins.open', side_effect=PermissionError('locked')):
            store = FrequentDirs(path)
        store.record(r'C:\new', now=1)
        self.assertFalse(store.save())
        self.assertTrue(store.dirty)
        self.assertEqual(path.read_text(encoding='utf-8'), '[]')

    def test_workflow_invalid_shape_is_preserved_and_unreadable_file_is_not_replaced(self):
        path = self.root / 'ai_workflows.json'
        host = types.SimpleNamespace(workflow_file=str(path), main_window=None,
                                     _default_workflow_templates=lambda: {'sample': 'default'})
        host._write_default_workflow_file = types.MethodType(ChatPanel._write_default_workflow_file, host)
        path.write_text('{"workflows": []}', encoding='utf-8')
        with patch('tabexplorer.chat.show_toast'):
            self.assertEqual(ChatPanel._load_workflow_templates(host), {'sample': 'default'})
        self.assertEqual(Path(host._workflow_store.recovered_backup).read_text(encoding='utf-8'),
                         '{"workflows": []}')
        path.write_text('{"workflows": {"sample": "custom"}}', encoding='utf-8')
        with patch('builtins.open', side_effect=PermissionError('locked')), patch('tabexplorer.chat.show_toast'):
            ChatPanel._load_workflow_templates(host)
        with patch('tabexplorer.chat.show_toast'):
            host._write_default_workflow_file({'sample': 'new'})
        self.assertIn('custom', path.read_text(encoding='utf-8'))

    def test_failed_encryption_never_writes_plaintext_key_or_overwrites_config(self):
        path = self.root / 'config.json'
        path.write_text('{"old": true}', encoding='utf-8')
        store = ConfigStore(path)
        with patch('tabexplorer.persistence.protect_secret', return_value=''):
            self.assertFalse(store.save({'ai_chat': {'api_key': 'private-key'}}))
        self.assertEqual(path.read_text(encoding='utf-8'), '{"old": true}')
        self.assertIn('encryption failed', store.last_error)
        for invalid in (None, []):
            self.assertFalse(store.save(invalid))
            self.assertEqual(path.read_text(encoding='utf-8'), '{"old": true}')

    def test_undecryptable_key_is_preserved_during_unrelated_config_save(self):
        path = self.root / 'config.json'
        path.write_text('{"ai_chat": {"api_key_protected": "preserved-ciphertext"}}', encoding='utf-8')
        store = ConfigStore(path)
        with patch('tabexplorer.persistence.unprotect_secret', return_value=None):
            config = store.load()
        self.assertEqual(config['ai_chat']['api_key'], '')
        self.assertTrue(store.secret_unreadable)
        self.assertTrue(store.save(config))
        self.assertEqual(json.loads(path.read_text(encoding='utf-8'))['ai_chat']['api_key_protected'],
                         'preserved-ciphertext')

    def test_atomic_save_failure_is_retryable_and_identical_data_is_not_rewritten(self):
        path = self.root / 'data.json'
        path.write_text('{"old": true}', encoding='utf-8')
        store = JsonFileStore(path, dict)
        with patch('os.replace', side_effect=PermissionError('busy')):
            self.assertFalse(store.save({'new': True}))
        self.assertEqual(path.read_text(encoding='utf-8'), '{"old": true}')
        self.assertFalse(Path(str(path) + '.tmp').exists())
        self.assertTrue(store.save({'new': True}))
        with patch('builtins.open', side_effect=AssertionError('unchanged save must not perform IO')):
            self.assertTrue(store.save({'new': True}))


class NavigationOrderingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.first, self.second = self.root / 'A', self.root / 'B'
        self.first.mkdir()
        self.second.mkdir()
        host = self.host = QWidget()
        self.addCleanup(host.deleteLater)
        host._is_cleaning_up = False
        host.folder_checker = None
        host.current_path = str(self.root)
        host._is_control_panel_path = lambda _path: False
        host.path_bar = Mock()
        host.explorer = Mock()
        host._is_mycomputer_shell_target = lambda _path: False
        for name in ('_show_loading_indicator', '_hide_loading_indicator', '_show_network_error',
                     '_mark_expected_navigation', '_resume_path_sync_after_navigation', 'update_tab_title',
                     '_add_to_history', 'update_explorer_status'):
            setattr(host, name, Mock())
        for name in ('_normalize_local_path', '_release_folder_checker', '_check_folder_size_async',
                     '_on_folder_size_checked', '_is_slow_path', 'navigate_to', '_on_navigation_failed',
                     '_handle_navigation_failure', '_on_async_nav_started', '_async_nav_safety_timeout',
                     '_on_async_nav_finished'):
            setattr(host, name, types.MethodType(getattr(FileExplorerTab, name), host))
        self.applied = []
        host._perform_navigation = lambda path, history: self.applied.append(path)
        self.addCleanup(lambda: host._release_folder_checker(2000))

    def test_finished_but_queued_folder_check_cannot_override_new_direct_navigation(self):
        host = self.host
        host._check_folder_size_async(str(self.first), True)
        self.assertTrue(host.folder_checker.wait(2000))
        self.assertFalse(host.folder_checker.isRunning())
        host.navigate_to(str(self.second), skip_async_check=True)
        self.qt.processEvents()
        self.assertEqual(self.applied, [str(self.second)])

    def test_queued_folder_check_cannot_override_shell_or_network_navigation(self):
        host = self.host
        for target, shell in (('shell:Desktop', True), (r'\\server\share', False)):
            with self.subTest(target=target):
                self.applied.clear()
                host.explorer.reset_mock()
                host._check_folder_size_async(str(self.first), True)
                self.assertTrue(host.folder_checker.wait(2000))
                host.navigate_to(target, is_shell=shell)
                self.qt.processEvents()
                self.assertNotIn(str(self.first), self.applied)
                if shell:
                    host.explorer.dynamicCall.assert_called_once_with('Navigate(const QString&)', target)
                else:
                    self.assertEqual(self.applied, [target])

    def test_queued_failure_does_not_hide_or_unlock_new_navigation(self):
        host = self.host
        host._on_navigation_failed(str(self.first), -2147024894)
        host.navigate_to(str(self.second), skip_async_check=True)
        host._nav_in_progress = True
        self.qt.processEvents()
        host._show_network_error.assert_not_called()
        self.assertTrue(host._nav_in_progress)

    def test_old_timeout_does_not_cancel_a_retry_of_the_same_path(self):
        host = self.host
        callbacks = []
        with patch.object(QTimer, 'singleShot', side_effect=lambda _delay, callback: callbacks.append(callback)):
            host._on_async_nav_started(str(self.first))
        host._nav_in_progress = False
        host.navigate_to(str(self.first), skip_async_check=True)
        host._nav_in_progress = True
        host._nav_in_progress_path = str(self.first)
        callbacks[0]()
        self.assertTrue(host._nav_in_progress)
        host._show_network_error.assert_not_called()

    def test_old_async_failure_does_not_clear_new_path_loading_state(self):
        self.host._nav_in_progress = True
        self.host._nav_in_progress_path = str(self.second)
        self.host._on_async_nav_finished(str(self.first), False)
        self.assertTrue(self.host._nav_in_progress)
        self.host._hide_loading_indicator.assert_not_called()

    def test_duplicate_completion_of_current_view_does_not_cancel_new_folder_check(self):
        host = self.host
        host._check_folder_size_async(str(self.first), True)
        checker = host.folder_checker
        self.assertTrue(checker.wait(2000))
        FileExplorerTab._on_navigation_succeeded(host, host.current_path)
        self.qt.processEvents()
        self.assertEqual(self.applied, [str(self.first)])

    def test_native_navigation_to_different_location_invalidates_queued_folder_check(self):
        host = self.host
        host._check_folder_size_async(str(self.first), True)
        self.assertTrue(host.folder_checker.wait(2000))
        FileExplorerTab._on_navigation_succeeded(host, str(self.second))
        self.qt.processEvents()
        self.assertEqual(self.applied, [])

    def test_hidden_shell_widget_invalidates_old_pidl_resolution(self):
        host = types.SimpleNamespace(_nav_generation=3, _resolving_generation=3,
                                     isVisible=lambda: False, _pending_path=None)
        IExplorerBrowserWidget._navigate(host, str(self.second))
        self.assertGreater(host._nav_generation, 3)
        self.assertIsNone(host._resolving_generation)
        self.assertEqual(host._pending_path, str(self.second))

    def test_slow_address_is_classified_before_any_filesystem_probe(self):
        import ctypes
        host = self.host
        host.navigate_to = Mock()
        for path, drive_type in ((r'Z:\offline', 4), (r'C:\Users\test\OneDrive\offline', 3),
                                 (r'\\server\share\offline', 4)):
            with self.subTest(path=path), \
                    patch.object(ctypes.windll.kernel32, 'GetDriveTypeW', return_value=drive_type), \
                    patch('os.path.exists', side_effect=AssertionError('UI filesystem probe')), \
                    patch('os.path.isdir', side_effect=AssertionError('UI filesystem probe')):
                host.navigate_to.reset_mock()
                FileExplorerTab.on_path_bar_changed(host, path)
                host.navigate_to.assert_called_once_with(path)
                self.assertEqual(translate_common_path(path), path)


class SessionRestoreWorkTests(unittest.TestCase):
    """真实标签栈与导航流程；拦截原生浏览、Git 查询和目录快照，不接管桌面或读取用户配置。"""

    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        stack = ExitStack()
        self.addCleanup(stack.close)
        root = Path(stack.enter_context(tempfile.TemporaryDirectory()))
        stack.enter_context(patch.dict(os.environ, TABEX_DATA_DIR=str(root)))
        self.paths = [str(root / f'work-{i}') for i in range(18)]
        for path in self.paths:
            Path(path).mkdir()
        self.left = self.paths[1:] + [self.paths[1]]
        self.right = ['shell:RecycleBinFolder', 'shell:Desktop']
        config = {
            'pinned_tabs': [self.paths[0]], 'enable_cache_tabs': True,
            'cached_tabs': [{'path': p, 'bookmark_group_color': '#64B5F6'} for p in self.left],
            'last_active_tab_path': self.paths[-1],
            'split_session': {'active': True, 'active_index': 1, 'tabs': [
                {'path': p, 'is_shell': True} for p in self.right]},
            'auto_update_check': False, 'enable_explorer_monitor': False,
            'enable_title_shortcuts': False, 'hotkeys': {'screenshot': False}, 'named_workspaces': {},
        }
        (root / 'config').mkdir()
        (root / 'config' / 'config.json').write_text(json.dumps(config), encoding='utf-8')
        stack.enter_context(patch.object(BookmarkManager, 'save_bookmarks', return_value=True))
        stack.enter_context(patch.object(MainWindow, '_start_shortcut_listener'))
        stack.enter_context(patch.object(MainWindow, '_repair_startup_shortcut'))
        stack.enter_context(patch('tabexplorer.navigation.ASYNC_LOAD_ENABLED', False))
        stack.enter_context(patch.object(FileExplorerTab, '_request_dir_snapshot'))
        self.browse = stack.enter_context(patch.object(IExplorerBrowserWidget, '_navigate', autospec=True))
        self.git = stack.enter_context(patch.object(FileExplorerTab, '_request_git_status_async', autospec=True))
        # Drive initialization explicitly, both before and after the window is shown.
        with patch.object(MainWindow, '_delayed_initialization', lambda self: None):
            self.window = MainWindow()
        self.window.setAttribute(Qt.WA_ShowWithoutActivating)
        self.addCleanup(self.dispose)

    def dispose(self):
        self.window.close()
        self.window.deleteLater()
        self.qt.sendPostedEvents(None, QEvent.DeferredDelete)

    def browsed(self):
        return [call.args[1] for call in self.browse.call_args_list]

    def test_restore_loads_only_selected_pages_then_browses_once_on_first_switch(self):
        w = self.window
        w.show()
        with patch.object(w, 'populate_bookmark_bar_menu', wraps=w.populate_bookmark_bar_menu) as menu:
            w._delayed_initialization()
            menu.assert_called_once()
        self.assertEqual(self.browsed(), [self.paths[-1], self.right[1]])
        self.assertEqual({call.args[1] for call in self.git.call_args_list}, {self.paths[-1]})
        saved = w._capture_named_workspace()['groups']
        self.assertEqual([t['path'] for t in saved[0]['tabs']], [self.paths[0], *self.left])
        self.assertEqual([t['path'] for t in saved[1]['tabs']], self.right)
        self.assertTrue(saved[0]['tabs'][0]['is_pinned'])
        self.assertEqual(saved[0]['tabs'][1]['bookmark_group_color'], '#64B5F6')
        self.assertEqual([g['active_index'] for g in saved], [len(self.paths) - 1, 1])
        w.tab_widget.setCurrentIndex(0)
        w.tab_widget.setCurrentIndex(2)
        w.tab_widget.setCurrentIndex(0)
        self.assertEqual(self.browsed()[2:], [self.paths[0], self.paths[2]])
        unused = w.content_stack.widget(3)
        self.assertFalse(hasattr(unused, 'explorer'))
        w.close_tab(3)
        self.assertNotIn(self.paths[3], self.browsed())

    def test_hidden_restore_fallback_and_workspace_switch_skip_transient_pages(self):
        w = self.window
        w.config['last_active_tab_path'] = ''
        w._delayed_initialization()
        self.browse.assert_not_called()
        self.git.assert_not_called()
        w.show()
        self.assertCountEqual(self.browsed(), [self.paths[0], self.right[1]])
        w.config['named_workspaces']['two'] = {'groups': [
            {'active_index': 2, 'tabs': [{'path': p} for p in self.paths[:5]]},
            {'active_index': 1, 'tabs': [{'path': p} for p in self.paths[5:8]]}]}
        self.browse.reset_mock()
        with patch.object(w, 'populate_bookmark_bar_menu', wraps=w.populate_bookmark_bar_menu) as menu:
            self.assertTrue(w.switch_named_workspace('two'))
            menu.assert_called_once()
        self.assertEqual(self.browsed(), [self.paths[2], self.paths[6]])
        self.assertEqual((w.tab_widget.count(), w.split_tab_widget.count()), (5, 3))

    def test_close_and_reopen_preserves_duplicate_instances_pin_order_active_side_and_lazy_views(self):
        w = self.window
        w.show()
        w._delayed_initialization()
        a, b, c = self.paths[:3]
        w.config['named_workspaces']['duplicates'] = {'groups': [
            {'active_index': 3, 'tabs': [{'path': a, 'is_pinned': True}, {'path': b},
                                          {'path': a}, {'path': b, 'bookmark_group_color': '#64B5F6'}]},
            {'active_index': 1, 'tabs': [{'path': b, 'is_pinned': True}, {'path': c}, {'path': b}]}],
            'active_group': 1}
        with patch('tabexplorer.mainwindow.show_toast'):
            self.assertTrue(w.switch_named_workspace('duplicates'))
        expected = w._capture_named_workspace()
        self.assertEqual([g['active_index'] for g in expected['groups']], [3, 1])
        self.assertEqual(expected['active_group'], 1)
        self.dispose()
        with patch.object(MainWindow, '_delayed_initialization', lambda self: None):
            self.window = MainWindow()
        self.window.setAttribute(Qt.WA_ShowWithoutActivating)
        self.window.show()
        self.browse.reset_mock()
        self.window._delayed_initialization()
        self.assertEqual(self.window._capture_named_workspace(), expected)
        self.assertEqual(self.browsed(), [b, c])

    def test_workspace_switch_undo_restores_split_and_append_does_not_replace_undo(self):
        w = self.window
        w.show()
        w._delayed_initialization()
        before = w._capture_named_workspace()
        w.config['named_workspaces']['one'] = {'groups': [{'tabs': [{'path': self.paths[4]}], 'active_index': 0}]}
        with patch('tabexplorer.mainwindow.show_toast') as toast:
            w.switch_named_workspace('one')
            undo_click = toast.call_args.kwargs['action']
            self.assertTrue(w.undo_workspace_action.isEnabled())
            self.assertFalse(w._split_active)
            w.open_named_workspace('one')
            self.assertEqual(w.tab_widget.count(), 2)
            self.assertTrue(undo_click())
            self.assertEqual(w._capture_named_workspace(), before)
            self.assertFalse(w.undo_workspace_action.isEnabled())
            self.assertFalse(undo_click())

    def test_restored_missing_local_folder_keeps_tab_and_can_retry(self):
        missing = Path(self.paths[0]) / 'restored-later'
        w = self.window
        w.config['tab_session'] = {'groups': [{'tabs': [{'path': str(missing)}]}]}
        w.show()
        w._delayed_initialization()
        tab = w.get_current_tab_widget()
        self.assertEqual(tab.current_path, str(missing))
        self.assertTrue(tab.net_banner.isVisible())
        self.assertTrue(tab.explorer.isHidden())
        self.assertTrue(tab.net_banner.reconnect_button.isHidden())
        tab.net_banner.retry_button.click()
        self.assertTrue(tab.net_banner.retry_button.isEnabled())
        self.browse.assert_not_called()
        missing.mkdir()
        tab.net_banner.retry_button.click()
        self.assertEqual(self.browsed(), [str(missing)])


class DebugLogRotationTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.root = Path(self.stack.enter_context(tempfile.TemporaryDirectory()))
        self.path = self.root / 'TabEx_debug_latest.log'
        self.stack.enter_context(patch.object(debuglog, '_DEBUG_LOG_PATH', str(self.path)))
        self.stack.enter_context(patch.object(debuglog, '_DEBUG_LOG_READY', False))
        self.stack.enter_context(patch.object(debuglog, '_DEBUG_LOG_ERROR', ''))
        self.stack.enter_context(patch.object(debuglog, '_diagnostic_event'))
        debuglog._init_debug_log_file()

    def test_exact_five_mib_limit_rotates_before_crossing_boundary(self):
        self.assertEqual(debuglog.DEBUG_LOG_MAX_BYTES, 5 * 1024 * 1024)
        self.assertEqual(debuglog.DEBUG_LOG_BACKUP_COUNT, 3)
        self.path.write_bytes(b'x' * (debuglog.DEBUG_LOG_MAX_BYTES - 1))
        self.assertTrue(debuglog._append_debug_log_line('[ClosedTabs] next'))
        self.assertEqual(Path(str(self.path) + '.1').stat().st_size, debuglog.DEBUG_LOG_MAX_BYTES - 1)
        self.assertLessEqual(self.path.stat().st_size, debuglog.DEBUG_LOG_MAX_BYTES)

    def test_utf8_rotation_keeps_at_most_three_backups(self):
        with patch.object(debuglog, 'DEBUG_LOG_MAX_BYTES', 256):
            for index in range(30):
                self.assertTrue(debuglog._append_debug_log_line(f'[ClosedTabs] seq:{index} ' + '\u4e2d' * 30))
        files = sorted(self.root.glob('TabEx_debug_latest.log*'))
        self.assertEqual(len(files), 4)
        self.assertTrue(all(path.stat().st_size <= 256 for path in files))
        for path in files:
            path.read_text(encoding='utf-8')
        self.assertIn('seq:29', self.path.read_text(encoding='utf-8'))
        self.assertFalse(Path(str(self.path) + '.4').exists())

    def test_oversized_single_entry_is_truncated_to_valid_utf8(self):
        with patch.object(debuglog, 'DEBUG_LOG_MAX_BYTES', 128):
            self.assertTrue(debuglog._append_debug_log_line('[ClosedTabs] ' + '\u4e2d' * 1000))
        self.assertLessEqual(self.path.stat().st_size, 128)
        self.assertIn('oversized entry truncated', self.path.read_text(encoding='utf-8'))

    def test_rotation_failure_is_reported_once_and_retry_can_recover(self):
        original = b'x' * 120
        self.path.write_bytes(original)
        with patch.object(debuglog, 'DEBUG_LOG_MAX_BYTES', 128):
            with patch('os.replace', side_effect=PermissionError('locked')), \
                    patch('sys.stderr', new_callable=io.StringIO) as stderr:
                self.assertFalse(debuglog._append_debug_log_line('next'))
                self.assertFalse(debuglog._append_debug_log_line('next'))
            self.assertEqual(stderr.getvalue().count('Write failed'), 1)
            self.assertEqual(self.path.read_bytes(), original)
            self.assertTrue(debuglog._DEBUG_LOG_ERROR)
            self.assertTrue(debuglog._append_debug_log_line('recovered'))
            self.assertEqual(debuglog._DEBUG_LOG_ERROR, '')

    def test_export_includes_only_filtered_and_bounded_rotation_backups(self):
        service = Diagnostics(self.root / 'diagnostics', 'test')
        self.addCleanup(service.close)
        for index in range(1, 5):
            Path(str(self.path) + f'.{index}').write_text(
                f'[AI] PRIVATE\n[ClosedTabs] backup-{index} path="C:\\secret\\file"\n', encoding='utf-8')
        service.capture_debug(self.path)
        bundle = self.root / 'bundle.zip'
        service.export(bundle, self.path)
        with zipfile.ZipFile(bundle) as archive:
            combined = '\n'.join(archive.read(name).decode('utf-8') for name in archive.namelist())
            self.assertNotIn('PRIVATE', combined)
            self.assertNotIn('secret', combined)
            for index in range(1, 4):
                self.assertIn(f'debug-filtered.{index}.log', archive.namelist())
                self.assertIn(f'sessions/{service.session.name}/previous-debug-filtered.{index}.log',
                              archive.namelist())
            self.assertNotIn('backup-4', combined)

    def test_write_error_is_notified_in_ui_once_per_failure_episode(self):
        from tabexplorer.mainwindow import MainWindow
        host = types.SimpleNamespace()
        with patch.object(debuglog, '_diagnostics', None), patch('tabexplorer.mainwindow.show_toast') as toast:
            debuglog._DEBUG_LOG_ERROR = 'disk full'
            MainWindow._capture_diagnostic_snapshot(host)
            MainWindow._capture_diagnostic_snapshot(host)
            toast.assert_called_once()
            debuglog._DEBUG_LOG_ERROR = ''
            MainWindow._capture_diagnostic_snapshot(host)
            debuglog._DEBUG_LOG_ERROR = 'disk full'
            MainWindow._capture_diagnostic_snapshot(host)
            self.assertEqual(toast.call_count, 2)


class ModuleBoundaryTests(unittest.TestCase):
    def test_tab_keeps_navigation_api_after_extraction(self):
        from tabexplorer.navigation import NavigationMixin
        self.assertTrue(issubclass(FileExplorerTab, QWidget))
        self.assertTrue(issubclass(FileExplorerTab, NavigationMixin))
        for name in ('navigate_to', 'on_path_bar_changed', '_perform_navigation',
                     '_handle_navigation_failure', '_check_folder_size_async'):
            self.assertIs(getattr(FileExplorerTab, name), getattr(NavigationMixin, name))
            self.assertNotIn(name, FileExplorerTab.__dict__)

    def test_persistent_stores_share_failure_protection(self):
        for store in (ConfigStore, BookmarkManager, FrequentDirs):
            self.assertTrue(issubclass(store, JsonFileStore))


if __name__ == '__main__':
    unittest.main()

"""Scope consistency, external tool readiness, file collections and runtime identity."""

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
import types
import unittest
from unittest.mock import Mock, patch

from PyQt5.QtCore import QCoreApplication, QEvent, QItemSelectionModel, Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QMessageBox, QTreeWidgetItem, QWidget
from PyQt5.QtTest import QTest

from tabexplorer import context_tools as tools
from tabexplorer import search
from tabexplorer import widgets
from tabexplorer import i18n
from tabexplorer.file_basket import FileBasketDialog
from tabexplorer.fileops import DirectoryCompareDialog, FileTaskPanel
from tabexplorer.mainwindow import MainWindow


class QtTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def wait_for(self, condition, timeout=5):
        deadline = time.monotonic() + timeout
        while not condition() and time.monotonic() < deadline:
            self.qt.processEvents()
            QTest.qWait(5)
        self.assertTrue(condition())


class ProjectSearchTests(QtTests):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.rg = search.detect_ripgrep()
        self.assertTrue(self.rg)
        self.names = ('src/a.txt', '.hidden.txt', 'build/b.txt', 'nested/node_modules/dep.txt',
                      'ignored.txt', 'keep.log')
        for name in self.names:
            file = self.root / name
            file.parent.mkdir(parents=True, exist_ok=True)
            file.write_text('needleneedle\n', encoding='utf-8')
        (self.root / '.gitignore').write_text('ignored.txt\n', encoding='utf-8')

    def run_search(self, project=False, exclusions=(), filename=False, keyword='needle', **options):
        task = search._SearchTask(str(self.root), None, 100, self.rg, project, exclusions)
        task.run(keyword, filename, True, '*.txt', None, False, False, False,
                 options.get('regex', False), None)
        messages = []
        while not task.result_queue.empty():
            messages.append(task.result_queue.get_nowait())
        rows = [row for message in messages if message['type'] == 'result_batch' for row in message['items']]
        return {Path(row['path']).relative_to(self.root).as_posix() for row in rows}, messages

    def test_project_ignores_hidden_build_dependencies_and_gitignored_even_with_extension_filter(self):
        files, messages = self.run_search(project=True)
        self.assertEqual(files, {'src/a.txt'})
        self.assertFalse(any('搜索未完成' in message.get('text', '') for message in messages))
        all_files, _messages = self.run_search()
        self.assertEqual(all_files, set(self.names) - {'keep.log'})

    def test_exclusions_override_extension_include_in_all_mode(self):
        exclusions = search.parse_search_exclusions('build/;node_modules/;ignored.txt')
        files, _messages = self.run_search(exclusions=exclusions)
        self.assertEqual(files, {'src/a.txt', '.hidden.txt'})
        self.assertEqual(search.parse_search_exclusions(r'build\;build/;*.log'),
                         ('**/build/**', '*.log'))
        with self.assertRaises(ValueError):
            search.parse_search_exclusions('!*.txt')

    def test_filename_listing_and_python_regex_fallback_keep_identical_scope(self):
        task = search._SearchTask(str(self.root), None, 100, self.rg, True)
        entries = list(search.iter_search_entries(task))
        self.assertEqual({name for _root, _dirs, files in entries for name in files}, {'a.txt', 'keep.log'})
        with patch.object(search.os, 'walk', side_effect=AssertionError('unscoped traversal')):
            files, messages = self.run_search(project=True, keyword=r'(needle)\1', regex=True)
        self.assertEqual(files, {'src/a.txt'})
        self.assertTrue(any(search.tr("ripgrep 不可用，已回退到普通搜索: {}").split('{}')[0]
                            in message.get('text', '') for message in messages))

    def test_scoped_enumerator_releases_process_when_closed_early(self):
        task = search._SearchTask(str(self.root), None, 100, self.rg, True)
        entries = search.iter_search_entries(task)
        self.assertTrue(next(entries)[2])
        entries.close()
        self.assertFalse(any(thread.name == 'rg-output' for thread in threading.enumerate()))

    def test_unavailable_scoped_engine_reports_error_instead_of_broadening_scope(self):
        task = search._SearchTask(str(self.root), None, 100, None, True)
        with patch.object(search.os, 'walk') as walk:
            task.run('needle', True, True)
            walk.assert_not_called()
        messages = list(task.result_queue.queue)
        self.assertTrue(any(search.tr("工程模式和排除规则需要 ripgrep，请恢复组件或使用全量搜索")
                            in message.get('text', '') for message in messages))

    def test_presets_roundtrip_scope_and_disable_everything(self):
        with patch.object(search, 'detect_everything', return_value='es.exe'):
            dialog = search.SearchDialog(str(self.root))
        self.addCleanup(dialog.deleteLater)
        self.addCleanup(dialog._release_search_resources)
        dialog.scope_combo.setCurrentIndex(1)
        dialog.exclusions_input.setText('vendor/;*.log')
        self.assertFalse(dialog.use_everything_cb.isChecked())
        self.assertFalse(dialog.use_everything_cb.isEnabled())
        preset = dialog._current_search_conditions()
        dialog.scope_combo.setCurrentIndex(0)
        dialog.exclusions_input.clear()
        self.assertTrue(dialog.use_everything_cb.isEnabled())
        dialog.apply_search_preset(preset)
        self.assertTrue(dialog.scope_combo.currentData())
        self.assertEqual(dialog.exclusions_input.text(), 'vendor/;*.log')


class ToolReadinessTests(QtTests):
    def setUp(self):
        self.old_cache = dict(tools._tool_readiness)
        tools._tool_readiness.clear()
        self.addCleanup(self.restore_cache)

    def restore_cache(self):
        tools._tool_readiness.clear()
        tools._tool_readiness.update(self.old_cache)

    def test_probe_deduplicates_programs_and_reports_failure(self):
        worker = tools.ToolProbeWorker(['Code.exe', 'Code.exe', 'WinMergeU.exe'])
        results = []
        worker.completed.connect(results.append)
        with patch.object(tools, 'resolve_program', side_effect=[r'C:\Code.exe', FileNotFoundError('missing')]) as resolve:
            worker.run()
        self.assertEqual(resolve.call_count, 2)
        self.assertEqual(results[0]['Code.exe'], (r'C:\Code.exe', ''))
        self.assertIn('missing', results[0]['WinMergeU.exe'][1])

    def test_per_user_winmerge_installation_is_detected(self):
        expected = r'C:\Local\Programs\WinMerge\WinMergeU.exe'
        with patch.dict(os.environ, {'LOCALAPPDATA': r'C:\Local'}), \
                patch.object(tools.shutil, 'which', return_value=None), \
                patch.object(tools.os.path, 'isfile', side_effect=lambda path: path == expected):
            self.assertEqual(tools.resolve_program('WinMergeU.exe'), expected)

    def test_dialog_detection_preview_and_download_do_not_launch_tools(self):
        def resolve(program):
            if program == 'Code.exe':
                return r'C:\VS Code\Code.exe'
            raise FileNotFoundError('missing WinMerge')

        with patch.object(tools, 'resolve_program', side_effect=resolve), \
                patch.object(tools.subprocess, 'Popen') as launch:
            dialog = tools.ContextToolsDialog(tools.default_tools(), context=tools.ToolContext(
                cwd=r'C:\project', selected=(r'C:\project\a & b.txt', r'C:\project\c.txt')))
            self.addCleanup(dialog.deleteLater)
            self.wait_for(lambda: dialog._probe_worker is None)
            self.assertIn(search.tr("已就绪"), dialog.list.item(0).text())
            self.assertIn(search.tr("未安装或未找到"), dialog.list.item(3).text())
            dialog.list.setCurrentRow(5)
            dialog.preview_arguments()
            self.assertIn('--diff', dialog.argument_preview.toPlainText())
            self.assertIn('a & b.txt', dialog.argument_preview.toPlainText())
            dialog.list.setCurrentRow(3)
            with patch.object(tools.QDesktopServices, 'openUrl', return_value=True) as open_url:
                dialog.open_download_page()
                self.assertEqual(open_url.call_args.args[0].toString(), 'https://winmerge.org/downloads/')
            dialog.program.setText(r'C:\moved\WinMergeU.exe')
            dialog.detect_tools()
            self.wait_for(lambda: dialog._probe_worker is None)
            self.assertIn(search.tr("路径失效"), dialog.list.item(3).text())
            launch.assert_not_called()


class CompareAndBasketTests(QtTests):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.left = self.root / 'left'
        self.right = self.root / 'right'
        self.target = self.root / 'target'
        for directory in (self.left, self.right, self.target):
            directory.mkdir()
        (self.left / 'same.txt').write_text('left')
        (self.right / 'same.txt').write_text('right')
        self.owner = QWidget()
        self.owner.config = {}
        self.owner.get_active_pane = lambda: types.SimpleNamespace(current_path=str(self.target))
        self.panel = FileTaskPanel(self.owner)
        self.owner.get_file_task_panel = lambda: self.panel
        self.owner.show_file_tasks = Mock()
        self.owner.add_paths_to_basket = Mock()
        self.addCleanup(self.owner.deleteLater)

    def test_compare_filters_and_opens_exact_file_pair_without_copying(self):
        dialog = DirectoryCompareDialog(self.owner, str(self.left), str(self.right))
        self.addCleanup(dialog.deleteLater)
        rows = [
            QTreeWidgetItem(['same.txt', search.tr('内容不同'), str(self.left / 'same.txt'), str(self.right / 'same.txt')]),
            QTreeWidgetItem(['only.txt', search.tr('仅左侧'), str(self.left / 'only.txt'), '']),
            QTreeWidgetItem(['bad.txt', 'read error', '', '']),
        ]
        dialog.table.addTopLevelItems(rows)
        dialog.table.setCurrentItem(rows[0])
        self.assertTrue(dialog.view_diff_button.isEnabled())
        with patch.object(tools, 'launch_context_tool') as launch:
            dialog.view_content_difference()
            tool, context = launch.call_args.args[1:]
            self.assertEqual(tool['id'], 'vscode-diff')
            self.assertEqual(context.selected, (str(self.left / 'same.txt'), str(self.right / 'same.txt')))
            self.assertEqual(self.panel.records, [])
        dialog.difference_filter.setCurrentIndex(dialog.difference_filter.findData('left'))
        self.assertTrue(rows[0].isHidden())
        self.assertFalse(rows[1].isHidden())
        self.assertEqual(dialog.table.selectedItems(), [])
        self.assertFalse(dialog.view_diff_button.isEnabled())
        dialog.difference_filter.setCurrentIndex(dialog.difference_filter.findData('other'))
        self.assertFalse(rows[2].isHidden())

    def test_task_panel_prunes_oldest_finished_records_only(self):
        self.panel.max_finished_records = 2
        records = []
        for index, done in enumerate((True, True, False, True)):
            row = QTreeWidgetItem([str(index)])
            self.panel.table.addTopLevelItem(row)
            records.append({'row': row, 'done': done, 'worker': None})
        self.panel.records = list(records)
        self.panel._prune_finished_records(keep=records[3])
        self.assertEqual(self.panel.records, [records[1], records[2], records[3]])
        self.assertEqual(self.panel.table.topLevelItemCount(), 3)

    def test_basket_deduplicates_references_and_removal_never_deletes_files(self):
        basket = FileBasketDialog(self.owner)
        paths = [str(self.left / 'same.txt'), str(self.right / 'same.txt')]
        self.assertEqual(basket.add_paths(paths + paths), 2)
        basket.table.setCurrentItem(basket.table.topLevelItem(0))
        self.assertEqual(basket.selected_or_all_paths(), paths[:1])
        basket.remove_selected()
        self.assertEqual(basket.all_paths(), paths[1:])
        basket.clear()
        self.assertEqual(basket.all_paths(), [])
        self.assertTrue(all(Path(path).exists() for path in paths))

    def test_basket_copy_reuses_tasks_preserves_collisions_and_reports_missing_source(self):
        basket = FileBasketDialog(self.owner)
        paths = [str(self.left / 'same.txt'), str(self.right / 'same.txt'), str(self.root / 'gone.txt')]
        basket.add_paths(paths)
        with patch.object(QMessageBox, 'question', return_value=QMessageBox.Yes):
            basket.copy_to_current()
        self.wait_for(lambda: self.panel.records and self.panel.records[-1]['done'])
        self.assertEqual({path.read_text() for path in self.target.iterdir()}, {'left', 'right'})
        self.assertIn(paths[-1], self.panel.records[-1]['failed'])
        self.assertEqual(basket.all_paths(), paths)
        self.owner.show_file_tasks.assert_called_once()

    def test_search_multi_selection_can_be_collected(self):
        dialog = search.SearchDialog(str(self.root), self.owner)
        self.addCleanup(dialog.deleteLater)
        self.addCleanup(dialog._release_search_resources)
        paths = [str(self.left / 'same.txt'), str(self.right / 'same.txt')]
        dialog._append_results_to_table([{'path': path, 'name': path} for path in paths])
        for row in range(2):
            dialog.result_list.selectionModel().select(
                dialog.result_model.index(row, 0), QItemSelectionModel.Select | QItemSelectionModel.Rows)
        dialog.basket_button.click()
        self.assertEqual(set(self.owner.add_paths_to_basket.call_args.args[0]), set(paths))

    def test_sorting_search_results_preserves_the_files_selected_for_collection(self):
        dialog = search.SearchDialog(str(self.root), self.owner)
        self.addCleanup(dialog.deleteLater)
        self.addCleanup(dialog._release_search_resources)
        rows = [{'path': str(self.root / name), 'name': name} for name in ('b.txt', 'a.txt', 'c.txt')]
        dialog._append_results_to_table(rows)
        dialog.result_list.selectionModel().select(
            dialog.result_model.index(2, 0), QItemSelectionModel.Select | QItemSelectionModel.Rows)
        before = set(dialog.selected_result_paths())
        dialog.result_model.sort(0, Qt.AscendingOrder)
        self.assertEqual(set(dialog.selected_result_paths()), before)
        dialog.basket_button.click()
        self.assertEqual(set(self.owner.add_paths_to_basket.call_args.args[0]), before)

    def test_destroyed_dialog_does_not_leave_a_stale_notification(self):
        owner = QWidget()
        toast = widgets.show_toast(owner, 'Test', 'Task completed')
        owner.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        self.assertNotIn(toast, widgets._active_toasts)
        widgets._position_toasts()

    def test_collection_layout_fits_both_languages_and_font_scales(self):
        old_font, old_language = self.qt.font(), i18n._app_language
        try:
            for language in ('zh', 'en'):
                for scale in (1, 1.5):
                    with self.subTest(language=language, scale=scale):
                        i18n._set_app_language(language)
                        self.qt.setFont(QFont('Segoe UI', round(9 * scale)))
                        basket = FileBasketDialog(self.owner)
                        width = 800 if scale == 1 else 1000
                        basket.resize(width, 600)
                        basket.show()
                        self.qt.processEvents()
                        try:
                            self.assertEqual(basket.width(), width)
                            for button in (basket.copy_button, basket.paths_button,
                                           basket.remove_button, basket.clear_button):
                                self.assertTrue(basket.rect().contains(button.geometry()))
                        finally:
                            basket.close()
                            basket.deleteLater()
        finally:
            i18n._set_app_language(old_language)
            self.qt.setFont(old_font)


class RuntimeIdentityTests(unittest.TestCase):
    def test_debug_profile_uses_the_same_pinned_runtime_as_start_script(self):
        root = Path(__file__).resolve().parents[1]
        launch = json.loads((root / '.vscode' / 'launch.json').read_text())
        profile = launch['configurations'][0]
        executable = profile['python'].replace('${workspaceFolder}', str(root))
        self.assertEqual(Path(executable).resolve(), Path(sys.executable).resolve())
        self.assertIn("'3.13.13'", (root / 'tools' / 'project_python.ps1').read_text(encoding='utf-8-sig'))
        output = subprocess.check_output(
            [executable, *profile['pythonArgs'], '-c',
             'import sys; print(sys.flags.ignore_environment,sys.flags.no_user_site,sys.flags.utf8_mode)'],
            text=True)
        self.assertEqual(output.strip(), '1 1 1')

    def test_about_identifies_mode_version_startup_and_interpreter(self):
        owner = types.SimpleNamespace(_app_started_at=0)
        with patch.object(sys, 'frozen', False, create=True):
            text = MainWindow.runtime_information(owner)
            self.assertIn(search.tr('源码运行'), text)
            self.assertIn(sys.executable, text)
            self.assertIn('1970', text)
        with patch.object(sys, 'frozen', True, create=True):
            self.assertIn(search.tr('打包 EXE'), MainWindow.runtime_information(owner))


if __name__ == '__main__':
    unittest.main()

class MenuLifecycleTests(QtTests):
    def test_temporary_menus_and_rebuilt_submenus_are_released(self):
        from PyQt5.QtCore import QPoint, QTimer
        from PyQt5.QtWidgets import QMenu, QMenuBar
        from tabexplorer.menus import clear_menu, exec_menu
        owner = QWidget()
        self.addCleanup(owner.deleteLater)
        for _ in range(3):
            menu = QMenu(owner)
            chosen = menu.addAction('a')
            QTimer.singleShot(20, lambda m=menu, a=chosen: (m.setActiveAction(a), m.close()))
            exec_menu(menu, QPoint(0, 0))
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        self.assertEqual(owner.findChildren(QMenu), [])
        bar = QMenuBar(owner)
        for _ in range(4):
            clear_menu(bar)
            bar.addMenu('folder').addMenu('nested')
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        self.assertEqual(len(bar.findChildren(QMenu)), 2)
        self.assertEqual(len(bar.actions()), 1)

    def test_newer_view_mode_schedule_cancels_stale_retries(self):
        from unittest.mock import call
        from tabexplorer.explorer_tab import FileExplorerTab
        host = types.SimpleNamespace(_is_cleaning_up=False,
                                     _set_ieb_current_view_mode=Mock(return_value='unchanged'))
        FileExplorerTab._schedule_view_mode(host, 4, (0, 20), 'details')
        FileExplorerTab._schedule_view_mode(host, 6, (5,), 'tile')
        QTest.qWait(80)
        self.assertEqual(host._set_ieb_current_view_mode.call_args_list, [call(6)])

    def test_minimize_pauses_visible_tabs_and_restore_resumes_only_current(self):
        def make_tab():
            tab = types.SimpleNamespace(_refresh_active=True)
            tab.set_refresh_active = Mock(side_effect=lambda active: setattr(tab, '_refresh_active', active))
            return tab
        left, right, other = make_tab(), make_tab(), make_tab()
        current = {'left': left, 'right': right}
        group = lambda key: (types.SimpleNamespace(currentIndex=lambda: 0),
                             types.SimpleNamespace(widget=lambda _index: current[key]))
        window = types.SimpleNamespace(_all_groups=lambda: [group('left'), group('right')])
        MainWindow._pause_visible_tabs_for_minimize(window)
        self.assertFalse(left._refresh_active or right._refresh_active)
        current['right'] = other
        MainWindow._resume_tabs_after_minimize(window)
        self.assertTrue(left._refresh_active)
        self.assertFalse(right._refresh_active)
        other.set_refresh_active.assert_not_called()
        self.assertEqual(window._minimize_paused_tabs, [])

class ChunkedContentScanTests(unittest.TestCase):
    def test_chunks_keep_boundary_matches_and_utf16_alignment(self):
        import re as _re
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'big.txt'
            text = 'x' * 900 + 'BEGIN-' + 'y' * 300 + '-END' + 'z' * 1000
            path.write_bytes(text.encode('utf-8'))
            overlap = search.content_chunk_overlap_bytes(r'BEGIN-y+-END', True)
            chunks = list(search._iter_decoded_chunks(str(path), 'utf-8', 1024, min(overlap, 512), 10 ** 6))
            self.assertTrue(any(_re.search(r'BEGIN-y+-END', chunk) for chunk in chunks))
            # 旧实现按关键词长度×2 重叠时，该跨块正则匹配会漏掉
            narrow = list(search._iter_decoded_chunks(str(path), 'utf-8', 1024, 24, 10 ** 6))
            self.assertFalse(any(_re.search(r'BEGIN-y+-END', chunk) for chunk in narrow))
            wide = Path(directory) / 'wide.txt'
            wide.write_bytes(('a' * 509 + '关键词' + 'b' * 600).encode('utf-16-le'))
            chunks = list(search._iter_decoded_chunks(str(wide), 'utf-16-le', 1024, 17, 10 ** 6))
            self.assertTrue(any('关键词' in chunk for chunk in chunks))
            limited = list(search._iter_decoded_chunks(str(wide), 'utf-16-le', 256, 16, 512))
            self.assertEqual(len(limited), 2)
        self.assertEqual(search.content_chunk_overlap_bytes('abc', False) % 2, 0)
        self.assertGreaterEqual(search.content_chunk_overlap_bytes('a', True), search.REGEX_CHUNK_OVERLAP_BYTES)

class ChatStreamingTests(unittest.TestCase):
    def test_streaming_http_error_closes_response(self):
        import requests
        from tabexplorer.chat import ChatWorker
        response = Mock()
        response.raise_for_status.side_effect = requests.exceptions.HTTPError('500')
        session = Mock()
        session.post.return_value = response
        worker = ChatWorker([], 'http://localhost', '', 'model', stream=True)
        with self.assertRaises(requests.exceptions.HTTPError):
            worker._request(session)
        response.close.assert_called_once_with()

class TabPageReleaseTests(QtTests):
    def test_removing_tabs_releases_placeholder_pages(self):
        from PyQt5.QtWidgets import QTabWidget
        from tabexplorer.mainwindow import _remove_tab_and_page
        tabs = QTabWidget()
        self.addCleanup(tabs.deleteLater)
        for index in range(5):
            tabs.addTab(QWidget(), str(index))
        for _ in range(4):
            _remove_tab_and_page(tabs, 0)
        QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
        self.assertEqual(tabs.count(), 1)
        self.assertEqual(tabs.tabText(0), '4')
        stacked = tabs.findChild(QWidget, 'qt_tabwidget_stackedwidget')
        self.assertEqual(len(stacked.findChildren(QWidget, options=Qt.FindDirectChildrenOnly)), 1)

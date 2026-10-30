"""Streaming searches, navigation snapshot generations and contextual tool integration."""

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import types
import unittest
from unittest.mock import Mock, patch

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QMenu

from tabexplorer import context_tools as tools
from tabexplorer import explorer_tab as explorer
from tabexplorer import search
from tabexplorer import i18n
from tabexplorer.mainwindow import MainWindow


class SnapshotTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def host(self):
        host = types.SimpleNamespace(
            current_path=r'C:\A', dir_poll_timer=Mock(isActive=Mock(return_value=False)),
            _snapshot_signals=Mock(), _snapshot_generation=0, _snapshot_inflight=False,
            _snapshot_pending=False, _refresh_active=True, _last_dir_snapshot=None,
            _is_slow_path=lambda path: path.startswith('\\\\'),
            _should_ignore_internal_dir_entry=lambda path, name: False, _request_refresh=Mock())
        host._request_dir_snapshot = lambda path: explorer.FileExplorerTab._request_dir_snapshot(host, path)
        return host

    def test_initial_snapshot_is_one_background_scan_without_ui_filesystem_calls(self):
        host = self.host()
        with patch.object(explorer.QThreadPool, 'globalInstance') as pool, \
                patch.object(explorer.os, 'scandir', side_effect=AssertionError('UI scandir')), \
                patch.object(explorer.os, 'stat', side_effect=AssertionError('UI stat')):
            explorer.FileExplorerTab._update_dir_polling(host, host.current_path)
            pool.return_value.start.assert_called_once()
            runnable = pool.return_value.start.call_args.args[0]
            self.assertEqual(runnable._generation, host._snapshot_generation)
        with patch.object(explorer, '_compute_dir_snapshot', return_value=(1, 2, 3, 4)) as compute:
            runnable.run()
            compute.assert_called_once()
            host._snapshot_signals.done.emit.assert_called_once_with(host.current_path, (1, 2, 3, 4), 1)
        explorer.FileExplorerTab._on_dir_snapshot_ready(host, host.current_path, (1, 2, 3, 4), 1)
        host._request_refresh.assert_not_called()
        explorer.FileExplorerTab._on_dir_snapshot_ready(host, host.current_path, (1, 3, 3, 4), 1)
        host._request_refresh.assert_called_once_with(reason='poll')

    def test_a_b_a_coalesces_pending_navigation_and_rejects_old_a(self):
        host = self.host()
        with patch.object(explorer.QThreadPool, 'globalInstance') as pool:
            for path in (r'C:\A', r'C:\B', r'C:\A'):
                host.current_path = path
                explorer.FileExplorerTab._update_dir_polling(host, path)
            self.assertEqual(pool.return_value.start.call_count, 1)
            explorer.FileExplorerTab._on_dir_snapshot_ready(host, r'C:\A', (99,), 1)
            self.assertIsNone(host._last_dir_snapshot)
            self.assertEqual(pool.return_value.start.call_count, 2)
            self.assertEqual(pool.return_value.start.call_args.args[0]._generation, 3)
            explorer.FileExplorerTab._on_dir_snapshot_ready(host, r'C:\A', (3,), 3)
            self.assertEqual(host._last_dir_snapshot, (3,))
            host._request_refresh.assert_not_called()

    def test_slow_shell_and_cleanup_paths_do_not_schedule(self):
        for path in ('', 'shell:MyComputerFolder', r'\\server\share'):
            host = self.host()
            host.current_path = path
            with patch.object(explorer.QThreadPool, 'globalInstance') as pool:
                explorer.FileExplorerTab._update_dir_polling(host, path)
                pool.assert_not_called()
        host._is_cleaning_up = True
        host._snapshot_pending = True
        host._request_dir_snapshot = Mock()
        explorer.FileExplorerTab._on_dir_snapshot_ready(host, 'old', (1,), 1)
        host._request_dir_snapshot.assert_not_called()


class StreamingSearchTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.rg = search.detect_ripgrep()
        if not self.rg:
            self.skipTest('Bundled ripgrep is required for these integration tests')

    def run_search(self, keyword='needle', filename=False, limit=100, **options):
        task = search._SearchTask(str(self.root), None, limit, self.rg)
        search.SearchDialog.do_search(task, keyword, filename, True, **options)
        messages = []
        while not task.result_queue.empty():
            messages.append(task.result_queue.get_nowait())
        rows = [row for message in messages if message['type'] == 'result_batch' for row in message['items']]
        return rows, messages

    def test_content_only_streams_without_second_directory_walk(self):
        (self.root / 'one.txt').write_text('first\nprefix needle suffix\n', encoding='utf-8')
        (self.root / 'none.txt').write_text('unrelated', encoding='utf-8')
        with patch.object(search.os, 'walk', side_effect=AssertionError('second traversal')):
            rows, _messages = self.run_search()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['line'], 2)
        self.assertEqual(rows[0]['column'], 8)
        self.assertEqual(rows[0]['preview'], 'prefix needle suffix')

    def test_combined_search_deduplicates_filename_and_content_hits(self):
        (self.root / 'needle.txt').write_text('needle', encoding='utf-8')
        (self.root / 'needle-name.txt').write_text('unrelated', encoding='utf-8')
        (self.root / 'content.txt').write_text('needle', encoding='utf-8')
        rows, _messages = self.run_search(filename=True)
        self.assertEqual(len(rows), 3)
        self.assertEqual(len({row['path'] for row in rows}), 3)

    def test_encoding_filters_and_actual_line_numbers(self):
        for encoding in ('utf-8', 'gbk', 'utf-16'):
            (self.root / f'{encoding}.txt').write_bytes('第一行\n中文 测试 后文\n'.encode(encoding))
        (self.root / 'binary.exe').write_bytes('测试'.encode('utf-8'))
        rows, _messages = self.run_search('测试', file_types='txt')
        self.assertEqual({Path(row['path']).stem for row in rows}, {'utf-8', 'gbk', 'utf-16'})
        self.assertTrue(all(row['line'] == 2 and row['column'] == 4 for row in rows))
        self.assertTrue(all('测试' in row['preview'] for row in rows))
        rows, _messages = self.run_search('测试', filters={'min_size': 100000})
        self.assertEqual(rows, [])

    def test_ascii_match_in_gbk_reports_decoded_column(self):
        (self.root / 'gbk.txt').write_bytes('中文 needle 后文\n'.encode('gbk'))
        rows, _messages = self.run_search()
        self.assertEqual(rows[0]['line'], 1)
        self.assertEqual(rows[0]['column'], 4)
        self.assertEqual(rows[0]['preview'], '中文 needle 后文')

    def test_result_limit_stops_engine_and_does_not_cache_partial_results(self):
        for index in range(10):
            (self.root / f'{index}.txt').write_text('needle', encoding='utf-8')
        search._search_cache.clear()
        rows, messages = self.run_search(limit=2, cache_key='limited-stream')
        self.assertEqual(len(rows), 2)
        self.assertIsNone(search._search_cache.get('limited-stream'))
        self.assertTrue(any(search.tr("（已达结果上限）") in message.get('text', '') for message in messages))
        self.assertFalse(any(thread.name == 'rg-output' for thread in threading.enumerate()))

    def test_long_line_keeps_match_but_bounds_preview(self):
        (self.root / 'long.txt').write_bytes(b'x' * (2 * 1024 * 1024) + b'needle\n')
        rows, _messages = self.run_search()
        self.assertEqual(len(rows), 1)
        self.assertNotIn('line', rows[0])
        self.assertLess(len(rows[0]['preview']), 200)

    def test_failure_after_partial_results_falls_back_without_duplicates(self):
        (self.root / 'first.txt').write_text('needle', encoding='utf-8')
        (self.root / 'second.txt').write_text('needle', encoding='utf-8')

        def partial(*args, **kwargs):
            yield {'path': str(self.root / 'first.txt'), 'line': 1, 'preview': 'needle'}
            raise OSError('engine failed')

        with patch.object(search, 'iter_ripgrep_matches', side_effect=partial):
            rows, messages = self.run_search()
        self.assertEqual(len(rows), 2)
        self.assertTrue(any('engine failed' in message.get('text', '') for message in messages))

    def test_first_result_arrives_before_exit_and_cancel_reaps_child(self):
        task = search._SearchTask(str(self.root), None, 10, self.rg)
        event = {'type': 'match', 'data': {'path': {'text': str(self.root / 'one.txt')},
                 'lines': {'text': 'needle\n'}, 'line_number': 1, 'submatches': [{'start': 0}]}}
        script = f'import time; print({json.dumps(json.dumps(event))}, flush=True); time.sleep(20)'
        original_popen = subprocess.Popen
        children = []

        def spawn(_command, **kwargs):
            child = original_popen([sys.executable, '-u', '-c', script], **kwargs)
            children.append(child)
            return child

        with patch.object(subprocess, 'Popen', side_effect=spawn):
            matches = search.iter_ripgrep_matches(task, 'needle', str(self.root))
            try:
                self.assertEqual(next(matches)['line'], 1)
                self.assertIsNone(children[0].poll())
                task.cancel_event.set()
                with self.assertRaises(search._SearchCancelled):
                    next(matches)
            finally:
                matches.close()
        self.assertIsNotNone(children[0].poll())

    def test_cancel_before_any_output_does_not_wait_for_pipe_read(self):
        task = search._SearchTask(str(self.root), None, 10, self.rg)
        original_popen = subprocess.Popen
        spawned = threading.Event()
        children = []
        failures = []

        def spawn(_command, **kwargs):
            child = original_popen([sys.executable, '-c', 'import time; time.sleep(20)'], **kwargs)
            children.append(child)
            spawned.set()
            return child

        def consume():
            try:
                list(search.iter_ripgrep_matches(task, 'needle', str(self.root)))
            except search._SearchCancelled:
                failures.append('cancelled')

        with patch.object(subprocess, 'Popen', side_effect=spawn):
            thread = threading.Thread(target=consume)
            thread.start()
            try:
                self.assertTrue(spawned.wait(5))
            finally:
                task.cancel_event.set()
                thread.join(5)
        self.assertFalse(thread.is_alive())
        self.assertEqual(failures, ['cancelled'])
        self.assertIsNotNone(children[0].poll())

    def test_cancellation_unblocks_full_result_queue(self):
        for index in range(20):
            (self.root / f'{index}.txt').write_text('needle', encoding='utf-8')
        with patch.object(search, 'SEARCH_RESULT_QUEUE_MAXSIZE', 1):
            task = search._SearchTask(str(self.root), None, 100, self.rg)
        worker = threading.Thread(target=task.run, args=('needle', False, True))
        worker.start()
        try:
            task.result_queue.get(timeout=5)  # initial status
            first = task.result_queue.get(timeout=5)
            self.assertEqual(first['type'], 'result_batch')
        finally:
            task.cancel_event.set()
            worker.join(5)
        self.assertFalse(worker.is_alive())
        self.assertFalse(any(thread.name == 'rg-output' for thread in threading.enumerate()))


class ContextToolTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def test_expansion_preserves_spaces_metacharacters_and_multiple_arguments(self):
        selected = (r'C:\folder space\a & {line}.txt', r'C:\folder space\b.txt')
        context = tools.ToolContext(cwd=r'C:\folder space', selected=selected,
                                    left=r'C:\left', right=r'C:\right')
        tool = tools.default_tools()[1]
        args, cwd = tools.prepare_tool(tool, context)
        self.assertEqual(args, ['Code.exe', '--reuse-window', *selected])
        self.assertEqual(cwd, r'C:\folder space')
        args, _cwd = tools.prepare_tool(tools.default_tools()[3], context)
        self.assertEqual(args[-2:], list(selected))
        args, _cwd = tools.prepare_tool(tools.default_tools()[4], context)
        self.assertEqual(args[-2:], [r'C:\left', r'C:\right'])

    def test_missing_context_and_invalid_templates_are_explicit_errors(self):
        for index in (0, 1, 2, 3, 4):
            with self.subTest(preset=index), self.assertRaises(ValueError):
                tools.prepare_tool(tools.default_tools()[index], tools.ToolContext())
        for template in ('{missing}', '{cwd.__class__}', '{file!r}', 'prefix{selected}', '{file:20}'):
            tool = dict(tools.default_tools()[0], args=[template])
            with self.subTest(template=template), self.assertRaises(ValueError):
                tools.validate_tool(tool)
        with self.assertRaises(ValueError):
            tools.validate_tool(dict(tools.default_tools()[0], program='tool.cmd'))

    def test_goto_uses_configured_program_and_exact_match_position(self):
        tool = dict(tools.default_tools()[2], program=r'C:\custom\Code.exe')
        args, _cwd = tools.prepare_tool(
            tool, tools.ToolContext(cwd=r'C:\project', selected=(r'C:\project\a.txt',), line=7, column=3))
        self.assertEqual(args, [r'C:\custom\Code.exe', '--reuse-window', '--goto', r'C:\project\a.txt:7:3'])

    def test_launch_uses_argv_and_propagates_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            context = tools.ToolContext(cwd=directory)
            worker = tools.ToolLaunchWorker(tools.default_tools()[0], context)
            errors = []
            worker.completed.connect(errors.append)
            with patch.object(tools, 'resolve_program', return_value=r'C:\Code.exe'), \
                    patch.object(tools.subprocess, 'Popen') as launch:
                worker.run()
                self.assertEqual(errors, [''])
                self.assertEqual(launch.call_args.args[0], [r'C:\Code.exe', '--reuse-window', directory])
                self.assertIs(launch.call_args.kwargs['shell'], False)
            with patch.object(tools, 'resolve_program', side_effect=FileNotFoundError('missing tool')), \
                    patch.object(tools.subprocess, 'Popen') as launch:
                worker.run()
                self.assertIn('missing tool', errors[-1])
                launch.assert_not_called()

    def test_real_executable_receives_arguments_without_shell_interpretation(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            selected = (str(root / 'space & {file}.txt'), str(root / '中文.txt'))
            for path in selected:
                Path(path).touch()
            output = root / 'arguments.json'
            tool = {'name': 'Argument probe', 'program': sys.executable, 'cwd': '{cwd}',
                    'args': ['-c', "import json,sys; json.dump(sys.argv[2:],open(sys.argv[1],'w'))",
                             str(output), '{selected}']}
            worker = tools.ToolLaunchWorker(tool, tools.ToolContext(cwd=directory, selected=selected))
            errors = []
            worker.completed.connect(errors.append)
            original_popen = subprocess.Popen
            children = []

            def spawn(*args, **kwargs):
                child = original_popen(*args, **kwargs)
                children.append(child)
                return child

            with patch.object(tools.subprocess, 'Popen', side_effect=spawn):
                worker.run()
            self.assertEqual(errors, [''])
            try:
                self.assertEqual(children[0].wait(timeout=10), 0)
                self.assertEqual(json.loads(output.read_text()), list(selected))
            finally:
                for child in children:
                    if child.poll() is None:
                        child.kill()
                        child.wait()

    def test_invalid_saved_tool_configuration_is_not_silently_replaced(self):
        for value in (None, {}, [dict(tools.default_tools()[0], args='bad')]):
            with self.subTest(value=value), self.assertRaises(ValueError):
                tools.configured_tools({'external_tools': value})

    def test_menu_disables_missing_context_and_dispatches_captured_selection(self):
        menu = QMenu()
        self.addCleanup(menu.deleteLater)
        context = tools.ToolContext(cwd=r'C:\project')
        with patch.object(tools, 'launch_context_tool') as launch:
            tools.populate_tools_menu(menu, tools.default_tools(), context, menu)
            self.assertTrue(menu.actions()[0].isEnabled())
            self.assertFalse(menu.actions()[3].isEnabled())
            self.assertTrue(menu.actions()[3].property('unavailable_reason'))
            menu.actions()[0].trigger()
            self.assertEqual(launch.call_args.args[2], context)

    def test_main_window_context_uses_active_side_without_swapping_split_paths(self):
        left = types.SimpleNamespace(current_path=r'C:\left')
        right = types.SimpleNamespace(current_path=r'C:\right',
                                      _get_selected_paths=lambda: [r'C:\right\selected.txt'])
        owner = types.SimpleNamespace(
            _split_active=True, get_active_pane=lambda: right,
            get_current_tab_widget=lambda: left, _get_split_pane=lambda: right)
        context = MainWindow._external_tool_context(owner)
        self.assertEqual(context.cwd, r'C:\right')
        self.assertEqual(context.selected, (r'C:\right\selected.txt',))
        self.assertEqual((context.left, context.right), (r'C:\left', r'C:\right'))

    def test_tool_editor_layout_fits_languages_and_font_scales(self):
        old_font, old_language = self.qt.font(), i18n._app_language
        try:
            for language in ('zh', 'en'):
                for scale in (1, 1.5):
                    with self.subTest(language=language, scale=scale):
                        i18n._set_app_language(language)
                        self.qt.setFont(QFont('Segoe UI', round(9 * scale)))
                        dialog = tools.ContextToolsDialog(tools.default_tools())
                        try:
                            width = 800 if scale == 1 else 1000
                            dialog.resize(width, 760)
                            dialog.show()
                            self.qt.processEvents()
                            self.assertEqual(dialog.width(), width)
                            for widget in (dialog.list, dialog.name, dialog.program,
                                           dialog.arguments, dialog.cwd):
                                self.assertTrue(dialog.rect().contains(widget.geometry()))
                                self.assertGreater(widget.width(), 60, type(widget).__name__)
                        finally:
                            dialog.close()
                            dialog.deleteLater()
        finally:
            i18n._set_app_language(old_language)
            self.qt.setFont(old_font)

    def test_editor_cancel_is_isolated_and_accept_persists_arguments(self):
        original = tools.default_tools()
        dialog = tools.ContextToolsDialog(original)
        self.addCleanup(dialog.deleteLater)
        dialog.arguments.setPlainText('--reuse-window\n{selected}')
        dialog.reject()
        self.assertEqual(original[0]['args'], ['--reuse-window', '{cwd}'])
        dialog.accept()
        self.assertEqual(dialog.tools[0]['args'], ['--reuse-window', '{selected}'])
        self.assertEqual(tools.configured_tools({'external_tools': dialog.tools}), dialog.tools)
        dialog.remove_tool()
        dialog.add_presets()
        self.assertEqual(dialog.list.count(), len(original))

    def test_preview_and_editor_use_row_metadata_without_reading_file_on_ui(self):
        dialog = search.SearchDialog(r'C:\project')
        self.addCleanup(dialog.deleteLater)
        self.addCleanup(dialog._release_search_resources)
        with patch.object(tools, 'launch_context_tool') as launch:
            dialog._append_results_to_table([
                {'path': r'C:\project\a.txt', 'name': 'a', 'line': 5, 'column': 2, 'preview': 'needle'}])
            self.assertIn('needle', dialog.match_preview.toPlainText())
            self.assertTrue(dialog.open_match_button.isEnabled())
            dialog.open_match_button.click()
            self.assertEqual(launch.call_args.args[2].line, 5)
            dialog.result_model.clear()
            self.assertFalse(dialog.open_match_button.isEnabled())
            self.assertEqual(dialog.match_preview.toPlainText(), '')


if __name__ == '__main__':
    unittest.main()

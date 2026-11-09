"""Interaction tests with real Qt widgets and isolated application data."""

from contextlib import ExitStack
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QAction, QApplication, QMenu, QWidget

from app_modules import app, patch_all
from tabexplorer.commands import CommandPalette, menu_commands


class CommandPaletteTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def test_filter_keyboard_activation_and_disabled_explanation(self):
        dialog = CommandPalette()
        self.addCleanup(dialog.deleteLater)
        sync = QAction('Sync folders', dialog)
        sync.setEnabled(False)
        sync.setProperty('unavailable_reason', 'Open split view first')
        search = QAction('Search (Ctrl+F)', dialog)
        callback = Mock()
        search.triggered.connect(callback)
        dialog.set_commands([('Sync folders', sync), ('Search (Ctrl+F)', search)])
        dialog.show()
        dialog.query.setText('sync')
        dialog.choose_current()
        self.assertIsNone(dialog.selected_action)
        self.assertEqual(dialog.explanation.text(), 'Open split view first')
        dialog.query.setText('ctrl+f')
        QTest.keyClick(dialog.query, Qt.Key_Return)
        self.assertIs(dialog.selected_action, search)
        callback.assert_not_called()
        dialog.selected_action.trigger()
        callback.assert_called_once()

    def test_menu_recursion_and_checkable_action_use_original_action(self):
        menu = QMenu()
        self.addCleanup(menu.deleteLater)
        submenu = menu.addMenu('Workspace')
        action = submenu.addAction('Save')
        hidden = menu.addAction('Hidden')
        hidden.setVisible(False)
        menu.addSeparator()
        sync = menu.addAction('Sync')
        sync.setCheckable(True)
        entries = list(menu_commands(menu))
        self.assertEqual(entries, [('Workspace / Save', action), ('Sync', sync)])
        dialog = CommandPalette()
        self.addCleanup(dialog.deleteLater)
        dialog.set_commands(entries)
        QTest.keyClick(dialog.query, Qt.Key_Down)
        dialog.choose_current()
        dialog.selected_action.trigger()
        self.assertTrue(sync.isChecked())
        dialog.query.setText('nothing-matches')
        self.assertEqual(dialog.results.topLevelItemCount(), 0)
        self.assertTrue(dialog.explanation.text())

    def test_new_default_hotkey_yields_to_existing_user_binding(self):
        """新增命令的默认键（Ctrl+P）若已被用户分配给其他功能则让位；用户显式设置的按键保持不变。"""
        from tabexplorer.hotkeys import _hotkey_bindings
        self.assertEqual(_hotkey_bindings({})['quick_open'], 'Ctrl+P')
        bindings = _hotkey_bindings({'hotkey_bindings': {'frequent_dirs': 'Ctrl+P'}})
        self.assertEqual((bindings['frequent_dirs'], bindings['quick_open']), ('Ctrl+P', ''))
        bindings = _hotkey_bindings({'hotkey_bindings': {'quick_open': 'Ctrl+Alt+O'}})
        self.assertEqual(bindings['quick_open'], 'Ctrl+Alt+O')

    def test_fuzzy_match_highlights_words_and_prefers_contiguous_text(self):
        """连续子串优先于分散字符，词首加分；只在路径等附加文本中命中时可用但不高亮；任一词不中即排除。"""
        from tabexplorer.commands import fuzzy_match, split_hotkey
        self.assertEqual(fuzzy_match('sw', 'Save Workspace')[1], [0, 5])
        self.assertGreater(fuzzy_match('work', 'Save Workspace')[0], fuzzy_match('wsp', 'Save Workspace')[0])
        self.assertGreater(fuzzy_match('space', 'Space bar')[0], fuzzy_match('space', 'Workspace')[0])
        self.assertEqual(fuzzy_match('分屏 比较', '分屏对比与目录比较')[1], [0, 1, 7, 8])
        self.assertEqual(fuzzy_match('src', 'project', r'C:\work\src')[1], [])
        self.assertIsNone(fuzzy_match('xyz', 'Save Workspace'))
        self.assertEqual(split_hotkey('查找功能 (Ctrl+Shift+P)'), ('查找功能', 'Ctrl+Shift+P'))
        self.assertEqual(split_hotkey('目录差异比较...'), ('目录差异比较...', ''))

    def test_recent_commands_come_first_with_group_labels(self):
        """最近使用的命令排在最前并标出分组，其余命令保持原顺序；右侧显示从标签拆出的快捷键。"""
        from tabexplorer.commands import _ROLE_GROUP, _ROLE_KEYS
        dialog = CommandPalette()
        self.addCleanup(dialog.deleteLater)
        actions = [QAction(text, dialog) for text in ('Alpha', 'Beta (Ctrl+B)', 'Gamma')]
        dialog.set_commands([(action.text(), action) for action in actions], recent=['Gamma', 'Beta'])
        rows = [dialog.results.topLevelItem(index) for index in range(dialog.results.topLevelItemCount())]
        self.assertEqual([row.text(0) for row in rows], ['Gamma', 'Beta', 'Alpha'])
        groups = [row.data(0, _ROLE_GROUP) for row in rows]
        self.assertEqual(groups, [app.tr('最近使用'), None, app.tr('其他命令')])
        self.assertEqual(rows[1].data(0, _ROLE_KEYS), 'Ctrl+B')
        dialog.query.setText('gam')
        self.assertEqual(dialog.results.topLevelItem(0).text(0), 'Gamma')
        self.assertEqual(dialog.results.topLevelItemCount(), 1)

    def test_quick_open_prefixes_switch_between_files_commands_and_help(self):
        """快速打开：不带前缀查找标签/书签/目录，> 切到命令，? 显示用法并可一键切换；输入路径时首项为直接打开。"""
        opened = []
        entries = [{'group': 'tabs', 'label': 'src', 'description': r'C:\work\src', 'run': lambda: opened.append('src')},
                   {'group': 'bookmarks', 'label': 'Docs', 'description': r'D:\docs', 'run': lambda: opened.append('docs')}]
        dialog = CommandPalette(quick_open=lambda: entries, open_path=opened.append)
        self.addCleanup(dialog.deleteLater)
        command = QAction('Compare folders', dialog)
        dialog.set_commands([('Compare folders', command)])
        self.assertEqual(dialog.mode, 'files')
        self.assertEqual(dialog.results.topLevelItemCount(), 2)
        dialog.query.setText('docs')
        self.assertEqual(dialog.results.topLevelItem(0).text(0), 'Docs')
        dialog.choose_current()
        self.assertIs(dialog.selected_entry, entries[1])
        dialog.query.setText('>comp')
        self.assertEqual(dialog.mode, 'commands')
        self.assertEqual(dialog.results.topLevelItem(0).text(0), 'Compare folders')
        dialog.query.setText('?')
        self.assertEqual(dialog.mode, 'help')
        dialog.results.setCurrentItem(dialog.results.topLevelItem(1))
        dialog.choose_current()
        self.assertEqual(dialog.query.text(), '>')
        self.assertEqual(dialog.mode, 'commands')
        QTest.keyClick(dialog.query, Qt.Key_P, Qt.ControlModifier)
        self.assertEqual(dialog.mode, 'files')
        dialog.query.setText(r'C:\Windows')
        self.assertEqual(dialog.results.topLevelItem(0).text(0), app.tr('打开路径'))
        dialog.choose_current()
        dialog.selected_entry['run']()
        self.assertEqual(opened, [r'C:\Windows'])

    def test_theme_choices_preview_on_move_and_start_at_current_theme(self):
        """主题选择从当前主题开始；上下移动后预览所选主题，筛选只按主题名称。"""
        previewed = []
        dialog = CommandPalette()
        self.addCleanup(dialog.deleteLater)
        entries = [{'id': 'light', 'label': 'Light Modern', 'group': 'light'},
                   {'id': 'dark', 'label': 'Dark Modern', 'group': 'dark'},
                   {'id': 'monokai', 'label': 'Monokai', 'group': 'dark'}]
        dialog.set_theme_choices(entries, 'dark', previewed.append)
        dialog.show()
        self.assertEqual(dialog.mode, 'themes')
        self.assertEqual(dialog.results.currentItem().text(0), 'Dark Modern')
        QTest.keyClick(dialog.query, Qt.Key_Down)
        dialog._preview_current()
        self.assertEqual(previewed[-1], 'monokai')
        dialog.query.setText('light')
        self.assertEqual(dialog.results.topLevelItemCount(), 1)
        dialog.choose_current()
        self.assertEqual(dialog.selected_entry['id'], 'light')


class ContextControlsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        directory = self.stack.enter_context(tempfile.TemporaryDirectory())
        self.config = {'enable_explorer_monitor': False, 'enable_title_shortcuts': False,
                       'title_shortcuts': [], 'ai_chat': {'enabled': False, 'panel_visible': False},
                       'hotkeys': {'screenshot': False}}
        self.stack.enter_context(patch.object(app.MainWindow, 'load_config', return_value=self.config))
        for name in ('ensure_default_bookmarks', '_delayed_initialization', '_check_shortcuts',
                     '_start_shortcut_listener', '_maybe_auto_check_updates', 'save_config',
                     'save_session_snapshot', 'save_pinned_tabs', '_setup_window_icon', 'add_new_tab'):
            self.stack.enter_context(patch.object(app.MainWindow, name))
        manager = self.stack.enter_context(patch_all('BookmarkManager'))
        manager.return_value.get_tree.return_value = {
            'bookmark_bar': {'type': 'folder', 'children': [], 'id': '1'},
            'other': {'type': 'folder', 'children': [], 'id': '2'}}
        self.stack.enter_context(patch_all('get_app_data_path',
                                         side_effect=lambda *parts: os.path.join(directory, *parts)))
        self.window = app.MainWindow()
        self.window.server_socket = Mock()
        self.addCleanup(self.dispose_window)
        self.panes = []
        for tabs, content in self.window._all_groups():
            pane = QWidget()
            pane.current_path = directory
            pane.can_go_back = lambda: True
            pane.can_go_forward = lambda: False
            pane.path_bar = app.SimplePathBar(pane)
            content.addWidget(pane)
            tabs.addTab(QWidget(), 'folder')
            self.panes.append(pane)
        self.window.resize(1050, 650)
        self.window.show()
        self.qt.processEvents()

    def dispose_window(self):
        self.window.close()
        self.window.deleteLater()
        self.qt.processEvents()

    def test_split_action_state_scope_and_disable_reason(self):
        w = self.window
        w._refresh_context_actions()
        self.assertTrue(w.split_controls.isHidden())
        self.assertFalse(w.sync_browse_action.isEnabled())
        self.assertTrue(w.sync_browse_action.property('unavailable_reason'))
        w._activate_split_layout()
        w._active_pane = self.panes[1]
        w.update_navigation_buttons()
        self.assertFalse(w.split_controls.isHidden())
        self.assertTrue(w.sync_browse_action.isEnabled())
        self.assertIn(app.tr('右侧'), w.search_button.toolTip())
        self.assertIn(app.tr('右侧'), w.cmd_button.toolTip())
        w.sync_browse_button.click()
        self.assertTrue(w._sync_browse)
        self.assertTrue(w.sync_browse_action.isChecked())
        self.assertEqual(w.sync_browse_button.text(), app.tr('同步导航：开'))
        w.sync_browse_action.trigger()
        self.assertFalse(w.sync_browse_button.isChecked())
        self.panes[1].current_path = 'shell:MyComputerFolder'
        w._refresh_context_actions()
        self.assertFalse(w.search_button.isEnabled())
        self.assertTrue(w.search_button.property('unavailable_reason'))
        w._teardown_split_group()
        self.assertTrue(w.split_controls.isHidden())
        self.assertFalse(w._sync_browse)

    def test_command_shortcut_and_menu_follow_custom_binding(self):
        w = self.window
        self.config['hotkey_bindings'] = {'command_palette': 'Ctrl+Alt+P', 'frequent_dirs': 'Ctrl+Alt+J'}
        w._apply_hotkey_bindings()
        self.assertIn('Ctrl+Alt+P', w.command_palette_action.text())
        self.assertIn('Ctrl+Alt+J', w.frequent_dirs_action.text())
        with patch.object(w, 'show_command_palette') as show:
            w.command_palette_action.trigger()
            show.assert_called_once()
            show.reset_mock()
            self.assertTrue(w._run_shortcut(ord('P'), True, False, True))
            show.assert_called_once()
        self.config['hotkeys']['command_palette'] = False
        self.config['hotkeys']['search'] = False
        w._apply_hotkey_bindings()
        self.assertNotIn('Ctrl+Alt+P', w.command_palette_action.text())
        self.assertNotIn('Ctrl+F', w.search_button.toolTip())

    def test_palette_invokes_existing_menu_action_once(self):
        action = self.window.directory_compare_action
        action.triggered.disconnect()
        callback = Mock()
        action.triggered.connect(callback)

        def choose():
            dialog = next(widget for widget in QApplication.topLevelWidgets()
                          if isinstance(widget, CommandPalette) and widget.isVisible())
            dialog.query.setText('>' + app.tr('目录差异比较...'))
            dialog.choose_current()

        QTimer.singleShot(0, choose)
        self.window.show_command_palette()
        callback.assert_called_once()
        self.assertEqual(self.window.config['recent_commands'][0], app.tr('目录差异比较...'))

    def test_command_center_shows_folder_and_opens_quick_open(self):
        """命令中心显示活动标签的文件夹名，提示中带快捷键，点击打开快速打开。"""
        w = self.window
        w._refresh_context_actions()
        self.assertEqual(w.command_center._context, os.path.basename(self.panes[0].current_path))
        self.assertIn('Ctrl+P', w.command_center.toolTip())
        self.assertFalse(w.command_center.icon().isNull())
        with patch.object(w, 'show_quick_open') as show:
            w.command_center.click()
        show.assert_called_once_with()
        self.assertIn('Ctrl+P', w.quick_open_action.text())

    def test_quick_open_lists_each_path_once_and_activates_its_tab(self):
        """快速打开：同一路径只列一次（右侧标签成为别名），选中后切到该标签；最近关闭的标签可直接恢复。"""
        w = self.window
        w.closed_tabs_history = [{'path': r'C:\gone', 'title': 'gone', 'is_shell': False}]
        entries = w._quick_open_entries()
        tabs = [entry for entry in entries if entry['group'] == app.tr('打开的标签')]
        self.assertEqual(len(tabs), 1)
        self.assertIn(app.tr('右侧'), tabs[0]['aliases'])
        with patch.object(w, 'set_active_pane_to_group') as activate:
            tabs[0]['run']()
        activate.assert_called_once_with(w.tab_widget)
        closed = next(entry for entry in entries if entry['group'] == app.tr('最近关闭'))
        closed['run']()
        w.add_new_tab.assert_called_with(r'C:\gone', is_shell=False, target_tabwidget=w.get_active_group_tabwidget())
        self.assertEqual(w.closed_tabs_history, [])

    def test_color_theme_picker_previews_then_reverts_or_saves(self):
        """主题选择：上下键即时预览；Esc 恢复原主题且不保存，Enter 保存并保持所选主题。"""
        w = self.window
        native = patch_all('_set_native_app_mode')
        native.start()
        self.addCleanup(native.stop)
        self.addCleanup(app.set_theme, 'light')
        self.config['theme'] = 'light'
        w.apply_theme_config()
        previewed = []

        def step(final_key):
            palette = next(widget for widget in QApplication.topLevelWidgets()
                           if isinstance(widget, CommandPalette) and widget.isVisible())
            QTest.keyClick(palette.query, Qt.Key_Down)
            palette._preview_current()
            previewed.append(app.current_theme_id())
            QTest.keyClick(palette.query, final_key)

        QTimer.singleShot(0, lambda: step(Qt.Key_Escape))
        w.show_color_theme_picker()
        self.assertNotEqual(previewed[-1], 'light')
        self.assertEqual(app.current_theme_id(), 'light')
        self.assertEqual(self.config['theme'], 'light')
        QTimer.singleShot(0, lambda: step(Qt.Key_Return))
        w.show_color_theme_picker()
        self.assertEqual(self.config['theme'], previewed[-1])
        self.assertEqual(app.current_theme_id(), previewed[-1])

    def test_split_controls_fit_languages_and_font_scales(self):
        old_language, old_font = app._app_language, self.qt.font()
        output = Path(tempfile.gettempdir()) / 'tabex-usability-checks'
        output.mkdir(exist_ok=True)
        try:
            for language in ('zh', 'en'):
                for scale in (1, 1.5):
                    app._set_app_language(language)
                    self.window.setFont(QFont('Segoe UI', round(9 * scale)))
                    self.window._activate_split_layout()
                    for width in (800, 1280):
                        self.window.resize(width, 650)
                        self.qt.processEvents()
                        self.window._refresh_context_actions()
                        self.assertLess(self.window.split_controls.height(), 90)
                        for button in (self.window.sync_browse_button, self.window.compare_directories_button):
                            self.assertTrue(self.window.split_controls.rect().contains(button.geometry()))
                            self.assertGreaterEqual(button.width(),
                                                    button.fontMetrics().horizontalAdvance(button.text()) + 6)
                        self.assertEqual(self.window.workspace_tools_button.toolButtonStyle(),
                                         Qt.ToolButtonTextBesideIcon if width >= 900 else Qt.ToolButtonIconOnly)
                    self.assertTrue(self.window.grab().save(str(output / f'split-{language}-{scale}.png')))
        finally:
            app._set_app_language(old_language)
            self.qt.setFont(old_font)

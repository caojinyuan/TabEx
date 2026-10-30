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
            dialog = QApplication.activeModalWidget()
            self.assertIsInstance(dialog, CommandPalette)
            dialog.query.setText(app.tr('目录差异比较...'))
            dialog.choose_current()

        QTimer.singleShot(0, choose)
        self.window.show_command_palette()
        callback.assert_called_once()

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

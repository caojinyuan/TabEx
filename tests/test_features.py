"""本批优化与新功能的回归测试：证书校验、密钥加密、开机启动、书签、标签分组、常用目录、分屏同步、搜索筛选和翻译覆盖。

运行：python -m unittest discover -s tests -p test_features.py -v
覆盖边界：使用临时目录、轻量 Qt 控件和模拟对象；不访问网络，不修改用户配置、书签或启动文件夹。
"""

import ast
import json
import os
from pathlib import Path
import re
import ssl
import sys
import tempfile
import time
import types
import unittest
from unittest.mock import Mock, patch

from PyQt5.QtWidgets import QApplication, QStackedWidget, QTabWidget, QWidget

from app_modules import MODULE_FILES, ROOT, app, patch_all


class SecurityTests(unittest.TestCase):
    def test_http_session_verifies_certificates_unless_explicitly_insecure(self):
        """AI 请求默认校验证书和主机名；只有显式勾选“不受信任证书”时才关闭校验。"""
        from tabexplorer.system import create_http_session
        with create_http_session() as session:
            self.assertIsNot(session.verify, False)
            pool = session.get_adapter('https://example.invalid').poolmanager
            context = pool.connection_pool_kw['ssl_context']
            self.assertEqual(context.verify_mode, ssl.CERT_REQUIRED)
            self.assertTrue(context.check_hostname)
            self.assertGreater(context.cert_store_stats()['x509_ca'], 0)
        with create_http_session(allow_insecure=True) as session:
            self.assertIs(session.verify, False)

    def test_api_key_is_encrypted_on_disk_and_decrypted_on_load(self):
        """API 密钥以 DPAPI 密文写盘、读取时还原；内容未变时不重复写盘，旧版明文配置下次保存即迁移。"""
        from tabexplorer.persistence import ConfigStore
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.json'
            path.write_text(json.dumps({'ai_chat': {'api_key': 'sk-legacy-secret', 'model': 'm'}}), encoding='utf-8')
            store = ConfigStore(path)
            config = store.load()
            self.assertEqual(config['ai_chat']['api_key'], 'sk-legacy-secret')
            self.assertTrue(store.save(config))
            text = path.read_text(encoding='utf-8')
            self.assertNotIn('sk-legacy-secret', text)
            self.assertIn('api_key_protected', json.loads(text)['ai_chat'])
            self.assertEqual(config['ai_chat']['api_key'], 'sk-legacy-secret')
            mtime = path.stat().st_mtime_ns
            self.assertTrue(store.save(config))
            self.assertEqual(path.stat().st_mtime_ns, mtime)
            reloaded = ConfigStore(path).load()
            self.assertEqual(reloaded['ai_chat']['api_key'], 'sk-legacy-secret')
            self.assertNotIn('api_key_protected', reloaded['ai_chat'])

    def test_undecryptable_key_is_reported_and_cleared(self):
        """无法解密的密文（如来自其他用户）不导致加载失败，标记提示并清空密钥。"""
        from tabexplorer.persistence import ConfigStore
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'config.json'
            path.write_text(json.dumps({'ai_chat': {'api_key_protected': 'bm90LWRwYXBp'}}), encoding='utf-8')
            store = ConfigStore(path)
            config = store.load()
            self.assertTrue(store.secret_unreadable)
            self.assertEqual(config['ai_chat']['api_key'], '')

    def test_model_list_parsing_handles_openai_and_plain_lists(self):
        """模型列表兼容 {data:[{id}]}、{data:[{name}]} 与字符串列表，去重并忽略无效项。"""
        from tabexplorer.settings import parse_model_names
        self.assertEqual(parse_model_names({'data': [{'id': 'a'}, {'name': 'b'}, {'id': 'a'}, {}, 3]}), ['a', 'b'])
        self.assertEqual(parse_model_names(['x', 'y']), ['x', 'y'])
        self.assertEqual(parse_model_names({'error': 'denied'}), [])


class StartupShortcutTests(unittest.TestCase):
    def test_stale_startup_shortcut_is_retargeted_only_when_old_target_missing(self):
        """快捷方式指向的程序已不存在时改为当前程序；旧目标仍存在时不抢占。"""
        from tabexplorer.system import read_shortcut_target, repair_stale_startup_shortcut, write_startup_shortcut
        with tempfile.TemporaryDirectory() as directory:
            shortcut = os.path.join(directory, 'TabExplorer.lnk')
            current = sys.executable
            write_startup_shortcut(os.path.join(directory, 'moved', 'TabExplorer.exe'), shortcut)
            self.assertTrue(repair_stale_startup_shortcut(current, shortcut))
            self.assertEqual(os.path.normcase(read_shortcut_target(shortcut)), os.path.normcase(current))
            other = os.path.join(directory, 'other.exe')
            Path(other).write_bytes(b'')
            write_startup_shortcut(other, shortcut)
            self.assertFalse(repair_stale_startup_shortcut(current, shortcut))
            self.assertFalse(repair_stale_startup_shortcut(current, os.path.join(directory, 'missing.lnk')))


class BookmarkTests(unittest.TestCase):
    SPECS = [('computer', '🖥️', '此电脑', 'shell:MyComputerFolder'), ('downloads', '⬇️', '下载', r'C:\Users\u\Downloads')]

    def manager(self, children):
        manager = app.BookmarkManager.__new__(app.BookmarkManager)
        manager.bookmark_tree = {'bookmark_bar': {'id': 'bar', 'type': 'folder', 'name': 'bar', 'children': children}}
        manager.save_bookmarks = Mock()
        return manager

    def test_builtin_bookmarks_migrate_legacy_items_and_keep_user_bookmarks(self):
        """旧版内置项按图标与 URL 识别并保留 ID；同名但无图标的用户书签不再被删除，重复整理无改动。"""
        user_named_like_builtin = {'id': 'u1', 'type': 'url', 'name': '下载', 'url': r'D:\my-downloads'}
        user_rocket = {'id': 'u2', 'type': 'url', 'name': '🚀 tools', 'url': r'D:\tools'}
        legacy = {'id': 'old', 'type': 'url', 'name': '🖥️ 此电脑', 'url': 'shell:MyComputerFolder'}
        manager = self.manager([user_named_like_builtin, legacy, user_rocket])
        self.assertTrue(manager.ensure_builtin_bookmarks(self.SPECS))
        children = manager.bookmark_tree['bookmark_bar']['children']
        self.assertEqual([node.get('builtin_key') for node in children[:2]], ['computer', 'downloads'])
        self.assertEqual(children[0]['id'], 'old')
        self.assertEqual(children[2:], [user_named_like_builtin, user_rocket])
        self.assertTrue(app.is_builtin_bookmark(children[0]))
        self.assertFalse(app.is_builtin_bookmark(user_rocket))
        self.assertFalse(manager.ensure_builtin_bookmarks(self.SPECS))

    def test_add_folder_creates_child_with_unique_id(self):
        """在指定文件夹下新建子文件夹并返回新 ID；父文件夹不存在时返回 None。"""
        manager = self.manager([])
        folder_id = manager.add_folder('bar', 'Projects')
        children = manager.bookmark_tree['bookmark_bar']['children']
        self.assertEqual((children[0]['id'], children[0]['type'], children[0]['name']), (folder_id, 'folder', 'Projects'))
        self.assertIsNone(manager.add_folder('missing', 'x'))


class TabFeatureTests(unittest.TestCase):
    """标签右键操作与分组命名/折叠。使用真实 Qt 标签栏和内容栈，标签内容用轻量控件代替 Explorer。"""

    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def make_tab(self, path, pinned=False, color='', name='', collapsed=False):
        tab = QWidget()
        tab.current_path = path
        tab.is_pinned = pinned
        tab.bookmark_group_color = color
        tab.tab_group_name = name
        tab.tab_group_collapsed = collapsed
        tab.update_tab_title = Mock()
        return tab

    def host(self, tabs):
        window = app.MainWindow
        names = ('_resolve_group', '_content_stack_for', '_closable_tabs', '_close_tab_widgets', 'close_other_tabs',
                 'close_tabs_to_right', '_tab_group_range', '_tab_group_members', 'set_tab_group_collapsed',
                 'toggle_tab_group_collapsed', '_after_tab_group_change', '_apply_tab_grouping_for_pane',
                 '_apply_tab_group_color', '_reveal_hidden_current_tab')
        host = type('TabHost', (QWidget,), {name: getattr(window, name) for name in names})()
        self.addCleanup(host.deleteLater)
        host.config = {'show_tab_group_markers': True, 'hotkey_bindings': {}}
        host.tab_widget, host.content_stack = QTabWidget(host), QStackedWidget(host)
        bar = app.CustomTabBar()
        bar.main_window = host
        bar.owner_tabwidget = host.tab_widget
        host.tab_widget.setTabBar(bar)
        host.split_tab_widget = host.split_content_stack = None
        host._schedule_session_snapshot = Mock()
        host.close_tab = Mock(side_effect=lambda index, target_tabwidget=None: (
            host.content_stack.removeWidget(host.content_stack.widget(index)), host.tab_widget.removeTab(index)))
        for tab in tabs:
            host.content_stack.addWidget(tab)
            host.tab_widget.addTab(QWidget(), tab.current_path)
        host._apply_tab_grouping_for_pane(host.tab_widget)
        return host

    def paths(self, host):
        return [host.content_stack.widget(index).current_path for index in range(host.content_stack.count())]

    def test_close_others_and_right_keep_pinned_and_clicked_tab(self):
        """关闭其他/右侧标签：保留被点击的标签和固定标签，其余按对象逐个关闭。"""
        tabs = [self.make_tab('P', pinned=True), self.make_tab('A'), self.make_tab('B'), self.make_tab('C')]
        host = self.host(tabs)
        with patch_all('show_toast'):
            self.assertEqual(host.close_tabs_to_right(tabs[1], host.tab_widget), 2)
            self.assertEqual(self.paths(host), ['P', 'A'])
            host2 = self.host([self.make_tab('P', pinned=True), self.make_tab('A'), self.make_tab('B')])
            self.assertEqual(host2.close_other_tabs(host2.content_stack.widget(2), host2.tab_widget), 1)
            self.assertEqual(self.paths(host2), ['P', 'B'])

    def test_collapsed_group_hides_members_shows_chip_and_expands_when_selected(self):
        """折叠分组只显示组首并带名称标签；选中被隐藏的成员时自动展开；组名随组首同步给成员。"""
        tabs = [self.make_tab('A', color='#64B5F6', name='Work'), self.make_tab('B', color='#64B5F6'),
                self.make_tab('C', color='#64B5F6'), self.make_tab('D')]
        host = self.host(tabs)
        bar = host.tab_widget.tabBar()
        host.tab_widget.setCurrentIndex(2)
        self.assertTrue(host.set_tab_group_collapsed(tabs[1], host.tab_widget, True))
        self.assertEqual(host.tab_widget.currentIndex(), 0)
        self.assertEqual([bar.isTabVisible(index) for index in range(4)], [True, False, False, True])
        chip = bar.tabButton(0, bar.LeftSide)
        self.assertIn('Work', chip.text())
        self.assertIn('3', chip.text())
        self.assertEqual([tab.tab_group_name for tab in tabs[:3]], ['Work'] * 3)
        host.tab_widget.setCurrentIndex(1)
        host._reveal_hidden_current_tab(host.tab_widget, 1)
        self.assertTrue(all(bar.isTabVisible(index) for index in range(4)))
        self.assertFalse(any(tab.tab_group_collapsed for tab in tabs))
        self.assertEqual(app.tab_group_runs(tabs), [(0, 2, '#64B5F6')])


class FrequentDirsTests(unittest.TestCase):
    def test_ranking_filter_remove_and_persistence(self):
        """常用目录按次数与最近访问排序，可按关键词筛选、移除；只记录真实路径并能保存后重新加载。"""
        from tabexplorer.frequent import FrequentDirs
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'frequent_dirs.json'
            store = FrequentDirs(path)
            now = time.time()
            for _index in range(3):
                store.record(r'D:\work\project', now=now - 30 * 86400)
            store.record(r'C:\recent', now=now)
            store.record(r'c:\RECENT\\', now=now)
            self.assertFalse(store.record('shell:Desktop', now=now))
            self.assertEqual([entry['path'] for entry in store.ranked(now=now)], [r'c:\RECENT', r'D:\work\project'])
            self.assertEqual([entry['path'] for entry in store.ranked('work proj', now=now)], [r'D:\work\project'])
            self.assertTrue(store.save())
            reloaded = FrequentDirs(path)
            self.assertEqual(len(reloaded.entries), 2)
            self.assertTrue(reloaded.remove(r'D:\WORK\project'))
            self.assertEqual([entry['path'] for entry in reloaded.ranked(now=now)], [r'c:\RECENT'])


class SyncBrowseTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        for side in ('left', 'right'):
            (self.root / side / 'sub' / 'nested').mkdir(parents=True)
        names = ('on_tab_path_changed', '_queue_sync_browse', '_apply_sync_browse',
                 '_set_sync_browse', '_normalize_path_for_compare', '_all_groups',
                 'get_current_tab_widget', 'get_tab_widget', '_get_split_pane')
        host_type = type('SyncHost', (QWidget,), {name: getattr(app.MainWindow, name) for name in names})
        self.host = host_type()
        self.addCleanup(self.dispose_host)
        self.host._record_frequent_visit = Mock()
        self.host._refresh_context_actions = Mock()
        self.host._split_active = True
        self.left = self.make_tab('left')
        self.right = self.make_tab('right')
        for prefix, tab in (('', self.left), ('split_', self.right)):
            tabs, stack = QTabWidget(self.host), QStackedWidget(self.host)
            tabs.addTab(QWidget(), 'tab')
            stack.addWidget(tab)
            setattr(self.host, prefix + 'tab_widget', tabs)
            setattr(self.host, prefix + 'content_stack', stack)
        self.host._set_sync_browse(True)

    def dispose_host(self):
        from PyQt5 import sip
        sip.delete(self.host)

    def make_tab(self, side):
        tab_type = type('PathTab', (QWidget,), {'current_path': app.FileExplorerTab.current_path})
        tab = tab_type(self.host)
        tab.main_window = self.host
        tab.current_path = str(self.root / side)
        tab.navigate_to = Mock(side_effect=lambda path: setattr(tab, 'current_path', path))
        tab._is_slow_path = lambda path: False
        return tab

    def settle(self):
        from PyQt5.QtTest import QTest
        QTest.qWait(app.SYNC_BROWSE_SETTLE_MS + 150)

    def test_same_directory_navigation_follows_both_ways_without_dead_time(self):
        """两侧起始目录相同：进入子目录后立即从跟随侧返回，再次进入时仍双向联动。"""
        self.right._current_path = self.left.current_path
        base = self.root / 'left'
        self.left.current_path = str(base / 'sub')
        self.settle()
        self.assertEqual(self.right.current_path, str(base / 'sub'))
        self.right.current_path = str(base)
        self.settle()
        self.assertEqual(self.left.current_path, str(base))
        self.left.current_path = str(base / 'sub' / 'nested')
        self.settle()
        self.assertEqual(self.right.current_path, str(base / 'sub' / 'nested'))
        self.assertEqual(self.left.navigate_to.call_count, 1)
        self.assertEqual(self.right.navigate_to.call_count, 2)

    def test_delayed_follow_completion_does_not_schedule_reverse_navigation(self):
        """异步跟随超过原来的 3 秒后完成，也不应触发反向导航。"""
        self.right.navigate_to.side_effect = None
        self.left.current_path = str(self.root / 'left' / 'sub')
        self.settle()
        target = self.right.navigate_to.call_args.args[0]
        with patch('tabexplorer.mainwindow.time.monotonic', return_value=time.monotonic() + 10):
            self.right.current_path = target
        self.assertIsNone(self.host._sync_pending)
        self.assertFalse(self.host._sync_browse_timer.isActive())

    def test_pending_navigation_does_not_move_a_switched_out_tab(self):
        """防抖等待中切换任何一侧的标签，旧配对不应继续联动隐藏标签。"""
        for prefix in ('', 'split_'):
            with self.subTest(side=prefix or 'left'):
                self.left._current_path = str(self.root / 'left')
                self.right._current_path = str(self.root / 'right')
                self.left.current_path = str(self.root / 'left' / 'sub')
                stack = getattr(self.host, prefix + 'content_stack')
                tabs = getattr(self.host, prefix + 'tab_widget')
                extra = self.make_tab('right')
                stack.addWidget(extra)
                tabs.addTab(QWidget(), 'other')
                tabs.setCurrentIndex(tabs.count() - 1)
                self.settle()
                self.right.navigate_to.assert_not_called()
                tabs.setCurrentIndex(0)

    def test_disabled_missing_unrelated_and_background_paths_do_not_follow(self):
        """关闭同步、缺少对应目录、无关跳转及后台标签变化均不移动另一侧。"""
        self.host._set_sync_browse(False)
        self.left.current_path = str(self.root / 'left' / 'sub')
        self.settle()
        self.right.navigate_to.assert_not_called()
        self.host._set_sync_browse(True)
        self.left.current_path = str(self.root / 'right')
        self.settle()
        self.right.navigate_to.assert_not_called()
        self.left._current_path = str(self.root / 'left')
        (self.root / 'left' / 'missing').mkdir()
        with patch_all('show_toast') as toast:
            self.left.current_path = str(self.root / 'left' / 'missing')
            self.settle()
            toast.assert_called_once()
        self.right.navigate_to.assert_not_called()
        background = self.make_tab('left')
        background.current_path = str(self.root / 'left' / 'sub')
        self.settle()
        self.right.navigate_to.assert_not_called()

    def test_rapid_navigation_coalesces_to_the_final_relative_path(self):
        """连续进入多级目录合并为一次跟随，不丢失最初的相对路径基准。"""
        self.left.current_path = str(self.root / 'left' / 'sub')
        self.left.current_path = str(self.root / 'left' / 'sub' / 'nested')
        self.settle()
        self.right.navigate_to.assert_called_once_with(str(self.root / 'right' / 'sub' / 'nested'))
        self.left.navigate_to.assert_not_called()

    def test_disabling_sync_clears_pending_and_follow_targets(self):
        """关掉同步取消待执行导航；再次启用不沿用未完成跟随的屏蔽标记。"""
        self.left.current_path = str(self.root / 'left' / 'sub')
        self.host._set_sync_browse(False)
        self.settle()
        self.right.navigate_to.assert_not_called()
        self.assertIsNone(self.host._sync_pending)
        self.left._current_path = str(self.root / 'left')
        self.host._set_sync_browse(True)
        self.right.navigate_to.side_effect = None
        self.left.current_path = str(self.root / 'left' / 'sub')
        self.settle()
        self.assertIsNotNone(self.right._sync_follow_target)
        self.host._set_sync_browse(False)
        self.host._set_sync_browse(True)
        self.assertIsNone(self.right._sync_follow_target)
        self.left._current_path = str(self.root / 'left')
        self.right.current_path = str(self.root / 'right' / 'sub')
        self.settle()
        self.assertEqual(self.left.current_path, str(self.root / 'left' / 'sub'))

    def test_failed_follow_clears_only_its_own_target(self):
        """导航失败清理对应跟随标记，过期的其他目标失败不影响新请求。"""
        target = str(self.root / 'right' / 'sub')
        self.right._sync_follow_target = self.host._normalize_path_for_compare(target)
        self.right._handle_navigation_failure = Mock()
        app.FileExplorerTab._on_navigation_failed(self.right, str(self.root / 'right'), 1)
        self.assertIsNotNone(self.right._sync_follow_target)
        app.FileExplorerTab._on_navigation_failed(self.right, target, 1)
        self.assertIsNone(self.right._sync_follow_target)

    def test_relative_moves_are_mirrored_and_unrelated_jumps_ignored(self):
        """进入子目录、返回上级按相对路径映射到另一侧；跨盘或无关跳转、越过根目录和虚拟位置不跟随。"""
        map_sync_path = app.map_sync_path
        self.assertEqual(map_sync_path(r'C:\v1', r'C:\V1\src\core', r'D:\v2'), r'D:\v2\src\core')
        self.assertEqual(map_sync_path(r'C:\v1\src\core', r'C:\v1', r'D:\v2\src\core'), r'D:\v2')
        self.assertIsNone(map_sync_path(r'C:\v1', r'E:\other', r'D:\v2'))
        self.assertIsNone(map_sync_path(r'C:\a\b', 'C:\\', 'D:\\'))
        self.assertIsNone(map_sync_path('shell:Desktop', r'C:\x', r'D:\v2'))

    def test_compare_copy_plan_skips_children_of_missing_folders(self):
        """“复制到另一侧”：整目录缺失时只复制目录本身；内容不同的文件先移入回收站再复制；类型不同的项跳过。"""
        with tempfile.TemporaryDirectory() as directory:
            left, right = Path(directory, 'L'), Path(directory, 'R')
            (left / 'new' / 'sub').mkdir(parents=True)
            (left / 'new' / 'sub' / 'a.txt').write_text('a', encoding='utf-8')
            (left / 'changed.txt').write_text('left', encoding='utf-8')
            right.mkdir()
            only_left, differs, type_differs = app.tr('仅左侧'), app.tr('内容不同'), app.tr('类型不同')
            rows = [('new', only_left, str(left / 'new'), ''),
                    (r'new\sub\a.txt', only_left, str(left / 'new' / 'sub' / 'a.txt'), ''),
                    ('changed.txt', differs, str(left / 'changed.txt'), str(right / 'changed.txt')),
                    ('x', type_differs, str(left / 'x'), str(right / 'x'))]
            copies, replaced, skipped = app.plan_compare_copy(rows, str(left), str(right), True)
            self.assertEqual(copies, {str(right): [str(left / 'new'), str(left / 'changed.txt')]})
            self.assertEqual(replaced, [str(right / 'changed.txt')])
            self.assertEqual(skipped, ['x'])
            copies, replaced, skipped = app.plan_compare_copy(rows, str(left), str(right), False)
            self.assertEqual((copies, replaced), ({str(left): [str(right / 'changed.txt')]}, [str(left / 'changed.txt')]))


class SearchFilterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def search_rows(self, directory, keyword, **options):
        task = app._SearchTask(directory, '', 100, None)
        app.SearchDialog.do_search(task, keyword, True, options.pop('content', False), **options)
        rows = []
        while not task.result_queue.empty():
            message = task.result_queue.get_nowait()
            if message['type'] == 'result_batch':
                rows.extend(message['items'])
        return {Path(row['path']).name for row in rows}

    def test_size_and_pattern_helpers(self):
        """大小文本支持 B/KB/MB/GB 和小数；正则与全词可组合；大小筛选排除文件夹，日期筛选按修改时间。"""
        self.assertEqual([app.parse_size_text(text) for text in ('', '500', '10kb', '1.5 M', '2G')],
                         [None, 500, 10240, 1572864, 2147483648])
        with self.assertRaises(ValueError):
            app.parse_size_text('ten MB')
        pattern = app.compile_search_pattern(r'log_\d+', match_whole_word=True, use_regex=True)
        self.assertTrue(pattern.search('x LOG_42 y'))
        self.assertFalse(pattern.search('xlog_42'))
        self.assertIsNone(app.compile_search_pattern('plain'))
        filters = {'min_size': 10, 'max_size': None, 'modified_after': 100}
        self.assertTrue(app.passes_search_filters(filters, 10, 100, False))
        self.assertFalse(app.passes_search_filters(filters, 9, 100, False))
        self.assertFalse(app.passes_search_filters(filters, 50, 99, False))
        self.assertFalse(app.passes_search_filters(filters, 50, 200, True))
        self.assertTrue(app.passes_search_filters({'modified_after': 100}, 0, 200, True))

    def test_regex_and_size_filters_in_python_scan(self):
        """Python 扫描路径下，正则同时作用于文件名和内容，大小/日期筛选在内容扫描前排除文件。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'Log_01.txt').write_bytes(b'x' * 2048)
            (root / 'Log_02.txt').write_bytes(b'x' * 10)
            (root / 'Logbook.txt').write_bytes(b'x' * 4096)
            (root / 'code.c').write_text('int value_42 = 1;', encoding='utf-8')
            old = root / 'old.c'
            old.write_text('int value_7 = 1;', encoding='utf-8')
            os.utime(old, (time.time() - 40 * 86400, time.time() - 40 * 86400))
            self.assertEqual(self.search_rows(directory, r'^Log_\d+', use_regex=True), {'Log_01.txt', 'Log_02.txt'})
            self.assertEqual(self.search_rows(directory, r'^Log_\d+', use_regex=True,
                                              filters={'min_size': 1024}), {'Log_01.txt'})
            self.assertEqual(self.search_rows(directory, r'value_\d+', use_regex=True, content=True),
                             {'code.c', 'old.c'})
            recent = {'modified_after': time.time() - 7 * 86400}
            self.assertEqual(self.search_rows(directory, r'value_\d+', use_regex=True, content=True, filters=recent),
                             {'code.c'})


class TranslationCoverageTests(unittest.TestCase):
    def test_every_translated_ui_string_has_english_text(self):
        """所有 tr("中文") 文本及动态标签都有英文翻译，占位符 {} 数量一致，避免英文界面显示中文。"""
        from tabexplorer.i18n import _LANG_EN
        from tabexplorer.hotkeys import HOTKEY_COMMANDS, _TOOLBAR_HOTKEY_TOOLTIPS
        from tabexplorer.search import SEARCH_MODIFIED_RANGES
        cjk = re.compile(r'[\u4e00-\u9fff]')
        texts = {}
        for path in MODULE_FILES:
            if path.name == 'i18n.py':
                continue
            for node in ast.walk(ast.parse(path.read_text(encoding='utf-8'))):
                if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == 'tr'
                        and node.args and isinstance(node.args[0], ast.Constant)
                        and isinstance(node.args[0].value, str) and cjk.search(node.args[0].value)):
                    texts.setdefault(node.args[0].value, f'{path.name}:{node.lineno}')
        for label in ([command[3] for command in HOTKEY_COMMANDS] + [label for label, _cmd in _TOOLBAR_HOTKEY_TOOLTIPS.values()]
                      + [label for label, _days in SEARCH_MODIFIED_RANGES] + ['此电脑', '桌面', '回收站', '下载']):
            texts.setdefault(label, 'dynamic label')
        missing = sorted(f'{where}: {text!r}' for text, where in texts.items() if text not in _LANG_EN)
        self.assertEqual(missing, [])
        mismatched = sorted(text for text in texts if text.count('{}') != _LANG_EN[text].count('{}'))
        self.assertEqual(mismatched, [])


if __name__ == '__main__':
    unittest.main()

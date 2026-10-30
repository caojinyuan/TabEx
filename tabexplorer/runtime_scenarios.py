"""Native lifecycle exercises; called only by isolated --self-test runs."""

import time

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication, QMessageBox


def lifecycle_steps(check):
    w, root = check.window, check.root

    def ready(tab, path):
        explorer = getattr(tab, 'explorer', None)
        if not getattr(explorer, '_init_ok', False):
            return False
        location = explorer._url_to_path(explorer.property('LocationURL'))
        return check.same_path(tab.current_path, path) and check.same_path(location, path)

    def wait(seconds):
        until = time.monotonic() + seconds
        return lambda: time.monotonic() >= until

    def confirm_retry(panel):
        previous_count = len(panel.records)
        def accept_confirmation():
            box = QApplication.activeModalWidget()
            if isinstance(box, QMessageBox) and box.parent() is panel:
                box.button(QMessageBox.Yes).click()
        QTimer.singleShot(50, accept_confirmation)
        panel.retry_selected()
        if len(panel.records) != previous_count + 1:
            raise AssertionError('Retry confirmation did not start a new task')

    while True:
        left = w.get_current_tab_widget()
        yield 'base view ready', lambda: ready(left, root)
        w.add_new_tab(str(root))
        right = w.get_current_tab_widget()
        yield 'second view ready', lambda: ready(right, root)
        w._enter_split_view()
        w.sync_browse_action.trigger()
        if not w._sync_browse or w.split_controls.isHidden():
            raise AssertionError('Split controls or synchronization not enabled')
        left.explorer.dynamicCall('Navigate(const QString&)', str(root / 'navigation'))
        yield 'left to right native navigation', lambda: (
            ready(left, root / 'navigation') and ready(right, root / 'navigation'))
        right.explorer.dynamicCall('Navigate(const QString&)', str(root))
        yield 'immediate right to left native navigation', lambda: ready(left, root) and ready(right, root)
        right.explorer.dynamicCall('Navigate(const QString&)', str(root / 'navigation' / 'nested'))
        yield 'nested right to left navigation', lambda: (
            ready(left, root / 'navigation' / 'nested') and ready(right, root / 'navigation' / 'nested'))
        left.navigate_to(str(root))
        yield 'return both panes', lambda: ready(left, root) and ready(right, root)
        w._merge_split_back()
        if w._sync_browse or not w.split_controls.isHidden():
            raise AssertionError('Split state survived merge')
        w.tab_widget.setCurrentIndex(0)
        yield 'background tab inactive', lambda: not right.isVisible() and not right._refresh_active
        if not right.hibernate_shell_view():
            raise AssertionError('Background Shell view did not hibernate')
        w.tab_widget.setCurrentIndex(w.content_stack.indexOf(right))
        yield 'hibernated view wakes', lambda: ready(right, root)
        w.close_tab(w.content_stack.indexOf(right))
        yield 'tab closed', lambda: w.tab_widget.count() == 1 and ready(left, root)

        for cancel in (True, False):
            from .search import _search_cache
            _search_cache.clear()
            w.show_search_dialog()
            dialog = w.search_dialogs[-1]
            dialog.path_input.setText(str(root / 'fixture'))
            dialog.use_everything_cb.setChecked(False)
            dialog.search_filename_cb.setChecked(True)
            dialog.search_content_cb.setChecked(True)
            dialog.file_type_input.clear()
            dialog.search_input.setEditText('sample')
            dialog.start_search()
            thread = dialog.search_thread
            if thread is None:
                raise AssertionError('Search did not start')
            if cancel:
                dialog.stop_search()
            else:
                yield 'search finishes with results', lambda: (
                    not thread.is_alive() and dialog.current_result_count > 0)
            dialog.close()
            yield 'search thread exits', lambda: not thread.is_alive()
            if w.search_dialogs:
                raise AssertionError('Closed search dialog is still tracked')

        panel = w.get_file_task_panel()
        source, missing = root / 'fixture' / 'sample.bin', root / 'fixture' / 'missing.bin'
        destination = root / 'destination'
        panel.max_running_tasks = 0
        panel.start_task('copy', [str(source)], str(destination))
        panel.cancel_selected()
        yield 'queued copy cancelled', lambda: not panel.has_running_tasks()
        if not panel.records[-1]['cancelled']:
            raise AssertionError('Cancellation was not recorded')
        panel.max_running_tasks = 2
        confirm_retry(panel)
        yield 'cancelled copy resumed', lambda: not panel.has_running_tasks()
        if (destination / source.name).read_bytes() != source.read_bytes():
            raise AssertionError('Resumed copy content mismatch')
        panel.start_task('copy', [str(missing)], str(destination))
        yield 'missing source failure reported', lambda: not panel.has_running_tasks()
        if not panel.records[-1]['failed']:
            raise AssertionError('Missing source was not reported')
        missing.write_bytes(b'retry payload')
        confirm_retry(panel)
        yield 'failed copy retried', lambda: not panel.has_running_tasks()
        if (destination / missing.name).read_bytes() != missing.read_bytes():
            raise AssertionError('Retry copy content mismatch')
        panel.start_task('permanent_delete', [str(destination / source.name),
                                             str(destination / missing.name), str(missing)])
        yield 'temporary output removed', lambda: not panel.has_running_tasks()
        if any(path.exists() for path in (destination / source.name, destination / missing.name, missing)):
            raise AssertionError('Temporary output cleanup failed')
        panel.clear_completed()
        if panel.records or panel.timer.isActive():
            raise AssertionError('Completed task state did not drain')
        panel.hide()
        w.showMinimized()
        QTimer.singleShot(100, w.showMaximized)
        yield 'minimize and restore', lambda: not w.isMinimized() and ready(left, root)
        check.restored = True
        # Allow deferred QObject/COM destruction and filesystem notifications to settle.
        yield 'settling at baseline state', wait(check.settle_seconds)
        yield 'resource checkpoint', None

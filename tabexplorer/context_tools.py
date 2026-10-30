"""Context-aware external tools. Arguments are expanded individually, never as shell commands."""

from copy import deepcopy
from dataclasses import dataclass
import os
import json
import shutil
from string import Formatter
import subprocess
import uuid

from PyQt5 import sip
from PyQt5.QtCore import QThread, Qt, pyqtSignal
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtCore import QUrl
from PyQt5.QtWidgets import (
    QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMessageBox, QPlainTextEdit, QPushButton, QStyle, QToolButton, QVBoxLayout,
)

from .debuglog import debug_print
from .i18n import tr
from .system import _retain_thread_until_finished
from .widgets import show_toast

_tool_readiness = {}


@dataclass(frozen=True)
class ToolContext:
    cwd: str = ''
    selected: tuple[str, ...] = ()
    left: str = ''
    right: str = ''
    line: int = 0
    column: int = 1


def default_tools():
    return [
        {'id': 'vscode-folder', 'name': tr("VS Code：打开当前目录"), 'program': 'Code.exe',
         'args': ['--reuse-window', '{cwd}'], 'cwd': '{cwd}'},
        {'id': 'vscode-files', 'name': tr("VS Code：打开选中项"), 'program': 'Code.exe',
         'args': ['--reuse-window', '{selected}'], 'cwd': '{cwd}'},
        {'id': 'vscode-line', 'name': tr("VS Code：打开匹配行"), 'program': 'Code.exe',
         'args': ['--reuse-window', '--goto', '{file}:{line}:{column}'], 'cwd': '{cwd}'},
        {'id': 'winmerge-files', 'name': tr("WinMerge：比较两个选中项"), 'program': 'WinMergeU.exe',
         'args': ['/e', '/u', '{file1}', '{file2}'], 'cwd': '{cwd}'},
        {'id': 'winmerge-folders', 'name': tr("WinMerge：比较左右目录"), 'program': 'WinMergeU.exe',
         'args': ['/r', '/e', '/u', '{left}', '{right}'], 'cwd': '{cwd}'},
        {'id': 'vscode-diff', 'name': tr("VS Code：比较两个文件"), 'program': 'Code.exe',
         'args': ['--reuse-window', '--diff', '{file1}', '{file2}'], 'cwd': '{cwd}'},
    ]


def configured_tools(config):
    if isinstance(config, dict) and 'external_tools' not in config:
        return default_tools()
    tools = config.get('external_tools') if isinstance(config, dict) else None
    if not isinstance(tools, list) or not all(isinstance(tool, dict) for tool in tools):
        raise ValueError(tr("外部工具配置格式无效"))
    for tool in tools:
        validate_tool(tool)
    return deepcopy(tools)


def validate_tool(tool):
    for key in ('name', 'program', 'cwd'):
        if not isinstance(tool.get(key), str) or not tool[key].strip():
            raise ValueError(tr("名称、程序和工作目录不能为空"))
    if not tool['program'].lower().endswith('.exe'):
        raise ValueError(tr("请选择 EXE 程序；脚本请通过解释器 EXE 和独立参数启动"))
    if not isinstance(tool.get('args'), list) or not all(isinstance(arg, str) for arg in tool['args']):
        raise ValueError(tr("参数必须是字符串列表"))
    if any('\x00' in value for value in [tool['program'], tool['cwd'], *tool['args']]):
        raise ValueError(tr("工具配置不能包含空字符"))
    for template in [tool['cwd'], *tool['args']]:
        for _text, field, spec, conversion in Formatter().parse(template):
            if field is None:
                continue
            if field not in ('cwd', 'selected', 'file', 'file1', 'file2', 'left', 'right', 'line', 'column'):
                raise ValueError(tr("未知参数占位符: {}").format(field))
            if spec or conversion or (field == 'selected' and template != '{selected}'):
                raise ValueError(tr("占位符不支持格式修饰；{selected} 必须单独占一行"))
            if field == 'selected' and template == tool['cwd']:
                raise ValueError(tr("工作目录不能使用多选占位符"))


def prepare_tool(tool, context):
    validate_tool(tool)

    def value(field):
        if field in ('selected', 'file', 'file1', 'file2') and any(
                not os.path.isabs(path) or path.startswith(('shell:', '::')) for path in context.selected):
            raise ValueError(tr("选中项必须是文件系统路径"))
        if field == 'selected':
            if not context.selected:
                raise ValueError(tr("请先选择文件或文件夹"))
            return context.selected
        if field == 'file':
            if len(context.selected) != 1:
                raise ValueError(tr("需要恰好一个选中项"))
            return context.selected[0]
        if field in ('file1', 'file2'):
            if len(context.selected) != 2:
                raise ValueError(tr("需要恰好两个选中项"))
            return context.selected[0 if field == 'file1' else 1]
        if field in ('line', 'column'):
            if context.line < 1:
                raise ValueError(tr("仅内容搜索匹配行可用"))
            return str(getattr(context, field))
        path = getattr(context, field)
        if not path or path.startswith(('shell:', '::')) or not os.path.isabs(path):
            raise ValueError(tr("缺少文件夹上下文: {}").format(field))
        return path

    def expand(template):
        output = []
        for text, field, _spec, _conversion in Formatter().parse(template):
            output.append(text)
            if field is not None:
                output.append(value(field))
        return ''.join(output)

    args = []
    for template in tool['args']:
        if template == '{selected}':
            args.extend(value('selected'))
        else:
            args.append(expand(template))
    cwd = expand(tool['cwd'])
    if not os.path.isabs(cwd) or cwd.startswith(('shell:', '::')):
        raise ValueError(tr("工作目录必须是绝对路径"))
    return [tool['program'], *args], cwd


def resolve_program(program):
    program = os.path.expandvars(os.path.expanduser(program))
    if os.path.dirname(program):
        if os.path.isabs(program) and os.path.isfile(program):
            return program
        raise FileNotFoundError(tr("找不到工具程序: {}").format(program))
    found = shutil.which(program)
    if found:
        return found
    locations = []
    if program.casefold() == 'code.exe':
        locations = [
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Microsoft VS Code', 'Code.exe'),
            os.path.join(os.environ.get('ProgramFiles', ''), 'Microsoft VS Code', 'Code.exe'),
            os.path.join(os.environ.get('ProgramFiles(x86)', ''), 'Microsoft VS Code', 'Code.exe'),
        ]
    elif program.casefold() == 'winmergeu.exe':
        locations = [os.path.join(os.environ.get(key, ''), 'WinMerge', 'WinMergeU.exe')
                     for key in ('ProgramFiles', 'ProgramFiles(x86)', 'LOCALAPPDATA')]
        locations.append(os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'WinMerge', 'WinMergeU.exe'))
    for path in locations:
        if os.path.isabs(path) and os.path.isfile(path):
            return path
    raise FileNotFoundError(tr("找不到工具程序: {}").format(program))

def tool_download_url(program):
    return {'winmergeu.exe': 'https://winmerge.org/downloads/',
            'code.exe': 'https://code.visualstudio.com/Download'}.get(os.path.basename(program).casefold(), '')


class ToolProbeWorker(QThread):
    completed = pyqtSignal(object)

    def __init__(self, programs):
        super().__init__()
        self.programs = list(dict.fromkeys(programs))

    def run(self):
        results = {}
        for program in self.programs:
            try:
                results[program] = (resolve_program(program), '')
            except (OSError, ValueError) as error:
                results[program] = ('', str(error))
        self.completed.emit(results)


class ToolLaunchWorker(QThread):
    completed = pyqtSignal(str)

    def __init__(self, tool, context):
        super().__init__()
        self.tool = deepcopy(tool)
        self.context = context

    def run(self):
        error = ''
        try:
            args, cwd = prepare_tool(self.tool, self.context)
            args[0] = resolve_program(args[0])
            if not os.path.isdir(cwd):
                raise NotADirectoryError(cwd)
            selection_used = any(
                field in ('selected', 'file', 'file1', 'file2')
                for template in [self.tool['cwd'], *self.tool['args']]
                for _text, field, _spec, _conversion in Formatter().parse(template))
            if selection_used:
                for path in self.context.selected:
                    if not os.path.exists(path):
                        raise FileNotFoundError(path)
            flags = (getattr(subprocess, 'DETACHED_PROCESS', 0)
                     | getattr(subprocess, 'CREATE_NEW_PROCESS_GROUP', 0))
            subprocess.Popen(args, cwd=cwd, shell=False, close_fds=True, creationflags=flags)
        except (OSError, ValueError) as failure:
            error = str(failure)
            debug_print(f"[ExternalTool] {self.tool.get('name', '')}: {error}")
        self.completed.emit(error)


def launch_context_tool(owner, tool, context):
    worker = ToolLaunchWorker(tool, context)

    def finished(error):
        if not sip.isdeleted(owner):
            show_toast(owner, tr("外部工具"), error or tr("已启动"),
                       level='error' if error else 'success')

    worker.completed.connect(finished)
    worker.start()
    _retain_thread_until_finished(worker)
    return worker


def populate_tools_menu(menu, tools, context, owner):
    menu.clear()
    for tool in tools:
        action = menu.addAction(str(tool.get('name', tr("外部工具"))))
        try:
            args, cwd = prepare_tool(tool, context)
            action.setToolTip(subprocess.list2cmdline(args) + '\n' + cwd)
            resolved, error = _tool_readiness.get(tool['program'], ('', ''))
            if error:
                action.setEnabled(False)
                action.setProperty('unavailable_reason', error)
                action.setToolTip(error + '\n' + tr("请在配置外部工具中检测或重新选择程序"))
            elif resolved:
                action.setToolTip(resolved + '\n' + action.toolTip())
            action.triggered.connect(
                lambda checked=False, selected=tool: launch_context_tool(owner, selected, context))
        except ValueError as error:
            action.setEnabled(False)
            action.setProperty('unavailable_reason', str(error))
            action.setToolTip(str(error))


class ContextToolsDialog(QDialog):
    def __init__(self, tools, parent=None, context=None):
        super().__init__(parent)
        self.setWindowTitle(tr("配置外部工具"))
        self.resize(800, 570)
        self.tools = deepcopy(tools)
        self.context = context or ToolContext()
        self._probe_worker = None
        self._probe_again = False
        self._current = None
        layout = QVBoxLayout(self)
        body = QHBoxLayout()
        self.list = QListWidget(self)
        body.addWidget(self.list, 1)
        form = QFormLayout()
        self.name = QLineEdit(self)
        self.program = QLineEdit(self)
        self.program.setMinimumWidth(140)
        self.arguments = QPlainTextEdit(self)
        self.arguments.setPlaceholderText(tr("每行一个参数，无需添加引号"))
        self.cwd = QLineEdit(self)
        form.addRow(tr("名称"), self.name)
        program_row = QHBoxLayout()
        program_row.addWidget(self.program)
        browse = QToolButton(self)
        browse.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        browse.setToolTip(tr("选择工具程序"))
        browse.setAccessibleName(tr("选择工具程序"))
        browse.clicked.connect(self.browse_program)
        program_row.addWidget(browse)
        form.addRow(tr("程序"), program_row)
        form.addRow(tr("参数（每行一个）"), self.arguments)
        form.addRow(tr("工作目录"), self.cwd)
        body.addLayout(form, 2)
        layout.addLayout(body)
        self.tool_status = QLabel(self)
        self.tool_status.setWordWrap(True)
        self.tool_status.setTextFormat(Qt.PlainText)
        self.tool_status.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(self.tool_status)
        readiness = QHBoxLayout()
        self.detect_button = QPushButton(tr("检测工具"), self)
        self.detect_button.clicked.connect(self.detect_tools)
        self.download_button = QPushButton(tr("打开官方下载页"), self)
        self.download_button.clicked.connect(self.open_download_page)
        self.preview_button = QPushButton(tr("预览启动参数（不执行）"), self)
        self.preview_button.setToolTip(tr("仅预览参数，不会启动程序"))
        self.preview_button.clicked.connect(self.preview_arguments)
        for button in (self.detect_button, self.download_button, self.preview_button):
            readiness.addWidget(button)
        layout.addLayout(readiness)
        self.argument_preview = QPlainTextEdit(self)
        self.argument_preview.setReadOnly(True)
        self.argument_preview.setMaximumHeight(110)
        self.argument_preview.hide()
        layout.addWidget(self.argument_preview)
        hint = QLabel(tr("占位符：{cwd} 当前目录；{selected} 多选；{file} 单选；"
                        "{file1}/{file2} 双选；{left}/{right} 左右目录；{line}/{column} 匹配位置。\n"
                        "只启动本机工具，不自动安装或上传文件。自定义工具会按所填参数执行。"), self)
        hint.setWordWrap(True)
        layout.addWidget(hint)
        controls = QHBoxLayout()
        for label, callback in (("添加工具", self.add_tool), ("删除工具", self.remove_tool),
                                ("添加缺少的预设", self.add_presets)):
            button = QPushButton(tr(label), self)
            button.clicked.connect(callback)
            controls.addWidget(button)
        layout.addLayout(controls)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.list.currentItemChanged.connect(self.select_tool)
        for tool in self.tools:
            self._add_item(tool)
        self.list.setCurrentRow(0)
        self.program.editingFinished.connect(self.detect_tools)
        self.detect_tools()

    def detect_tools(self):
        self.save_current()
        if self._probe_worker is not None:
            self._probe_again = True
            return
        programs = [self.list.item(index).data(Qt.UserRole).get('program', '')
                    for index in range(self.list.count())]
        self.detect_button.setEnabled(False)
        self.tool_status.setText(tr("正在检测工具..."))
        worker = ToolProbeWorker([program for program in programs if program])
        self._probe_worker = worker
        worker.completed.connect(self._tools_detected)
        worker.start()
        _retain_thread_until_finished(worker)

    def _tools_detected(self, results):
        _tool_readiness.update(results)
        self._probe_worker = None
        self.detect_button.setEnabled(True)
        for index in range(self.list.count()):
            item = self.list.item(index)
            tool = item.data(Qt.UserRole)
            path, error = results.get(tool.get('program', ''), ('', ''))
            state = tr("已就绪") if path else (tr("路径失效") if os.path.dirname(tool.get('program', ''))
                                             else tr("未安装或未找到"))
            item.setText(tool['name'] + ' · ' + state)
            item.setToolTip(path or error)
        self._show_tool_status()
        if self._probe_again:
            self._probe_again = False
            self.detect_tools()

    def _show_tool_status(self):
        program = self.program.text().strip()
        self.download_button.setEnabled(bool(tool_download_url(program)))
        self.download_button.setToolTip(tool_download_url(program))
        path, error = _tool_readiness.get(program, ('', ''))
        self.tool_status.setText((tr("已就绪") + ': ' + path) if path
                                 else error or tr("尚未检测，请选择程序后点击检测"))

    def open_download_page(self):
        url = tool_download_url(self.program.text().strip())
        if not url or not QDesktopServices.openUrl(QUrl(url)):
            show_toast(self, tr("外部工具"), tr("无法打开官方下载页"), level='error')

    def preview_arguments(self):
        self.save_current()
        self.argument_preview.show()
        if self._current is None:
            self.argument_preview.setPlainText(tr("请先选择工具"))
            return
        try:
            args, cwd = prepare_tool(self._current.data(Qt.UserRole), self.context)
            self.argument_preview.setPlainText(
                tr("工作目录") + ': ' + cwd + '\n' + json.dumps(args, ensure_ascii=False, indent=2))
        except ValueError as error:
            self.argument_preview.setPlainText(tr("当前不可用: {}").format(error))

    def _add_item(self, tool):
        item = QListWidgetItem(str(tool.get('name', '')), self.list)
        item.setData(Qt.UserRole, tool)
        return item

    def save_current(self):
        if self._current is not None:
            tool = self._current.data(Qt.UserRole)
            tool.update(name=self.name.text().strip(), program=self.program.text().strip(),
                        args=self.arguments.toPlainText().splitlines(), cwd=self.cwd.text().strip())
            self._current.setData(Qt.UserRole, tool)
            self._current.setText(tool['name'])

    def select_tool(self, current, previous):
        self.save_current()
        self._current = current
        tool = current.data(Qt.UserRole) if current else {}
        self.name.setText(tool.get('name', ''))
        self.program.setText(tool.get('program', ''))
        self.arguments.setPlainText('\n'.join(tool.get('args', [])))
        self.cwd.setText(tool.get('cwd', '{cwd}'))
        self.argument_preview.clear()
        self._show_tool_status()

    def add_tool(self):
        item = self._add_item({'id': uuid.uuid4().hex, 'name': tr("新工具"),
                               'program': '', 'args': ['{selected}'], 'cwd': '{cwd}'})
        self.list.setCurrentItem(item)

    def remove_tool(self):
        self._current = None
        self.list.takeItem(self.list.currentRow())
        self.select_tool(self.list.currentItem(), None)

    def add_presets(self):
        self.save_current()
        ids = {self.list.item(index).data(Qt.UserRole).get('id') for index in range(self.list.count())}
        for tool in default_tools():
            if tool['id'] not in ids:
                self._add_item(tool)
        if self.list.currentRow() < 0:
            self.list.setCurrentRow(0)

    def browse_program(self):
        path, _filter = QFileDialog.getOpenFileName(self, tr("选择工具程序"), '', 'EXE (*.exe)')
        if path:
            self.program.setText(path)

    def accept(self):
        self.save_current()
        tools = [self.list.item(index).data(Qt.UserRole) for index in range(self.list.count())]
        try:
            for tool in tools:
                validate_tool(tool)
        except ValueError as error:
            QMessageBox.warning(self, tr("外部工具"), str(error))
            return
        self.tools = tools
        super().accept()

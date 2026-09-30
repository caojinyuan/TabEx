"""AI 助手侧边栏。"""

import json
import os

from PyQt5.QtCore import pyqtSignal, Qt, QThread
from PyQt5.QtWidgets import QHBoxLayout, QLabel, QMenu, QPushButton, QVBoxLayout, QWidget

from .paths import get_app_base_dir, get_app_data_path, translate_common_path
from .i18n import tr
from .constants import MAX_CHAT_HISTORY_MESSAGES, MAX_CHAT_MESSAGE_CHARS, MAX_CONTEXT_TOTAL_CHARS
from .debuglog import debug_print
from .system import launch_detached, _retain_thread_until_finished
from .widgets import show_toast
from .fileops import _atomic_reviewed_write, _confirm_file_preview
from .ai_actions import _AiActionMethods, AiActionWorker


# ─────────────────────────────────────────────────────────────────────────────
# AI 聊天面板：ChatWorker（异步 API 线程）+ ChatPanel（UI 面板）
# ─────────────────────────────────────────────────────────────────────────────


class ChatWorker(QThread):
    """在后台线程中发起 OpenAI 兼容 API 请求，避免阻塞 UI。支持流式输出。"""
    response_received = pyqtSignal(str)
    token_received = pyqtSignal(str)   # 流式逐块 token
    error_occurred = pyqtSignal(str)

    def __init__(self, messages, api_url, api_key, model, stream=True, parent=None):
        super().__init__(parent)
        self.messages = messages
        self.api_url = api_url.rstrip('/')
        self.api_key = api_key
        self.model = model
        self.stream = stream

    def run(self):
        try:
            import requests, json as _json
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}" if self.api_key else "",
            }
            payload = {
                "model": self.model,
                "messages": self.messages,
                "stream": self.stream,
            }
            url = self.api_url + '/chat/completions'
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            if self.stream:
                try:
                    resp = requests.post(url, headers=headers, json=payload,
                                         timeout=120, verify=False, stream=True)
                    resp.raise_for_status()
                except requests.exceptions.ConnectionError as _ce:
                    if 'prematurely' in str(_ce).lower() or 'protocol' in str(_ce).lower():
                        self.error_occurred.emit(
                            tr('❌ 流式响应被服务器提前关闭（上下文可能过长）。\n') + str(_ce)
                        )
                        return
                    raise
                full_content = []
                try:
                    for raw_line in resp.iter_lines():
                        if self.isInterruptionRequested():
                            break
                        if not raw_line:
                            continue
                        line = raw_line.decode('utf-8', errors='replace') if isinstance(raw_line, bytes) else raw_line
                        if not line.startswith('data: '):
                            continue
                        data_str = line[6:].strip()
                        if data_str == '[DONE]':
                            break
                        try:
                            data = _json.loads(data_str)
                            chunk = ((data.get('choices') or [{}])[0]
                                     .get('delta', {}).get('content')) or ''
                            if chunk:
                                full_content.append(chunk)
                                self.token_received.emit(chunk)
                        except Exception:
                            pass
                except Exception as _se:
                    _msg = str(_se)
                    if 'prematurely' in _msg.lower() or 'protocol' in _msg.lower():
                        partial = ''.join(full_content)
                        if partial:
                            # 已收到部分内容，仍然返回
                            self.response_received.emit(partial)
                        else:
                            self.error_occurred.emit(
                                tr('❌ 流式响应被服务器提前关闭（上下文可能过长，请清空聊天记录重试）。\n') + _msg
                            )
                        return
                    raise
                finally:
                    resp.close()
                if self.isInterruptionRequested():
                    return
                self.response_received.emit(''.join(full_content))
            else:
                with requests.post(url, headers=headers, json=payload, timeout=120, verify=False) as resp:
                    resp.raise_for_status()
                    data = resp.json()
                    content = data['choices'][0]['message']['content']
                if not self.isInterruptionRequested():
                    self.response_received.emit(content)
        except Exception as e:
            self.error_occurred.emit(str(e))


class ChatPanel(QWidget):
    """右侧 AI 聊天侧边栏面板。"""

    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.messages = []   # 对话历史（不含 system prompt）
        self.worker = None
        self._action_worker = None
        self._closing = False
        self._request_cancelled = False
        self.history_file = get_app_data_path("chat_history.json")
        self.workflow_file = get_app_data_path("ai_workflows.json")
        self.workflow_templates = self._load_workflow_templates()
        self._is_loading_history = False
        # 给 ChatPanel 独立的 Win32 HWND，避免 QAxWidget(Shell Explorer) 抢占鼠标/键盘头
        self.setAttribute(Qt.WA_NativeWindow, True)
        self.setFocusPolicy(Qt.ClickFocus)
        # 启用拖拽
        self.setAcceptDrops(True)
        self._setup_ui()
        self._load_history()  # 启动时加载聊天记录

    # ── UI 构建 ──────────────────────────────────────────────────────────────
    def _setup_ui(self):
        from PyQt5.QtWidgets import QTextBrowser, QPlainTextEdit
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        # 标题栏
        header = QWidget()
        header.setStyleSheet("background: #E3F2FD; border-radius: 4px;")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(8, 4, 8, 4)
        title_lbl = QLabel(tr("🤖 AI 助手"))
        title_lbl.setStyleSheet("font-weight: bold; font-size: 10.5pt; background: transparent;")
        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        clear_btn = QPushButton(tr("清空"))
        clear_btn.setFixedHeight(22)
        clear_btn.setStyleSheet(
            "QPushButton{background:#fff;border:1px solid #90CAF9;border-radius:3px;font-size:9pt;padding:0 6px;}"
            "QPushButton:hover{background:#BBDEFB;}"
        )
        clear_btn.clicked.connect(self.clear_chat)
        header_layout.addWidget(clear_btn)
        layout.addWidget(header)

        # 当前目录提示条
        self.context_label = QLabel(tr("当前目录: —"))
        self.context_label.setStyleSheet(
            "color:#555; font-size:8.5pt; padding:2px 6px;"
            "background:#f0f4f8; border-radius:3px;"
        )
        self.context_label.setWordWrap(True)
        layout.addWidget(self.context_label)

        # 聊天历史显示区
        self.chat_display = QTextBrowser()
        self.chat_display.setOpenExternalLinks(False)
        self.chat_display.setTextInteractionFlags(
            Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard | Qt.TextBrowserInteraction
        )
        self.chat_display.setStyleSheet(
            "QTextBrowser{background:#fafafa;border:1px solid #e0e0e0;"
            "border-radius:4px;font-family:'Microsoft YaHei UI','Segoe UI',Arial;"
            "font-size:9.5pt;}"
        )
        self.chat_display.setContextMenuPolicy(Qt.CustomContextMenu)
        self.chat_display.customContextMenuRequested.connect(self._show_context_menu)
        layout.addWidget(self.chat_display, 1)

        # 输入框
        self.input_box = QPlainTextEdit()
        self.input_box.setPlaceholderText(tr("输入问题… (Enter 发送，Shift+Enter 换行)"))
        self.input_box.setFixedHeight(72)
        self.input_box.setReadOnly(False)
        self.input_box.setFocusPolicy(Qt.StrongFocus)
        self.input_box.setAttribute(Qt.WA_InputMethodEnabled, True)
        self.input_box.setStyleSheet(
            "QPlainTextEdit{border:1px solid #d0d0d0;border-radius:4px;padding:4px;"
            "font-family:'Microsoft YaHei UI','Segoe UI',Arial;font-size:9.5pt;}"
            "QPlainTextEdit:focus{border:1px solid #42A5F5;}"
        )
        self.input_box.installEventFilter(self)
        # QPlainTextEdit 的鼠标事件实际落在 viewport 上，这里一并监听
        self.input_box.viewport().installEventFilter(self)
        layout.addWidget(self.input_box)

        # 发送按钮行
        btn_row = QHBoxLayout()
        self.send_btn = QPushButton(tr("发 送"))
        self.send_btn.setFixedHeight(28)
        self.send_btn.setStyleSheet(
            "QPushButton{background:#1976D2;color:white;border:none;border-radius:4px;"
            "font-size:9.5pt;font-weight:bold;padding:0 16px;}"
            "QPushButton:hover{background:#1565C0;}"
            "QPushButton:disabled{background:#BDBDBD;}"
        )
        self.send_btn.clicked.connect(self.send_message)
        self.cancel_btn = QPushButton(tr("取消"))
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_actions)
        btn_row.addStretch()
        btn_row.addWidget(self.cancel_btn)
        btn_row.addWidget(self.send_btn)
        layout.addLayout(btn_row)

        # 快捷工作流按钮（过渡版 AI workflow）
        wf_row = QHBoxLayout()
        wf_row.setSpacing(6)
        wf_row.addWidget(QLabel(tr("快捷工作流:")))

        self.wf_dir_btn = QPushButton(tr("目录巡检"))
        self.wf_dir_btn.setFixedHeight(24)
        self.wf_dir_btn.clicked.connect(lambda: self._run_quick_workflow('dir_audit'))
        wf_row.addWidget(self.wf_dir_btn)

        self.wf_review_btn = QPushButton(tr("代码审阅"))
        self.wf_review_btn.setFixedHeight(24)
        self.wf_review_btn.clicked.connect(lambda: self._run_quick_workflow('code_review'))
        wf_row.addWidget(self.wf_review_btn)

        self.wf_change_btn = QPushButton(tr("变更说明"))
        self.wf_change_btn.setFixedHeight(24)
        self.wf_change_btn.clicked.connect(lambda: self._run_quick_workflow('change_summary'))
        wf_row.addWidget(self.wf_change_btn)
        wf_row.addStretch()
        layout.addLayout(wf_row)

        # 状态提示
        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color:#888;font-size:8.5pt;")
        layout.addWidget(self.status_lbl)

    # ── 鼠标点击：确保焦点转移到输入框 ──────────────────────────────────────
    def mousePressEvent(self, event):
        """点击面板空白区域时，强制 input_box 获取焦点（解决 QAxWidget 抢焦点问题）。"""
        super().mousePressEvent(event)
        self.main_window.activateWindow()
        self.input_box.setFocus(Qt.MouseFocusReason)

    # ── 事件过滤（Enter 发送 / 鼠标点击夺回焦点）────────────────────────────
    def eventFilter(self, obj, event):
        from PyQt5.QtCore import QEvent, QTimer
        if obj is self.input_box or obj is self.input_box.viewport():
            if event.type() == QEvent.KeyPress:
                if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                    if not (event.modifiers() & Qt.ShiftModifier):
                        self.send_message()
                        return True
            elif event.type() == QEvent.MouseButtonPress:
                # input_box 被点击时，确保主窗口重新激活并把焦点交给输入框
                # 用 singleShot(0) 让 Qt 先处理完点击事件再设焦点
                self.main_window.activateWindow()
                QTimer.singleShot(0, lambda: self.input_box.setFocus(Qt.MouseFocusReason))
                QTimer.singleShot(20, lambda: self.input_box.setFocus(Qt.MouseFocusReason))
        return super().eventFilter(obj, event)

    # ── 公共方法 ─────────────────────────────────────────────────────────────
    def update_context(self, path: str):
        """更新当前目录提示。"""
        if path:
            display = path if len(path) <= 55 else '…' + path[-52:]
            self.context_label.setText(tr("当前目录: {}").format(display))

    def submit_external_prompt(self, prompt: str) -> bool:
        """外部入口：将预置问题发送到 AI 对话。"""
        text = (prompt or "").strip()
        if not text:
            return False
        if not self.send_btn.isEnabled():
            show_toast(self, tr("提示"), tr("请等待当前 AI 任务完成后再启动新的工作流"), level="warning")
            return False
        self.input_box.setPlainText(text)
        self.send_message()
        return True

    def _default_workflow_templates(self):
        return {
            "dir_audit": (
                "请执行目录巡检工作流。\n"
                "目标目录: {target_dir}\n"
                "请按需先使用 [LIST_DIR] / [READ_FILE] 收集必要信息，再输出：\n"
                "1) 目录分层概览\n"
                "2) 核心模块与职责\n"
                "3) 潜在风险（命名、耦合、结构、可维护性）\n"
                "4) 建议优先阅读的文件/目录清单（Top 10）\n"
                "5) 后续可执行改进建议\n"
                "如需读取文件，优先读关键实现文件，不要无意义重复列目录。"
            ),
            "code_review": (
                "请执行代码审阅工作流。\n"
                "目标目录: {target_dir}\n"
                "请按需使用 [LIST_DIR] / [READ_FILE] / [GIT_STATUS] / [GIT_DIFF] 收集信息，"
                "重点识别真实风险而非样式问题，并输出：\n"
                "1) 高风险问题（按严重级别排序）\n"
                "2) 可能的行为回归点\n"
                "3) 缺失的测试建议\n"
                "4) 最小改动修复建议\n"
                "如果未发现明显问题，请明确说明并给出残余风险。"
            ),
            "change_summary": (
                "请执行变更说明工作流。\n"
                "目标目录: {target_dir}\n"
                "请优先使用 [GIT_STATUS] [GIT_DIFF] [GIT_LOG] 收集变更上下文，然后输出：\n"
                "1) 本次改动摘要（面向开发者）\n"
                "2) 用户可感知变化\n"
                "3) 风险与回滚关注点\n"
                "4) 建议的提交信息（Conventional Commits 风格）\n"
                "5) 建议的验证清单\n"
                "输出请简洁、可直接用于发布说明或 MR 描述。"
            ),
        }

    def _write_default_workflow_file(self, templates):
        payload = {
            "version": 1,
            "workflows": templates,
        }
        try:
            with open(self.workflow_file, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            debug_print(f"[AI] Failed to write ai_workflows.json: {e}")

    def _load_workflow_templates(self):
        defaults = self._default_workflow_templates()
        try:
            if not os.path.exists(self.workflow_file):
                self._write_default_workflow_file(defaults)
                return defaults

            with open(self.workflow_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            workflows = data.get('workflows') if isinstance(data, dict) else None
            if not isinstance(workflows, dict):
                self._write_default_workflow_file(defaults)
                return defaults

            merged = dict(defaults)
            for key, value in workflows.items():
                if isinstance(key, str) and isinstance(value, str) and value.strip():
                    merged[key] = value
            return merged
        except Exception as e:
            debug_print(f"[AI] Failed to load ai_workflows.json, using defaults: {e}")
            return defaults

    def _build_quick_workflow_prompt(self, workflow_key: str) -> str:
        target_dir = self._get_action_base_dir()
        template = self.workflow_templates.get(workflow_key, "") if isinstance(self.workflow_templates, dict) else ""
        if not template:
            return ""
        try:
            return template.format(target_dir=target_dir)
        except Exception:
            # 模板格式错误时回退默认模板
            fallback = self._default_workflow_templates().get(workflow_key, "")
            return fallback.format(target_dir=target_dir) if fallback else ""

    def _run_quick_workflow(self, workflow_key: str):
        name_map = {
            'dir_audit': tr('目录巡检'),
            'code_review': tr('代码审阅'),
            'change_summary': tr('变更说明'),
        }
        prompt = self._build_quick_workflow_prompt(workflow_key)
        if not prompt:
            return
        if self.submit_external_prompt(prompt):
            show_toast(self, tr("提示"), tr("已启动工作流: {}").format(name_map.get(workflow_key, workflow_key)), level="info")

    def _show_context_menu(self, pos):
        """显示右键菜单（复制、全选等）。"""
        menu = QMenu(self)
        
        copy_action = menu.addAction(tr("复制"))
        copy_action.triggered.connect(self.chat_display.copy)
        
        select_all_action = menu.addAction(tr("全选"))
        select_all_action.triggered.connect(self.chat_display.selectAll)
        
        menu.addSeparator()
        clear_action = menu.addAction(tr("清空聊天"))
        clear_action.triggered.connect(self.clear_chat)
        
        menu.exec_(self.chat_display.mapToGlobal(pos))

    def clear_chat(self):
        self.cancel_actions()
        self.messages.clear()
        self.chat_display.clear()
        self._delete_history()  # 清空时删除保存文件

    _MAX_AGENTIC_STEPS = 12

    def cancel_actions(self):
        self._request_cancelled = True
        if self.worker is not None:
            self.worker.requestInterruption()
        if self._action_worker is not None:
            self._action_worker.cancel()
        self.cancel_btn.setEnabled(False)
        self.status_lbl.setText(tr("取消中"))

    def has_running_actions(self):
        return self._action_worker is not None and self._action_worker.isRunning()

    def _cleanup_worker(self):
        # 用 sender() 避免与 agentic loop 新建的 worker 混淆
        finished_worker = self.sender()
        target = finished_worker if finished_worker else self.worker
        if self.worker is target:
            self.worker = None
            if self._request_cancelled and not self.has_running_actions():
                self.send_btn.setEnabled(True)
                self.status_lbl.setText(tr("已取消"))
        if target:
            try:
                target.deleteLater()
            except Exception:
                pass

    def cleanup(self):
        self._closing = True
        self.cancel_actions()
        try:
            self.input_box.removeEventFilter(self)
        except Exception:
            pass
        try:
            self.input_box.viewport().removeEventFilter(self)
        except Exception:
            pass

        worker = self.worker
        self.worker = None
        if worker:
            for signal_name, handler in (
                ('response_received', self._on_response),
                ('error_occurred', self._on_error),
                ('finished', self._cleanup_worker),
                ('token_received', self._on_token),
            ):
                try:
                    getattr(worker, signal_name).disconnect(handler)
                except Exception:
                    pass

            try:
                worker.requestInterruption()
            except Exception:
                pass

            if worker.isRunning():
                _retain_thread_until_finished(worker)
            else:
                try:
                    worker.deleteLater()
                except Exception:
                    pass

    def closeEvent(self, event):
        self.cleanup()
        super().closeEvent(event)

    def _sanitize_chat_message(self, message):
        if not isinstance(message, dict):
            return None
        role = str(message.get("role", "system") or "system")
        if role not in ("system", "user", "assistant"):
            role = "system"
        content = str(message.get("content", "") or "")
        if len(content) > MAX_CHAT_MESSAGE_CHARS:
            omitted = len(content) - MAX_CHAT_MESSAGE_CHARS
            content = content[:MAX_CHAT_MESSAGE_CHARS] + tr("\n\n【已截断 {} 个字符】").format(omitted)
        return {"role": role, "content": content}

    def _trim_chat_history(self):
        trimmed = []
        for message in self.messages[-MAX_CHAT_HISTORY_MESSAGES:]:
            sanitized = self._sanitize_chat_message(message)
            if sanitized is not None:
                trimmed.append(sanitized)
        # 总字符预算：超出时从最老的消息开始丢弃，至少保留最新 1 条
        while len(trimmed) > 1:
            if sum(len(m.get('content', '')) for m in trimmed) <= MAX_CONTEXT_TOTAL_CHARS:
                break
            trimmed.pop(0)
        self.messages = trimmed

    @staticmethod
    def _md_to_html(text: str) -> str:
        """Convert common Markdown patterns to HTML for chat display."""
        import html as _html
        import re as _re

        def inline_fmt(s):
            # s is already HTML-escaped; apply inline formatting
            parts = _re.split(r'(`[^`]+`)', s)
            result = []
            for part in parts:
                if part.startswith('`') and part.endswith('`') and len(part) > 1:
                    inner = part[1:-1]
                    result.append(
                        f'<code style="background:#f0f0f0;padding:1px 4px;'
                        f'border-radius:3px;font-family:Consolas,monospace;font-size:12px">'
                        f'{inner}</code>'
                    )
                else:
                    p = _re.sub(r'\*\*\*(.+?)\*\*\*', r'<b><i>\1</i></b>', part)
                    p = _re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', p)
                    p = _re.sub(r'\*(.+?)\*', r'<i>\1</i>', p)
                    result.append(p)
            return ''.join(result)

        lines = text.split('\n')
        out = []
        i = 0
        while i < len(lines):
            line = lines[i]

            # --- Fenced code block ---
            if line.startswith('```'):
                lang = _html.escape(line[3:].strip())
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].startswith('```'):
                    code_lines.append(_html.escape(lines[i]))
                    i += 1
                code = '\n'.join(code_lines)
                lang_html = (
                    f'<div style="color:#aaa;font-size:11px;margin-bottom:3px">{lang}</div>'
                    if lang else ''
                )
                out.append(
                    f'<div style="background:#1e1e1e;color:#d4d4d4;padding:8px 10px;margin:6px 0;'
                    f'border-radius:4px;font-family:Consolas,monospace;font-size:12px;'
                    f'white-space:pre;overflow-x:auto">{lang_html}{code}</div>'
                )
                i += 1
                continue

            # --- Horizontal rule ---
            if _re.match(r'^[-*_]{3,}\s*$', line):
                out.append('<hr style="border:0;border-top:1px solid #ccc;margin:8px 0">')
                i += 1
                continue

            # --- Heading ---
            m = _re.match(r'^(#{1,6})\s+(.+)', line)
            if m:
                level = len(m.group(1))
                sizes = {1: '18px', 2: '16px', 3: '14px', 4: '13px'}
                size = sizes.get(level, '13px')
                border = 'border-bottom:1px solid #ccc;padding-bottom:3px;' if level <= 2 else ''
                content = inline_fmt(_html.escape(m.group(2)))
                out.append(
                    f'<div style="font-size:{size};font-weight:bold;'
                    f'margin:10px 0 4px 0;{border}">{content}</div>'
                )
                i += 1
                continue

            # --- Table ---
            if line.strip().startswith('|') and line.strip().endswith('|'):
                tbl_lines = []
                while i < len(lines) and lines[i].strip().startswith('|') and lines[i].strip().endswith('|'):
                    tbl_lines.append(lines[i])
                    i += 1
                rows = []
                sep_idx = -1
                for ri, tl in enumerate(tbl_lines):
                    cells = [c.strip() for c in tl.strip()[1:-1].split('|')]
                    if all(_re.match(r'^:?-+:?$', c) for c in cells if c):
                        sep_idx = ri
                    else:
                        rows.append((ri, cells))
                if rows:
                    tbl = ['<table style="border-collapse:collapse;width:100%;margin:6px 0;font-size:12px">']
                    header_rows = [r for r in rows if sep_idx < 0 or r[0] < sep_idx]
                    data_rows   = [r for r in rows if sep_idx >= 0 and r[0] > sep_idx]
                    if header_rows:
                        tbl.append('<thead>')
                        for _, cells in header_rows:
                            tbl.append('<tr>')
                            for cell in cells:
                                tbl.append(
                                    f'<th style="padding:4px 8px;border:1px solid #ccc;'
                                    f'background:#e8e8e8;text-align:left">'
                                    f'{inline_fmt(_html.escape(cell))}</th>'
                                )
                            tbl.append('</tr></thead>')
                    if data_rows:
                        tbl.append('<tbody>')
                        for _, cells in data_rows:
                            tbl.append('<tr>')
                            for cell in cells:
                                tbl.append(
                                    f'<td style="padding:4px 8px;border:1px solid #ccc">'
                                    f'{inline_fmt(_html.escape(cell))}</td>'
                                )
                            tbl.append('</tr>')
                        tbl.append('</tbody>')
                    tbl.append('</table>')
                    out.append(''.join(tbl))
                continue

            # --- Unordered list ---
            if _re.match(r'^\s*[-*+]\s+', line):
                out.append('<ul style="margin:2px 0;padding-left:20px">')
                while i < len(lines) and _re.match(r'^\s*[-*+]\s+', lines[i]):
                    m2 = _re.match(r'^\s*[-*+]\s+(.+)', lines[i])
                    out.append(f'<li>{inline_fmt(_html.escape(m2.group(1)))}</li>')
                    i += 1
                out.append('</ul>')
                continue

            # --- Ordered list ---
            if _re.match(r'^\s*\d+\.\s+', line):
                out.append('<ol style="margin:2px 0;padding-left:20px">')
                while i < len(lines) and _re.match(r'^\s*\d+\.\s+', lines[i]):
                    m2 = _re.match(r'^\s*\d+\.\s+(.+)', lines[i])
                    out.append(f'<li>{inline_fmt(_html.escape(m2.group(1)))}</li>')
                    i += 1
                out.append('</ol>')
                continue

            # --- Empty line ---
            if not line.strip():
                out.append('<div style="height:4px"></div>')
                i += 1
                continue

            # --- Normal line ---
            out.append(f'<div>{inline_fmt(_html.escape(line))}</div>')
            i += 1

        return ''.join(out)

    def _build_bubble_html(self, role: str, content: str):
        import html as _html
        if role == "user":
            color, prefix, bg = "#1976D2", tr("👤 你"), "#E3F2FD"
        elif role == "assistant":
            color, prefix, bg = "#2E7D32", "🤖 AI", "#F1F8E9"
        else:
            color, prefix, bg = "#888", tr("ℹ 系统"), "#F5F5F5"

        if role == "assistant":
            body = self._md_to_html(content)
        else:
            body = _html.escape(content).replace('\n', '<br>')
        return (
            f'<div style="margin:4px 0;padding:6px 10px;background:{bg};'
            f'border-radius:6px;border-left:3px solid {color};">'
            f'<b style="color:{color};">{prefix}</b><br>{body}</div>'
        )

    def _rebuild_chat_display(self):
        self.chat_display.clear()
        for message in self.messages:
            self.chat_display.append(self._build_bubble_html(message.get("role", "system"), message.get("content", "")))
        self.chat_display.verticalScrollBar().setValue(
            self.chat_display.verticalScrollBar().maximum()
        )

    # 气泡过多时 QTextDocument 不会自动截断——超过此块数则重建显示
    _MAX_CHAT_DISPLAY_BLOCKS = 400

    def append_bubble(self, role: str, content: str, save_history=True):
        """向聊天区追加一条消息气泡（HTML 格式）。"""
        bubble = self._build_bubble_html(role, content)
        self.chat_display.append(bubble)
        self.chat_display.verticalScrollBar().setValue(
            self.chat_display.verticalScrollBar().maximum()
        )
        # 防止 QTextDocument 无限增长：超块数时重建（已有 _rebuild_chat_display 函数）
        if self.chat_display.document().blockCount() > self._MAX_CHAT_DISPLAY_BLOCKS:
            self._rebuild_chat_display()
        if save_history and not self._is_loading_history:
            self._schedule_save_history()  # 防抖写盘，避免每条气泡都触发 I/O

    def _schedule_save_history(self):
        """聊天记录防抖写盘（500ms 内多次追加只写一次磁盘）。"""
        if not hasattr(self, '_history_save_timer') or self._history_save_timer is None:
            from PyQt5.QtCore import QTimer
            self._history_save_timer = QTimer(self)
            self._history_save_timer.setSingleShot(True)
            self._history_save_timer.timeout.connect(self._save_history)
        self._history_save_timer.start(500)

    # ── 发送消息 ─────────────────────────────────────────────────────────────
    def send_message(self):
        if not self.send_btn.isEnabled() or self.has_running_actions():
            return
        text = self.input_box.toPlainText().strip()
        if not text:
            return

        # 检查用户输入是否包含文件路径或链接，提前处理
        self._handle_user_file_input(text)

        cfg = self.main_window.config.get("ai_chat", {})
        api_url = cfg.get("api_url", "").strip()
        api_key = cfg.get("api_key", "").strip()
        model   = cfg.get("model", "gpt-3.5-turbo").strip() or "gpt-3.5-turbo"

        if not api_url:
            self.append_bubble("system", tr("❌ 请先在 设置 → AI 助手 中填写 API 地址"))
            return

        self.input_box.clear()
        self.append_bubble("user", text)

        # 获取当前路径作为上下文
        self._action_base_dir = None
        current_path = self._get_action_base_dir()
        self._action_base_dir = current_path
        import weakref
        action_tab = self.main_window.get_current_tab_widget()
        self._action_tab_ref = weakref.ref(action_tab) if action_tab is not None else lambda: None

        # 构建 system prompt
        system_prompt = cfg.get("system_prompt", "").strip()
        if not system_prompt:
            system_prompt = (
                tr("你是一个智能文件管理助手，帮助用户管理文件系统、打开目录和运行脚本。\n") +
                tr("⚠️ 只有当用户明确要求‘切换/打开目录’时，才在回复末尾添加：[OPEN_DIR: 目录完整路径]。") +
                tr("普通问答不要输出任何操作指令。\n") +
                tr("⚠️ 只有当用户明确要求运行脚本时，才添加：[RUN_SCRIPT: 脚本完整路径]。\n") +
                tr("⚠️ 仅当用户明确要求时，才可使用以下指令：\n") +
                tr("[READ_FILE: 文件路径] 读取文件（系统自动分段读取大文件，无需手动指定偏移）；\n") +
                tr("[PATCH_FILE: 路径|旧文本|新文本] 局部修改已有文件（安全，仅替换目标内容段）；\n") +
                tr("⚠️ 修改已有文件时必须使用 PATCH_FILE，禁止用 WRITE_FILE 覆盖已有大文件；\n") +
                tr("⚠️ PATCH_FILE 使用规范：\n") +
                tr("  1. 旧文本只需包含被修改的行及前后各1-2行（足够唯一定位即可），不要复制整个函数体；\n") +
                tr("  2. 对同一文件做多个 PATCH_FILE 时，第二个 patch 的旧文本必须是第一个 patch 应用后的实际内容；\n") +
                tr("  3. 最小改动原则：新文本只改动必要的行，不要重写周围未变更的代码。\n") +
                tr("[WRITE_FILE: 文件路径|文件内容] 仅用于创建全新文件（已有文件禁止使用）；\n") +
                tr("[LIST_DIR: 目录路径] 列出目录；\n") +
                tr("[MKDIR: 目录路径] 创建目录；\n") +
                tr("[DELETE: 路径] 删除文件或目录（需用户确认）。\n") +
                tr("[GIT_STATUS: 仓库路径] 查看 git 状态；\n") +
                tr("[GIT_DIFF: 仓库路径] 查看未提交 diff；\n") +
                tr("[GIT_LOG: 仓库路径] 查看最近提交日志；\n") +
                tr("[GIT_BRANCH: 仓库路径] 查看分支列表；\n") +
                tr("[GIT_ADD: 仓库路径|目标] 暂存变更（需用户确认）；\n") +
                tr("[GIT_COMMIT: 仓库路径|提交信息] 提交变更（需用户确认）；\n") +
                tr("[GIT_SWITCH: 仓库路径|分支名] 切换分支（需用户确认）。\n") +
                tr("[GIT_RESTORE: 仓库路径|目标] 还原工作区改动（需用户确认）；\n") +
                tr("[GIT_RESET_SOFT: 仓库路径|目标提交] 软重置 HEAD（需用户确认）；\n") +
                tr("[GIT_PULL: 仓库路径|远程|分支] 拉取远程更新（需用户确认）；\n") +
                tr("[GIT_PUSH: 仓库路径|远程|分支] 推送本地提交（需用户确认）。\n") +
                tr("路径使用 Windows 格式，例如 D:\\project\\src。\n") +
                tr("可以在同一回复中包含多个操作命令。")
            )

        ctx_prefix = tr("[当前目录: {}]\n").format(current_path) if current_path else ""
        full_messages = [{"role": "system", "content": system_prompt}]
        full_messages.extend(self.messages)
        full_messages.append({"role": "user", "content": ctx_prefix + text})

        # 保存对话历史（不含 system，不含 ctx_prefix）
        self.messages.append({"role": "user", "content": text})
        self._trim_chat_history()

        # 启动工作线程（流式输出，每个 token 实时更新状态栏）
        self._agentic_steps = 0
        self._request_cancelled = False
        self.cancel_btn.setEnabled(True)
        self._last_agentic_feed = None  # 重复调用检测用
        self._streaming_chars = 0
        self.send_btn.setEnabled(False)
        self.status_lbl.setText(tr("⏳ AI 思考中…"))
        self.worker = ChatWorker(full_messages, api_url, api_key, model, stream=True, parent=self)
        self.worker.token_received.connect(self._on_token)
        self.worker.response_received.connect(self._on_response)
        self.worker.error_occurred.connect(self._on_error)
        self.worker.finished.connect(self._cleanup_worker)
        self.worker.start()

    def _handle_user_file_input(self, text):
        """检查用户输入中是否包含文件链接，直接打开。仅处理显式的 file:// 链接，避免误触发。"""
        import re
        # 只匹配 file:/// 或 file:// 链接（显式指示）
        file_links = re.findall(r'file:///?([^\s\]]+)', text)
        for link in file_links:
            path = link.replace('%20', ' ').replace('%5C', '\\')
            if os.path.isfile(path):
                # 打开文件所在目录
                folder = os.path.dirname(path)
                if os.path.isdir(folder):
                    self.append_bubble("system", tr("📂 打开目录: {}").format(folder))
                    self.main_window.add_new_tab(folder)
            elif os.path.isdir(path):
                # 直接打开目录
                self.append_bubble("system", tr("📂 打开目录: {}").format(path))
                self.main_window.add_new_tab(path)

    def _on_token(self, chunk: str):
        """流式 token 到达——实时更新聊天气泡（VS Code Copilot 效果）。"""
        self._streaming_buffer = getattr(self, '_streaming_buffer', '') + chunk
        step = getattr(self, '_agentic_steps', 0)
        label = tr("第 {} 轮").format(step + 1) if step > 0 else ""
        self.status_lbl.setText(tr("⏳ AI 推理中{}…").format(label))

        if not getattr(self, '_streaming_bubble_open', False):
            # 第一个 token：记录插入位置，后续刷新替换这段内容
            from PyQt5.QtGui import QTextCursor
            cursor = QTextCursor(self.chat_display.document())
            cursor.movePosition(QTextCursor.End)
            self._streaming_pre_char_pos = cursor.position()
            self._streaming_bubble_open = True

        # 50ms 防抖刷新（约 20fps），避免每 token 都重建 HTML
        if not hasattr(self, '_streaming_flush_timer') or self._streaming_flush_timer is None:
            from PyQt5.QtCore import QTimer
            self._streaming_flush_timer = QTimer(self)
            self._streaming_flush_timer.setSingleShot(True)
            self._streaming_flush_timer.timeout.connect(self._flush_streaming_bubble)
        self._streaming_flush_timer.start(50)

    def _flush_streaming_bubble(self):
        """将累积的 token 刷新到聊天气泡（替换式更新，保持样式和光标）。"""
        if not getattr(self, '_streaming_bubble_open', False):
            return
        from PyQt5.QtGui import QTextCursor
        doc = self.chat_display.document()
        cursor = QTextCursor(doc)
        cursor.setPosition(self._streaming_pre_char_pos)
        cursor.movePosition(QTextCursor.End, QTextCursor.KeepAnchor)
        cursor.removeSelectedText()
        buf = getattr(self, '_streaming_buffer', '')
        cursor.insertHtml(self._build_bubble_html("assistant", buf + "▌"))
        self.chat_display.verticalScrollBar().setValue(
            self.chat_display.verticalScrollBar().maximum()
        )

    def _close_streaming_bubble(self):
        """移除流式临时气泡（由 _on_response / _on_error 调用，后续追加最终气泡）。"""
        if not getattr(self, '_streaming_bubble_open', False):
            return
        if hasattr(self, '_streaming_flush_timer') and self._streaming_flush_timer:
            self._streaming_flush_timer.stop()
        from PyQt5.QtGui import QTextCursor
        doc = self.chat_display.document()
        cursor = QTextCursor(doc)
        cursor.setPosition(self._streaming_pre_char_pos)
        cursor.movePosition(QTextCursor.End, QTextCursor.KeepAnchor)
        cursor.removeSelectedText()
        self._streaming_bubble_open = False
        self._streaming_buffer = ""

    def _on_response(self, content: str):
        if self._closing or self._request_cancelled:
            return
        self.status_lbl.setText("")
        self._streaming_chars = 0
        self.messages.append({"role": "assistant", "content": content})
        self._trim_chat_history()
        worker = AiActionWorker(content, self._get_action_base_dir(), self)
        self._action_worker = worker
        worker.ui_requested.connect(self._handle_action_ui)
        worker.completed.connect(self._on_actions_completed)
        worker.finished.connect(self._action_thread_finished)
        worker.start()
        _retain_thread_until_finished(worker)

    def _handle_action_ui(self, request):
        try:
            if self._closing or self._request_cancelled or self.sender() is not self._action_worker:
                return
            kind, arguments = request['kind'], request['arguments']
            if kind == 'confirm':
                request['result'] = self._confirm_danger_action(*arguments)
            elif kind == 'preview':
                request['result'] = _confirm_file_preview(self, *arguments)
            elif kind == 'open':
                from PyQt5 import sip
                tab = getattr(self, '_action_tab_ref', lambda: None)()
                if tab is not None and not sip.isdeleted(tab):
                    tab.navigate_to(arguments[0], skip_async_check=True)
                else:
                    self.main_window.add_new_tab(*arguments)
                request['result'] = True
        finally:
            request['ready'].set()

    def _on_actions_completed(self, display_content, feedable, cancelled):
        if self._closing or self.sender() is not self._action_worker:
            return
        if cancelled or self._request_cancelled:
            self._close_streaming_bubble()
            self.status_lbl.setText(tr("已取消"))
            self.send_btn.setEnabled(True)
            return
        self._finish_response(display_content, feedable)

    def _action_thread_finished(self):
        if self.sender() is self._action_worker:
            self._action_worker = None

    def _finish_response(self, display_content, feedable):

        if feedable:
            # 中间轮次：不追加气泡，清空缓冲区让流式气泡原地刷新为“等待”状态
            last_feed = getattr(self, '_last_agentic_feed', None)
            stuck = (last_feed is not None and feedable == last_feed)
            self._last_agentic_feed = feedable
            if stuck:
                self.status_lbl.setText(tr("⚠️ 检测到重复工具调用，强制要求给出结论…"))
            if hasattr(self, '_streaming_flush_timer') and self._streaming_flush_timer:
                self._streaming_flush_timer.stop()
            self._streaming_buffer = ""
            if getattr(self, '_streaming_bubble_open', False):
                self._flush_streaming_bubble()
            self._start_agentic_step(feedable, force_final=stuck)
        else:
            # 最终结果：关闭流式气泡，追加正式气泡
            self._close_streaming_bubble()
            self.append_bubble("assistant", display_content)
            self._last_agentic_feed = None
            self.send_btn.setEnabled(True)
            self.cancel_btn.setEnabled(False)
    def _start_agentic_step(self, feedable_results: list, force_final: bool = False):
        """将工具调用结果回传给 AI，启动下一轮自动推理。
        force_final=True 时禁止 AI 继续调工具，强制给出结论。
        """
        cfg = self.main_window.config.get("ai_chat", {})
        api_url = cfg.get("api_url", "").strip()
        api_key = cfg.get("api_key", "").strip()
        model   = cfg.get("model", "gpt-3.5-turbo").strip() or "gpt-3.5-turbo"
        if not api_url:
            self.send_btn.setEnabled(True)
            return

        self._agentic_steps = getattr(self, '_agentic_steps', 0) + 1
        step = self._agentic_steps
        if self._request_cancelled or (self._MAX_AGENTIC_STEPS and step > self._MAX_AGENTIC_STEPS):
            self._close_streaming_bubble()
            self.send_btn.setEnabled(True)
            self.cancel_btn.setEnabled(False)
            self.status_lbl.setText(tr("已停止自动工具调用"))
            return

        # 工具结果限制单条 8000 字，防止上下文爆炸
        _MAX_FEED = 8000
        feed_parts = [r[:_MAX_FEED] + (f"\n…（已截断，共 {len(r)} 字）" if len(r) > _MAX_FEED else "")
                      for r in feedable_results]
        feed_text = "\n\n".join(feed_parts)

        if force_final:
            continuation = (
                tr("以上是最后一批工具调用结果。不要再调用任何工具指令，") +
                tr("请直接基于已收集的全部信息给出具体的结论、优化建议或解决方案。")
            )
        else:
            continuation = (
                tr("以上是工具调用结果。请高效利用工具，") +
                tr("优先阅读源代码文件（.c/.h）而非反复列目录，") +
                tr("并尽快基于收集到的信息给出具体结论或优化建议。")
            )
        feed_msg = {"role": "user", "content": feed_text + "\n\n" + continuation}

        # 保存进历史（下一轮 AI 能看到工具结果）
        self.messages.append(feed_msg)
        self._trim_chat_history()

        # ── Agentic 续推专用 system prompt ─────────────────────────────────
        # 与初始 send_message 的提示词分开：这里鼓励 AI 自由使用工具，
        # 而不是重复“只有用户明确要求时才使用”的限制性提示词。
        system_prompt = cfg.get("system_prompt", "").strip()
        if not system_prompt:
            if force_final:
                system_prompt = (
                    tr("你是一个智能文件管理助手。") +
                    tr("不能将任何工具指令放入回复，") +
                    tr("必须直接基于已收集的所有信息给出具体的结论、优化建议或解决方案。") +
                    tr("路径使用 Windows 格式。")
                )
            else:
                system_prompt = (
                    tr("你是一个智能文件管理助手，正在通过工具调用完成用户任务。\n") +
                    tr("你可以自由使用以下工具指令：\n") +
                    tr("[READ_FILE: 路径] 读文件（优先读 .c/.h 源代码）；") +
                    tr("[PATCH_FILE: 路径|旧文本|新文本] 局部修改文件（修改已有文件时必须用此，禁止用 WRITE_FILE 覆盖已有大文件）；") +
                    tr("[LIST_DIR: 路径] 列目录（仅在不知道源码位置时使用）；") +
                    tr("[OPEN_DIR: 路径] [MKDIR: 路径] [WRITE_FILE: 路径|内容]（仅用于创建新文件） [DELETE: 路径]。\n") +
                    tr("[GIT_STATUS: 路径] [GIT_DIFF: 路径] [GIT_LOG: 路径] [GIT_BRANCH: 路径] [GIT_ADD: 路径|目标] [GIT_COMMIT: 路径|提交信息] [GIT_SWITCH: 路径|分支名]。\n") +
                    tr("[GIT_RESTORE: 路径|目标] [GIT_RESET_SOFT: 路径|目标提交] [GIT_PULL: 路径|远程|分支] [GIT_PUSH: 路径|远程|分支]。\n") +
                    tr("ℹ️ 工作原则：得到目录列表后尽快选择源代码文件直接阅读，") +
                    tr("而不要反复展开子目录；修改文件时优先使用 PATCH_FILE 安全替换具体代码段。\n") +
                    tr("⚠️ PATCH_FILE 最小改动规范：\n") +
                    tr("  1. 旧文本只取被修改行及前后各1-2行（能唯一定位即可），不要复制整个函数；\n") +
                    tr("  2. 对同一文件连续多个 PATCH_FILE 时，后一个的旧文本必须反映前一个 patch 已应用后的文件内容；\n") +
                    tr("  3. 新文本只改动必要的行，保留其余未变行，使 diff 最小化。\n") +
                    tr("路径使用 Windows 格式，例如 D:\\project\\src。")
                )
        full_messages = [{"role": "system", "content": system_prompt}]
        full_messages.extend(self.messages)

        self._streaming_chars = 0
        self.status_lbl.setText(tr("⏳ AI 第 {} 轮推理中…").format(step))
        self.worker = ChatWorker(full_messages, api_url, api_key, model, stream=True, parent=self)
        self.worker.token_received.connect(self._on_token)
        self.worker.response_received.connect(self._on_response)
        self.worker.error_occurred.connect(self._on_error)
        self.worker.finished.connect(self._cleanup_worker)
        self.worker.start()

    def _on_error(self, msg: str):
        if self._closing or self._request_cancelled:
            return
        self._close_streaming_bubble()   # 移除未完成的流式气泡
        self.send_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.status_lbl.setText("")
        self.append_bubble("system", tr("❌ 请求失败: {}").format(msg))

    # ── 拖拽文件支持 ──────────────────────────────────────────────────────────
    def dragEnterEvent(self, event):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dragMoveEvent(self, event):
        """拖拽移动事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        """拖拽释放事件 - 读取文件内容"""
        if event.mimeData().hasUrls():
            auto_analyze = not bool(self.input_box.toPlainText().strip())
            loaded_files = []
            loaded_dirs = []
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    local_path = url.toLocalFile()
                    if os.path.isfile(local_path):
                        if self._read_and_display_file(local_path):
                            loaded_files.append(local_path)
                    elif os.path.isdir(local_path):
                        dir_preview = self._read_and_display_directory(local_path)
                        if dir_preview:
                            loaded_dirs.append((local_path, dir_preview))
                else:
                    # 尝试从 URL 字符串中提取本地路径
                    url_str = url.toString()
                    if url_str.startswith('file:///'):
                        from urllib.parse import unquote
                        local_path = unquote(url_str[8:])
                        if os.name == 'nt' and local_path.startswith('/'):
                            local_path = local_path[1:]
                        if os.path.isfile(local_path):
                            if self._read_and_display_file(local_path):
                                loaded_files.append(local_path)
                        elif os.path.isdir(local_path):
                            dir_preview = self._read_and_display_directory(local_path)
                            if dir_preview:
                                loaded_dirs.append((local_path, dir_preview))

            # 输入框为空时，自动发起分析；若用户已手打内容，则不打断用户操作
            if auto_analyze and (loaded_files or loaded_dirs):
                if loaded_dirs and not loaded_files:
                    if len(loaded_dirs) == 1:
                        dir_path, dir_preview = loaded_dirs[0]
                        prompt = (
                            f"请分析我刚拖拽的目录结构：{dir_path}\n"
                            "请基于下面的目录结构预览输出：\n"
                            "1) 目录分层概览\n2) 可能的核心模块\n3) 潜在风险点（过深目录/职责混杂/命名问题）\n"
                            "4) 建议的阅读顺序与重构建议\n\n"
                            f"目录结构预览:\n{dir_preview}"
                        )
                    else:
                        dir_lines = "\n\n".join([f"目录: {p}\n{pv}" for p, pv in loaded_dirs[:3]])
                        prompt = (
                            f"请综合分析我刚拖拽的 {len(loaded_dirs)} 个目录结构。\n"
                            "请输出：\n1) 各目录职责\n2) 共同结构模式\n3) 风险点\n4) 建议的查看优先级\n\n"
                            f"目录结构预览:\n{dir_lines}"
                        )
                elif loaded_files and not loaded_dirs:
                    if len(loaded_files) == 1:
                        prompt = tr("请分析我刚拖拽的文件：{}\n请按以下结构输出：\n1) 文件用途\n2) 关键逻辑\n3) 风险点\n4) 建议下一步查看的相关文件").format(loaded_files[0])
                    else:
                        file_list = "\n".join(f"- {p}" for p in loaded_files[:8])
                        prompt = tr("请综合分析我刚拖拽的 {} 个文件：\n{}\n请按以下结构输出：\n1) 文件分工\n2) 关键逻辑链路\n3) 主要风险点\n4) 建议的阅读顺序").format(len(loaded_files), file_list)
                else:
                    file_list = "\n".join(f"- {p}" for p in loaded_files[:6])
                    dir_list = "\n\n".join([f"目录: {p}\n{pv}" for p, pv in loaded_dirs[:2]])
                    prompt = (
                        "请综合分析我刚拖拽的文件和目录。\n"
                        "请输出：\n1) 文件与目录的关系\n2) 核心实现链路\n3) 风险点\n4) 建议阅读顺序\n\n"
                        f"文件列表:\n{file_list}\n\n目录结构预览:\n{dir_list}"
                    )
                if self.submit_external_prompt(prompt):
                    pass
            event.acceptProposedAction()

    def _read_and_display_directory(self, dir_path):
        """读取目录结构摘要并在聊天窗口中显示。返回用于 AI 提示词的结构文本。"""
        if not os.path.isdir(dir_path):
            return ""

        max_depth = 2
        max_total_entries = 180
        max_entries_per_dir = 30
        total_entries = 0
        total_dirs = 0
        total_files = 0
        truncated = False
        lines = []

        def _walk(path, depth, indent):
            nonlocal total_entries, total_dirs, total_files, truncated
            if depth > max_depth or truncated:
                return
            try:
                with os.scandir(path) as it:
                    entries = sorted(list(it), key=lambda e: (not e.is_dir(follow_symlinks=False), e.name.lower()))
            except Exception as e:
                lines.append(f"{indent}⚠ 无法读取: {e}")
                return

            if len(entries) > max_entries_per_dir:
                visible_entries = entries[:max_entries_per_dir]
                hidden = len(entries) - max_entries_per_dir
            else:
                visible_entries = entries
                hidden = 0

            for entry in visible_entries:
                if total_entries >= max_total_entries:
                    truncated = True
                    break
                try:
                    is_dir = entry.is_dir(follow_symlinks=False)
                except Exception:
                    is_dir = False
                icon = "📁" if is_dir else "📄"
                lines.append(f"{indent}{icon} {entry.name}")
                total_entries += 1
                if is_dir:
                    total_dirs += 1
                    _walk(entry.path, depth + 1, indent + "  ")
                else:
                    total_files += 1

            if hidden > 0 and not truncated:
                lines.append(f"{indent}... 其余 {hidden} 项省略")

        root_name = os.path.basename(os.path.normpath(dir_path)) or dir_path
        lines.append(f"📁 {root_name}")
        _walk(dir_path, 0, "  ")

        if truncated:
            lines.append(f"... 已达到预览上限（最多 {max_total_entries} 项）")

        preview = "\n".join(lines)
        summary = f"目录: {dir_path}\n子目录: {total_dirs}，文件: {total_files}\n\n{preview}"
        self.append_bubble("system", f"📁 目录结构预览\n{'='*50}\n{summary}")
        return summary
    
    def _read_and_display_file(self, file_path):
        """读取文件并在聊天窗口中显示"""
        if not os.path.isfile(file_path):
            self.append_bubble("system", tr("❌ 文件不存在: {}").format(file_path))
            return False
        
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        if file_size_mb > 10:  # 超过10MB不读取
            self.append_bubble("system", tr("⚠️ 文件太大（{:.1f}MB），无法读取").format(file_size_mb))
            return False
        
        try:
            # 尝试以 UTF-8 编码读取
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            try:
                # 降级为 GBK/GB2312
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()
            except Exception:
                try:
                    # 最后尝试二进制显示
                    with open(file_path, 'rb') as f:
                        raw = f.read(1000)
                    content = f"[二进制文件，前 {len(raw)} 字节]\n{raw[:200]}"
                except Exception as e:
                    self.append_bubble("system", tr("❌ 无法读取文件: {}").format(e))
                    return False
        
        # 截断长内容（超过5000字符）
        max_chars = 5000
        if len(content) > max_chars:
            content = content[:max_chars] + f"\n\n【省略 {len(content) - max_chars} 个字符】"
        
        # 获取文件名
        file_name = os.path.basename(file_path)
        
        # 显示文件内容
        display_text = f"📄 {file_name}\n" + "="*50 + "\n" + content
        self.append_bubble("system", display_text)
        return True

    # ── 聊天记录持久化 ──────────────────────────────────────────────────────────
    def _save_history(self):
        """保存聊天记录到 JSON 文件。"""
        import json
        try:
            before_len = len(self.messages)
            self._trim_chat_history()
            if len(self.messages) != before_len:
                self._rebuild_chat_display()
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.messages, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(tr("保存聊天记录失败: {}").format(e))

    def _load_history(self):
        """从 JSON 文件加载聊天记录。"""
        import json
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    loaded_messages = json.load(f)
                self.messages = loaded_messages if isinstance(loaded_messages, list) else []
                self._trim_chat_history()
                # 重新显示所有消息
                self._is_loading_history = True
                self._rebuild_chat_display()
                self._is_loading_history = False
                self._save_history()
            except Exception as e:
                self._is_loading_history = False
                print(tr("加载聊天记录失败: {}").format(e))

    def _delete_history(self):
        """删除保存的聊天记录文件。"""
        try:
            if os.path.exists(self.history_file):
                os.remove(self.history_file)
        except Exception as e:
            print(tr("删除聊天记录文件失败: {}").format(e))

    def _get_action_base_dir(self):
        """获取动作执行的基准目录（当前标签目录）。"""
        captured = getattr(self, '_action_base_dir', None)
        if captured is not None:
            return captured
        try:
            tab = self.main_window.get_current_tab_widget()
            p = getattr(tab, 'current_path', '') if tab else ''
            if isinstance(p, str) and os.path.isabs(p) and not p.startswith(('shell:', '::')):
                return os.path.normpath(p)
        except Exception:
            pass
        return ''

    def _resolve_action_path(self, raw_path: str):
        return _AiActionMethods._resolve_action_path(self, raw_path)

    def _confirm_danger_action(self, title: str, text: str) -> bool:
        """危险操作二次确认。"""
        try:
            from PyQt5.QtWidgets import QMessageBox
            ret = QMessageBox.question(
                self,
                title,
                text,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            return ret == QMessageBox.Yes
        except Exception:
            # 任何异常都按拒绝处理，确保安全
            return False

    def _resolve_git_repo_dir(self, raw_path: str):
        return _AiActionMethods._resolve_git_repo_dir(self, raw_path)

    def _resolve_git_target(self, repo_dir, target):
        return _AiActionMethods._resolve_git_target(self, repo_dir, target)

    def _run_git_command(self, repo_dir: str, git_args: list, timeout_sec: int = 20):
        return _AiActionMethods._run_git_command(self, repo_dir, git_args, timeout_sec)

    # ── 解析并执行 AI 操作命令 ────────────────────────────────────────────────
    def _apply_actions(self, content: str) -> tuple:
        return _AiActionMethods._apply_actions(self, content)

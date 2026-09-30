"""AI action parsing and execution, separated from chat presentation."""

import os
import subprocess
import tempfile
import threading
import time

from PyQt5.QtCore import QThread, pyqtSignal

from .paths import translate_common_path
from .i18n import tr
from .system import launch_detached
from .fileops import _atomic_reviewed_write, _confirm_file_preview


def _action_ui(host, kind, *arguments):
    dispatch = getattr(host, '_request_ui', None)
    if dispatch is not None:
        return dispatch(kind, *arguments)
    if kind == 'preview':
        return _confirm_file_preview(host, *arguments)
    if kind == 'open':
        return host.main_window.add_new_tab(*arguments)
    raise ValueError(kind)


class _AiActionMethods:
    def _resolve_action_path(self, raw_path: str):
        """解析动作路径，限制在当前目录范围内。"""
        p = (raw_path or '').strip().strip('"\'')
        if not p:
            return None, tr("空路径")

        p = p.replace('/', '\\')
        base_dir = self._get_action_base_dir()
        if not base_dir:
            return None, tr("当前标签不是可操作的文件目录")

        # 相对路径按当前目录解析
        if not os.path.isabs(p):
            p = os.path.join(base_dir, p)

        p = translate_common_path(os.path.normpath(p))

        try:
            # 解析目录链接/符号链接后再比较，防止经由当前目录内的链接访问目录外文件
            base_real = os.path.normcase(os.path.realpath(base_dir)).rstrip(os.sep)
            path_real = os.path.normcase(os.path.realpath(p))
            if not (path_real == base_real or path_real.rstrip(os.sep) == base_real
                    or path_real.startswith(base_real + os.sep)):
                return None, tr("超出当前目录范围: {}").format(p)
        except Exception:
            return None, tr("路径非法: {}").format(p)

        return p, None

    def _resolve_git_repo_dir(self, raw_path: str):
        """解析 Git 操作目录，限制在当前目录范围内。"""
        path, err = self._resolve_action_path(raw_path)
        if err:
            return None, err
        repo_dir = path if os.path.isdir(path) else os.path.dirname(path)
        if not repo_dir or not os.path.isdir(repo_dir):
            return None, tr("❌ 目录不存在: {}").format(path)
        return repo_dir, None

    def _resolve_git_target(self, repo_dir, target):
        if target.startswith(':') or any(character in target for character in ('*', '?', '[', '\x00', '\n', '\r')):
            return None, tr("Git 目标必须是当前目录内的明确路径")
        return self._resolve_action_path(os.path.join(repo_dir, target))

    def _run_git_command(self, repo_dir: str, git_args: list, timeout_sec: int = 20):
        """执行 Git 命令并返回 (ok, stdout, stderr)。"""
        import shutil
        if not shutil.which("git"):
            return False, "", tr("未检测到 Git，请先安装 Git for Windows")
        command = git_args[0]
        operands = git_args[2:] if command == 'reset' else git_args[1:]
        if command in ('switch', 'checkout', 'reset', 'pull', 'push') and any(value.startswith('-') for value in operands):
            return False, "", tr("Git 参数不能是命令选项")
        if command in ('commit', 'switch', 'checkout', 'reset', 'pull', 'push'):
            ok, root, error = _AiActionMethods._run_git_command(self, repo_dir, ['rev-parse', '--show-toplevel'])
            if not ok:
                return False, '', error
            _, error = self._resolve_action_path(root)
            if error:
                return False, '', tr("整仓库操作要求仓库根目录位于授权目录内")
        cancelled = getattr(self, '_cancel_event', None) or threading.Event()
        environment = dict(os.environ, GIT_TERMINAL_PROMPT='0', GCM_INTERACTIVE='Never', GIT_LITERAL_PATHSPECS='1')
        try:
            with tempfile.TemporaryFile() as output, tempfile.TemporaryFile() as errors:
                proc = subprocess.Popen(
                    ["git", "--no-pager", "-C", repo_dir] + list(git_args),
                    stdin=subprocess.DEVNULL, stdout=output, stderr=errors, env=environment,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                deadline = time.monotonic() + timeout_sec
                try:
                    while proc.poll() is None:
                        if cancelled.is_set():
                            raise RuntimeError(tr("操作已取消"))
                        if time.monotonic() >= deadline:
                            raise subprocess.TimeoutExpired(git_args, timeout_sec)
                        if os.fstat(output.fileno()).st_size + os.fstat(errors.fileno()).st_size > 2 * 1024 * 1024:
                            raise RuntimeError(tr("Git 输出超过 2 MB 上限"))
                        cancelled.wait(0.05)
                finally:
                    if proc.poll() is None:
                        proc.kill()
                    proc.wait()
                output.seek(0)
                errors.seek(0)
                stdout = output.read(65536).decode('utf-8', errors='replace').strip()
                stderr = errors.read(65536).decode('utf-8', errors='replace').strip()
            if proc.returncode != 0:
                return False, stdout, (stderr or tr("❌ Git 执行失败: {}").format(" ".join(git_args)))
            return True, stdout, stderr
        except subprocess.TimeoutExpired:
            return False, "", tr("❌ Git 执行超时: {}").format(" ".join(git_args))
        except Exception as e:
            return False, "", tr("❌ Git 执行失败: {}").format(e)

    def _apply_actions(self, content: str) -> tuple:
        """解析 AI 回复中的操作标记并按出现顺序执行。
        返回 (display_content, feedable)：
        - display_content: 附带执行注释的显示文本
        - feedable: READ_FILE/LIST_DIR 结果列表，供 agentic loop 回传给 AI
        """
        import re, subprocess
        notes = []

        # ── 收集所有指令及其在文本中的位置，按顺序执行 ──────────────────────
        feedable = []  # 可回传给 AI 的工具结果（READ_FILE / LIST_DIR）
        actions = []  # list of (start_pos, action_type, match_obj)

        patterns = {
            'OPEN_DIR':   re.compile(r'\[OPEN_DIR:\s*([^\]]+)\]'),
            'RUN_SCRIPT': re.compile(r'\[RUN_SCRIPT:\s*([^\]]+)\]'),
            'LIST_DIR':   re.compile(r'\[LIST_DIR:\s*([^\]]+)\]'),
            'READ_FILE':  re.compile(r'\[READ_FILE:\s*([^\]]+)\]'),
            'MKDIR':      re.compile(r'\[MKDIR:\s*([^\]]+)\]'),
            'GIT_STATUS': re.compile(r'\[GIT_STATUS:\s*([^\]]+)\]'),
            'GIT_DIFF':   re.compile(r'\[GIT_DIFF:\s*([^\]]+)\]'),
            'GIT_LOG':    re.compile(r'\[GIT_LOG:\s*([^\]]+)\]'),
            'GIT_BRANCH': re.compile(r'\[GIT_BRANCH:\s*([^\]]+)\]'),
            # PATCH_FILE / WRITE_FILE 内容可含代码中的 ] (如 arr[i])，
            # 用 .*? + 行末 lookahead 确保只在真正的指令结束处停止
            'PATCH_FILE': re.compile(r'\[PATCH_FILE:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'WRITE_FILE': re.compile(r'\[WRITE_FILE:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'DELETE':     re.compile(r'\[DELETE:\s*([^\]]+)\]'),
            'GIT_ADD':    re.compile(r'\[GIT_ADD:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'GIT_COMMIT': re.compile(r'\[GIT_COMMIT:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'GIT_SWITCH': re.compile(r'\[GIT_SWITCH:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'GIT_RESTORE': re.compile(r'\[GIT_RESTORE:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'GIT_RESET_SOFT': re.compile(r'\[GIT_RESET_SOFT:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'GIT_PULL': re.compile(r'\[GIT_PULL:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
            'GIT_PUSH': re.compile(r'\[GIT_PUSH:\s*(.*?)\][ \t]*(?=\n|$)', re.DOTALL),
        }
        for atype, pat in patterns.items():
            for m in pat.finditer(content):
                actions.append((m.start(), atype, m))

        # 按出现位置排序，保证 WRITE_FILE 在 RUN_SCRIPT 之前（若 AI 如此安排）
        actions.sort(key=lambda x: x[0])

        for _, atype, m in actions:
            cancelled = getattr(self, '_cancel_event', None)
            if cancelled is not None and cancelled.is_set():
                break

            # ── OPEN_DIR ────────────────────────────────────────────────────
            if atype == 'OPEN_DIR':
                raw = m.group(1).strip().strip('"\'')
                if not os.path.isabs(raw) and self._get_action_base_dir():
                    raw = os.path.join(self._get_action_base_dir(), raw)
                path = translate_common_path(os.path.normpath(raw))
                if os.path.isdir(path):
                    try:
                        _action_ui(self, 'open', path)
                        notes.append(tr("✅ 已打开目录: {}").format(path))
                    except Exception as e:
                        notes.append(tr("❌ 打开目录失败: {}（{}）").format(path, e))
                else:
                    notes.append(tr("⚠️ 目录不存在: {}").format(path))

            # ── RUN_SCRIPT ──────────────────────────────────────────────────
            elif atype == 'RUN_SCRIPT':
                script, err = self._resolve_action_path(m.group(1))
                if err:
                    notes.append(tr("❌ 运行脚本失败: {}（{}）").format(m.group(1), err))
                    continue
                if os.path.isfile(script):
                    try:
                        if not self._confirm_danger_action(
                            tr("确认运行脚本"),
                            tr("AI 请求运行脚本：\n{}\n\n是否继续？").format(script)
                        ):
                            notes.append(tr("⏸ 已取消运行脚本: {}").format(script))
                            continue
                        checked, error = self._resolve_action_path(m.group(1))
                        if error or checked != script:
                            raise ValueError(tr("文件路径已变化"))
                        launch_detached(script if isinstance(script, list) else [script], cwd=os.path.dirname(script))
                        notes.append(tr("✅ 已启动脚本: {}").format(script))
                    except Exception as e:
                        notes.append(tr("❌ 运行脚本失败: {}（{}）").format(script, e))
                else:
                    notes.append(tr("⚠️ 脚本不存在: {}").format(script))

            # ── LIST_DIR ────────────────────────────────────────────────────
            elif atype == 'LIST_DIR':
                raw = m.group(1)
                path, err = self._resolve_action_path(raw)
                if err:
                    _emsg = tr("❌ 列目录失败: {}").format(err)
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                if not os.path.isdir(path):
                    _emsg = tr("❌ 目录不存在: {}").format(path)
                    if os.path.isfile(path):
                        _emsg += tr("（这是文件，请用 READ_FILE）")
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                try:
                    import itertools
                    with os.scandir(path) as entries:
                        items = [entry.name for entry in itertools.islice(entries, 51)]
                    preview = "\n".join(items[:50]) if items else tr("(空目录)")
                    if len(items) > 50:
                        preview += tr("\n... 其余项目省略")
                    notes.append(tr("📁 目录列表: {}\n{}").format(path, preview))
                    feedable.append(tr("[LIST_DIR 结果] 目录 {}:\n{}").format(path, preview))
                except Exception as e:
                    notes.append(tr("❌ 列目录失败: {}（{}）").format(path, e))

            # ── READ_FILE ───────────────────────────────────────────────────
            elif atype == 'READ_FILE':
                raw = m.group(1)
                parts = raw.split('|')
                raw_path = parts[0].strip()
                start_offset = 0
                if len(parts) > 1:
                    try:
                        start_offset = int(parts[1].strip())
                    except ValueError:
                        pass
                path, err = self._resolve_action_path(raw_path)
                if err:
                    _emsg = tr("❌ 读取失败: {}").format(err)
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                if not os.path.isfile(path):
                    _emsg = tr("❌ 文件不存在: {}").format(path)
                    if os.path.isdir(path):
                        _emsg += tr("（这是目录，请用 LIST_DIR 列目录，或直接指定具体 .c/.h 文件路径）")
                    notes.append(_emsg); feedable.append(_emsg)
                    continue
                CHUNK = 5000
                MAX_TOTAL = 50000
                try:
                    enc = 'utf-8'
                    try:
                        with open(path, 'r', encoding='utf-8') as _probe_f:
                            _probe_f.read(1)
                    except UnicodeDecodeError:
                        enc = 'gbk'
                    file_size = os.path.getsize(path)
                    all_text = []
                    offset = start_offset
                    with open(path, 'r', encoding=enc, errors='replace') as f:
                        f.seek(offset)
                        while True:
                            chunk = f.read(CHUNK)
                            if not chunk:
                                break
                            all_text.append(chunk)
                            offset += len(chunk.encode(enc, errors='replace'))
                            if sum(len(t) for t in all_text) >= MAX_TOTAL:
                                break
                    full_text = ''.join(all_text)
                    total_read = sum(len(t) for t in all_text)
                    truncated = (start_offset + total_read) < file_size
                    info = tr("📄 文件内容（{}-{}/{}字节，{}）: {}").format(start_offset, start_offset+total_read, file_size, enc, path)
                    if truncated:
                        info += tr("\n⚠️ 文件过大，仅读取前 {} 字符，剩余 {} 字节未读").format(MAX_TOTAL, file_size - start_offset - total_read)
                    notes.append(f"{info}\n{full_text}")
                    feedable.append(f"{info}\n{full_text}")
                except Exception as e:
                    notes.append(tr("❌ 读取失败: {}（{}）").format(path, e))

            # ── MKDIR ───────────────────────────────────────────────────────
            elif atype == 'MKDIR':
                raw = m.group(1)
                path, err = self._resolve_action_path(raw)
                if err:
                    notes.append(tr("❌ 创建目录失败: {}").format(err))
                    continue
                try:
                    os.makedirs(path, exist_ok=True)
                    notes.append(tr("✅ 已创建目录: {}").format(path))
                except Exception as e:
                    notes.append(tr("❌ 创建目录失败: {}（{}）").format(path, e))

            # ── PATCH_FILE ──────────────────────────────────────────────────
            elif atype == 'PATCH_FILE':
                raw = m.group(1)
                parts = raw.split('|', 2)
                if len(parts) < 3:
                    notes.append(tr("❌ PATCH_FILE 格式: [PATCH_FILE: 路径|旧文本|新文本]"))
                    continue
                raw_path, old_text, new_text = parts
                path, err = self._resolve_action_path(raw_path.strip())
                if err:
                    notes.append(tr("❌ 补丁失败: {}").format(err))
                    continue
                if not os.path.isfile(path):
                    notes.append(tr("❌ 文件不存在: {}").format(path))
                    continue
                try:
                    if os.path.getsize(path) > 5 * 1024 * 1024:
                        raise ValueError(tr("预览文件上限为 5 MB"))
                    with open(path, 'rb') as source:
                        original_bytes = source.read()
                    enc = 'utf-8-sig' if original_bytes.startswith(b'\xef\xbb\xbf') else 'utf-8'
                    try:
                        file_src = original_bytes.decode(enc)
                    except UnicodeDecodeError:
                        enc = 'gbk'
                        file_src = original_bytes.decode(enc)
                    if not old_text:
                        raise ValueError(tr("补丁目标文本不能为空"))
                    count = file_src.count(old_text)
                    if count == 0:
                        notes.append(tr("❌ 补丁失败：在 {} 中未找到目标文本").format(path))
                        continue
                    if count > 1:
                        notes.append(tr("补丁目标不唯一，请提供更多上下文"))
                        continue
                    patched = file_src.replace(old_text, new_text, 1)
                    if not _action_ui(self, 'preview', tr("确认 AI 补丁"), path, file_src, patched):
                        notes.append(tr("已取消补丁: ") + path)
                        continue
                    checked_path, checked_error = self._resolve_action_path(raw_path.strip())
                    if checked_error or checked_path != path:
                        raise ValueError(tr("文件路径已变化"))
                    _atomic_reviewed_write(path, original_bytes, patched.encode(enc))
                    notes.append(tr("✅ 补丁成功: {}").format(path))
                except Exception as e:
                    notes.append(tr("❌ 补丁失败: {}（{}）").format(path, e))

            # ── WRITE_FILE ──────────────────────────────────────────────────
            elif atype == 'WRITE_FILE':
                raw = m.group(1)
                if '|' not in raw:
                    notes.append(tr("❌ 写入失败: 格式应为 [WRITE_FILE: 路径|内容]"))
                    continue
                raw_path, file_content = raw.split('|', 1)
                path, err = self._resolve_action_path(raw_path)
                if err:
                    notes.append(tr("❌ 写入失败: {}").format(err))
                    continue
                try:
                    file_exists = os.path.exists(path)
                    # 已有文件且较大时，拒绝全量覆盖，引导用 PATCH_FILE
                    if file_exists and os.path.getsize(path) > 5000:
                        notes.append(
                            f"❌ 已拒绝覆盖：{os.path.basename(path)} 已存在且大于 5KB，" +
                            tr("全量覆盖会丢失未读内容。") +
                            tr("请改用 [PATCH_FILE: 路径|旧文本|新文本] 进行局部修改。")
                        )
                        continue
                    original_bytes = None
                    before = ''
                    if file_exists:
                        with open(path, 'rb') as source:
                            original_bytes = source.read()
                        before = original_bytes.decode('utf-8-sig')
                    if not _action_ui(self, 'preview', tr("确认 AI 写入"), path, before, file_content):
                        notes.append(tr("已取消写入: ") + path)
                        continue
                    checked_path, checked_error = self._resolve_action_path(raw_path)
                    if checked_error or checked_path != path:
                        raise ValueError(tr("文件路径已变化"))
                    _atomic_reviewed_write(path, original_bytes, file_content.encode('utf-8'))
                    notes.append(tr("✅ 已写入文件: {}").format(path))
                except Exception as e:
                    notes.append(tr("❌ 写入失败: {}（{}）").format(path, e))

            # ── DELETE ──────────────────────────────────────────────────────
            elif atype == 'DELETE':
                raw = m.group(1).strip()
                path, err = self._resolve_action_path(raw)
                if err:
                    notes.append(tr("❌ 删除失败: {}").format(err))
                    continue
                if not os.path.exists(path):
                    notes.append(tr("❌ 路径不存在: {}").format(path))
                    continue
                is_dir = os.path.isdir(path)
                type_label = tr("目录（及其所有内容）") if is_dir else tr("文件")
                if not self._confirm_danger_action(
                    tr("确认删除"),
                    tr("AI 请求将{}移入回收站：\n{}\n\n是否继续？").format(type_label, path)
                ):
                    notes.append(tr("⏸ 已取消删除: {}").format(path))
                    continue
                try:
                    from send2trash import send2trash
                    checked, error = self._resolve_action_path(raw)
                    if error or checked != path:
                        raise ValueError(tr("文件路径已变化"))
                    send2trash(path)
                    notes.append(tr("✅ 已删除{}: {}").format(type_label, path))
                except Exception as e:
                    notes.append(tr("❌ 删除失败: {}（{}）").format(path, e))

            # ── GIT_STATUS / GIT_DIFF / GIT_LOG / GIT_BRANCH ───────────────
            elif atype in ('GIT_STATUS', 'GIT_DIFF', 'GIT_LOG', 'GIT_BRANCH'):
                raw = m.group(1).strip()
                repo_dir, err = self._resolve_git_repo_dir(raw)
                if err:
                    _emsg = tr("❌ Git 执行失败: {}").format(err)
                    notes.append(_emsg)
                    feedable.append(_emsg)
                    continue

                git_map = {
                    'GIT_STATUS': (['status', '--short', '--branch'], tr("✅ Git 状态 ({})\n{}")),
                    'GIT_DIFF': (['diff'], tr("✅ Git Diff ({})\n{}")),
                    'GIT_LOG': (['log', '--oneline', '-n', '20'], tr("✅ Git Log ({})\n{}")),
                    'GIT_BRANCH': (['branch', '--all', '--verbose', '--no-abbrev'], tr("✅ Git Branch ({})\n{}")),
                }
                args, fmt = git_map[atype]
                ok, out, err_msg = self._run_git_command(repo_dir, args)
                if ok:
                    text = out if out else tr("⚠️ Git 无输出")
                    if len(text) > 30000:
                        text = text[:30000] + tr("\n\n【已截断 {} 个字符】").format(len(text) - 30000)
                    msg = fmt.format(repo_dir, text)
                    notes.append(msg)
                    feedable.append(msg)
                else:
                    _emsg = tr("❌ Git 执行失败: {}").format(err_msg)
                    notes.append(_emsg)
                    feedable.append(_emsg)

            # ── GIT_ADD ─────────────────────────────────────────────────────
            elif atype == 'GIT_ADD':
                raw = m.group(1)
                parts = raw.split('|', 1)
                if len(parts) < 2:
                    notes.append(tr("❌ Git 指令格式应为 [GIT_ADD: 路径|目标]"))
                    continue
                repo_raw, target = parts[0].strip(), parts[1].strip() or '.'
                repo_dir, err = self._resolve_git_repo_dir(repo_raw)
                if not err:
                    _, err = _AiActionMethods._resolve_git_target(self, repo_dir, target)
                if err:
                    notes.append(tr("❌ Git 执行失败: {}").format(err))
                    continue
                if not self._confirm_danger_action(
                    tr("确认 Git 暂存"),
                    tr("AI 请求暂存变更：\n仓库: {}\n目标: {}\n\n是否继续？").format(repo_dir, target)
                ):
                    notes.append(tr("⏸ 已取消 Git 暂存: {}").format(repo_dir))
                    continue
                ok, _, err_msg = self._run_git_command(repo_dir, ['add', '--', target])
                if ok:
                    notes.append(tr("✅ Git 已暂存 ({}) 目标: {}").format(repo_dir, target))
                else:
                    notes.append(tr("❌ Git 执行失败: {}").format(err_msg))

            # ── GIT_COMMIT ──────────────────────────────────────────────────
            elif atype == 'GIT_COMMIT':
                raw = m.group(1)
                parts = raw.split('|', 1)
                if len(parts) < 2:
                    notes.append(tr("❌ Git 指令格式应为 [GIT_COMMIT: 路径|提交信息]"))
                    continue
                repo_raw, commit_msg = parts[0].strip(), parts[1].strip()
                if not commit_msg:
                    notes.append(tr("❌ Git 指令格式应为 [GIT_COMMIT: 路径|提交信息]"))
                    continue
                repo_dir, err = self._resolve_git_repo_dir(repo_raw)
                if err:
                    notes.append(tr("❌ Git 执行失败: {}").format(err))
                    continue
                if not self._confirm_danger_action(
                    tr("确认 Git 提交"),
                    tr("AI 请求提交变更：\n仓库: {}\n提交信息: {}\n\n是否继续？").format(repo_dir, commit_msg)
                ):
                    notes.append(tr("⏸ 已取消 Git 提交: {}").format(repo_dir))
                    continue
                ok, out, err_msg = self._run_git_command(repo_dir, ['commit', '-m', commit_msg])
                if ok:
                    notes.append(tr("✅ Git 提交成功 ({})\n{}").format(repo_dir, out or tr("⚠️ Git 无输出")))
                else:
                    notes.append(tr("❌ Git 执行失败: {}").format(err_msg))

            # ── GIT_SWITCH ──────────────────────────────────────────────────
            elif atype == 'GIT_SWITCH':
                raw = m.group(1)
                parts = raw.split('|', 1)
                if len(parts) < 2:
                    notes.append(tr("❌ Git 指令格式应为 [GIT_SWITCH: 路径|分支名]"))
                    continue
                repo_raw, branch = parts[0].strip(), parts[1].strip()
                if not branch:
                    notes.append(tr("❌ Git 指令格式应为 [GIT_SWITCH: 路径|分支名]"))
                    continue
                repo_dir, err = self._resolve_git_repo_dir(repo_raw)
                if err:
                    notes.append(tr("❌ Git 执行失败: {}").format(err))
                    continue
                if not self._confirm_danger_action(
                    tr("确认 Git 切换分支"),
                    tr("AI 请求切换分支：\n仓库: {}\n分支: {}\n\n是否继续？").format(repo_dir, branch)
                ):
                    notes.append(tr("⏸ 已取消 Git 切换分支: {}").format(repo_dir))
                    continue
                ok, _, err_msg = self._run_git_command(repo_dir, ['switch', branch])
                if ok:
                    notes.append(tr("✅ Git 已切换到分支 ({}) -> {}").format(repo_dir, branch))
                else:
                    ok2, _, err_msg2 = self._run_git_command(repo_dir, ['checkout', branch])
                    if ok2:
                        notes.append(tr("✅ Git 已切换到分支 ({}) -> {}").format(repo_dir, branch))
                    else:
                        notes.append(tr("❌ Git 执行失败: {}").format(err_msg2 or err_msg))

            # ── GIT_RESTORE ────────────────────────────────────────────────
            elif atype == 'GIT_RESTORE':
                raw = m.group(1)
                parts = raw.split('|', 1)
                if len(parts) < 2:
                    notes.append(tr("❌ Git 指令格式应为 [GIT_RESTORE: 路径|目标]"))
                    continue
                repo_raw, target = parts[0].strip(), (parts[1].strip() or '.')
                repo_dir, err = self._resolve_git_repo_dir(repo_raw)
                if not err:
                    _, err = _AiActionMethods._resolve_git_target(self, repo_dir, target)
                if err:
                    notes.append(tr("❌ Git 执行失败: {}").format(err))
                    continue
                if not self._confirm_danger_action(
                    tr("确认 Git 还原"),
                    tr("AI 请求还原工作区改动：\n仓库: {}\n目标: {}\n\n是否继续？").format(repo_dir, target)
                ):
                    notes.append(tr("⏸ 已取消 Git 还原: {}").format(repo_dir))
                    continue
                ok, _, err_msg = self._run_git_command(repo_dir, ['restore', '--', target])
                if ok:
                    notes.append(tr("✅ Git 已还原 ({}) 目标: {}").format(repo_dir, target))
                else:
                    notes.append(tr("❌ Git 执行失败: {}").format(err_msg))

            # ── GIT_RESET_SOFT ─────────────────────────────────────────────
            elif atype == 'GIT_RESET_SOFT':
                raw = m.group(1)
                parts = raw.split('|', 1)
                repo_raw = parts[0].strip() if parts else ''
                target = (parts[1].strip() if len(parts) > 1 else '') or 'HEAD~1'
                repo_dir, err = self._resolve_git_repo_dir(repo_raw)
                if err:
                    notes.append(tr("❌ Git 执行失败: {}").format(err))
                    continue
                if not self._confirm_danger_action(
                    tr("确认 Git 软重置"),
                    tr("AI 请求软重置 HEAD：\n仓库: {}\n目标: {}\n\n是否继续？").format(repo_dir, target)
                ):
                    notes.append(tr("⏸ 已取消 Git 软重置: {}").format(repo_dir))
                    continue
                ok, _, err_msg = self._run_git_command(repo_dir, ['reset', '--soft', target])
                if ok:
                    notes.append(tr("✅ Git 软重置成功 ({}) -> {}").format(repo_dir, target))
                else:
                    notes.append(tr("❌ Git 执行失败: {}").format(err_msg))

            # ── GIT_PULL ───────────────────────────────────────────────────
            elif atype == 'GIT_PULL':
                raw = m.group(1)
                parts = [p.strip() for p in raw.split('|')]
                if not parts or not parts[0]:
                    notes.append(tr("❌ Git 执行失败: {}").format(tr("空路径")))
                    continue
                repo_raw = parts[0]
                remote = parts[1] if len(parts) > 1 and parts[1] else ''
                branch = parts[2] if len(parts) > 2 and parts[2] else ''
                repo_dir, err = self._resolve_git_repo_dir(repo_raw)
                if err:
                    notes.append(tr("❌ Git 执行失败: {}").format(err))
                    continue
                if not self._confirm_danger_action(
                    tr("确认 Git 拉取"),
                    tr("AI 请求拉取远程更新：\n仓库: {}\n远程: {}\n分支: {}\n\n是否继续？").format(
                        repo_dir, remote or '(default)', branch or '(default)'
                    )
                ):
                    notes.append(tr("⏸ 已取消 Git 拉取: {}").format(repo_dir))
                    continue
                args = ['pull']
                if remote:
                    args.append(remote)
                if branch:
                    args.append(branch)
                ok, out, err_msg = self._run_git_command(repo_dir, args, timeout_sec=60)
                if ok:
                    notes.append(tr("✅ Git 拉取成功 ({})\n{}").format(repo_dir, out or tr("⚠️ Git 无输出")))
                else:
                    notes.append(tr("❌ Git 执行失败: {}").format(err_msg))

            # ── GIT_PUSH ───────────────────────────────────────────────────
            elif atype == 'GIT_PUSH':
                raw = m.group(1)
                parts = [p.strip() for p in raw.split('|')]
                if not parts or not parts[0]:
                    notes.append(tr("❌ Git 执行失败: {}").format(tr("空路径")))
                    continue
                repo_raw = parts[0]
                remote = parts[1] if len(parts) > 1 and parts[1] else ''
                branch = parts[2] if len(parts) > 2 and parts[2] else ''
                repo_dir, err = self._resolve_git_repo_dir(repo_raw)
                if err:
                    notes.append(tr("❌ Git 执行失败: {}").format(err))
                    continue
                if not self._confirm_danger_action(
                    tr("确认 Git 推送"),
                    tr("AI 请求推送本地提交：\n仓库: {}\n远程: {}\n分支: {}\n\n是否继续？").format(
                        repo_dir, remote or '(default)', branch or '(default)'
                    )
                ):
                    notes.append(tr("⏸ 已取消 Git 推送: {}").format(repo_dir))
                    continue
                args = ['push']
                if remote:
                    args.append(remote)
                if branch:
                    args.append(branch)
                ok, out, err_msg = self._run_git_command(repo_dir, args, timeout_sec=60)
                if ok:
                    notes.append(tr("✅ Git 推送成功 ({})\n{}").format(repo_dir, out or tr("⚠️ Git 无输出")))
                else:
                    notes.append(tr("❌ Git 执行失败: {}").format(err_msg))

        if notes:
            content = content + "\n\n" + "\n".join(notes)
        return content, feedable


class AiActionWorker(QThread, _AiActionMethods):
    completed = pyqtSignal(str, object, bool)
    ui_requested = pyqtSignal(object)

    def __init__(self, content, base_dir, parent=None):
        super().__init__(parent)
        self.content = content
        self.base_dir = base_dir
        self._cancel_event = threading.Event()

    def cancel(self):
        self._cancel_event.set()
        self.requestInterruption()

    def _get_action_base_dir(self):
        return self.base_dir

    def _request_ui(self, kind, *arguments):
        request = {'kind': kind, 'arguments': arguments, 'ready': threading.Event(), 'result': False}
        self.ui_requested.emit(request)
        while not request['ready'].wait(0.05):
            if self._cancel_event.is_set():
                return False
        return False if self._cancel_event.is_set() else request['result']

    def _confirm_danger_action(self, title, text):
        return self._request_ui('confirm', title, text)

    def run(self):
        try:
            display, feedable = self._apply_actions(self.content)
        except Exception as error:
            display, feedable = self.content + '\n' + str(error), []
        self.completed.emit(display, feedable, self._cancel_event.is_set())

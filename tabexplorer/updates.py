"""检查 GitHub 新版本。"""

from PyQt5.QtCore import pyqtSignal, QThread, QUrl

from .i18n import tr
from .constants import APP_VERSION
from .debuglog import debug_print


UPDATE_RELEASE_API = "https://api.github.com/repos/caojinyuan/TabEx/releases/latest"
UPDATE_SOURCE_URL = "https://raw.githubusercontent.com/caojinyuan/TabEx/{tag}/TabEx.py"
UPDATE_RELEASES_PAGE = "https://github.com/caojinyuan/TabEx/releases"
UPDATE_CHECK_INTERVAL_S = 24 * 60 * 60
UPDATE_CHECK_DELAY_MS = 30 * 1000


def _version_tuple(text):
    """'3.75' / 'v3.75' -> (3, 75)；无法识别返回 None。"""
    import re
    match = re.fullmatch(r'v?(\d+(?:\.\d+)*)', str(text or '').strip())
    return tuple(int(part) for part in match.group(1).split('.')) if match else None


def _fetch_latest_release(timeout=10):
    """查询 GitHub 最新发布，并读取该发布源码开头的 APP_VERSION（发布标签与软件版本号不同）。"""
    import json
    import re
    import urllib.request
    headers = {'User-Agent': f'TabEx/{APP_VERSION}', 'Accept': 'application/vnd.github+json'}
    request = urllib.request.Request(UPDATE_RELEASE_API, headers=headers)
    with urllib.request.urlopen(request, timeout=timeout) as response:
        release = json.loads(response.read(1024 * 1024).decode('utf-8'))
    tag = str(release.get('tag_name') or '')
    if not re.fullmatch(r'[A-Za-z0-9._-]{1,64}', tag):
        raise ValueError(tr("发布标签无效"))
    page = str(release.get('html_url') or '')
    if not page.startswith(UPDATE_RELEASES_PAGE + '/'):
        page = UPDATE_RELEASES_PAGE
    version = ''
    try:
        source = urllib.request.Request(UPDATE_SOURCE_URL.format(tag=tag),
                                        headers={'User-Agent': headers['User-Agent'], 'Range': 'bytes=0-4095'})
        with urllib.request.urlopen(source, timeout=timeout) as response:
            head = response.read(4096).decode('utf-8', 'replace')
        match = re.search(r'^APP_VERSION\s*=\s*["\']([^"\']+)["\']', head, re.M)
        if match and _version_tuple(match.group(1)):
            version = match.group(1)
    except Exception as error:
        debug_print(f"[Update] cannot read version of {tag}: {error}")
    return {'tag': tag, 'version': version, 'url': page}


def _update_error_text(error):
    import socket
    import urllib.error
    if isinstance(error, urllib.error.HTTPError):
        if error.code in (403, 429):
            return tr("GitHub 请求过于频繁，请稍后再试")
        return tr("GitHub 返回错误 {}").format(error.code)
    if isinstance(error, (socket.timeout, TimeoutError)):
        return tr("连接 GitHub 超时")
    if isinstance(error, urllib.error.URLError):
        return tr("无法连接 GitHub：{}").format(error.reason)
    return tr("无法读取发布信息：{}").format(error)


def _open_release_page(url):
    from PyQt5.QtGui import QDesktopServices
    target = url if str(url).startswith(UPDATE_RELEASES_PAGE + '/') else UPDATE_RELEASES_PAGE
    QDesktopServices.openUrl(QUrl(target))


class UpdateCheckWorker(QThread):
    completed = pyqtSignal(object)  # dict：成功含 tag/version/url，失败含 error；均含 manual

    def __init__(self, manual):
        super().__init__()
        self.manual = bool(manual)

    def run(self):
        try:
            result = _fetch_latest_release()
        except Exception as error:
            result = {'error': _update_error_text(error)}
        result['manual'] = self.manual
        self.completed.emit(result)

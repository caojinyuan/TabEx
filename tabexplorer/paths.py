"""程序数据目录与路径工具。"""

import os
import sys


def get_app_base_dir():
    """Return a stable base directory for persistent app data."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    # 源码运行时为项目根目录（本包的上一级），与拆分前的数据文件位置一致
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_app_data_path(*parts):
    override = os.environ.get('TABEX_DATA_DIR')
    base = os.path.abspath(override) if override else get_app_base_dir()
    return os.path.join(base, *parts)


def get_app_log_path(*parts):
    """Return a stable log path outside the bundled application resources."""
    return get_app_data_path('logs', *parts)


def get_app_config_path(file_name=''):
    """Resolve a configuration file, migrating a legacy root file without overwriting."""
    directory = get_app_data_path('config')
    os.makedirs(directory, exist_ok=True)
    if not file_name:
        return directory
    target = os.path.join(directory, file_name)
    legacy = get_app_data_path(file_name)
    if not os.path.exists(target) and os.path.isfile(legacy):
        try:
            os.rename(legacy, target)
        except (FileExistsError, FileNotFoundError):
            # Another process may have migrated the same file during startup.
            if not os.path.isfile(target):
                raise
    return target


def get_bundled_path(file_name):
    """返回随程序内置的文件路径（打包后位于 PyInstaller 解压目录，源码运行时位于项目根目录）。"""
    roots = [getattr(sys, '_MEIPASS', None), get_app_base_dir()]
    for root in roots:
        if root:
            candidate = os.path.join(root, file_name)
            if os.path.isfile(candidate):
                return candidate
    return None


def map_sync_path(leader_old, leader_new, follower_base):
    """分屏同步浏览：一侧从 leader_old 进入子目录或返回上级到 leader_new 时，返回另一侧应前往的路径。

    另一侧从 follower_base 出发做相同的相对移动；跳到无关目录、虚拟位置或越过盘符根时返回 None。"""
    from pathlib import PureWindowsPath
    paths = (leader_old, leader_new, follower_base)
    if not all(isinstance(path, str) and path and not path.startswith(('shell:', '::')) for path in paths):
        return None
    old_parts = PureWindowsPath(leader_old).parts
    new_parts = PureWindowsPath(leader_new).parts

    def folded(parts):
        return [part.casefold() for part in parts]

    if len(new_parts) > len(old_parts) and folded(new_parts[:len(old_parts)]) == folded(old_parts):
        return str(PureWindowsPath(follower_base, *new_parts[len(old_parts):]))
    if len(new_parts) < len(old_parts) and folded(old_parts[:len(new_parts)]) == folded(new_parts):
        target = PureWindowsPath(follower_base)
        for _level in range(len(old_parts) - len(new_parts)):
            if target.parent == target:
                return None
            target = target.parent
        return str(target)
    return None


# 常用中英文目录名对照表
_COMMON_PATH_PAIRS = [
    ("Users", "用户"),
    ("Documents", "文档"),
    ("Desktop", "桌面"),
    ("Downloads", "下载"),
    ("Pictures", "图片"),
    ("Music", "音乐"),
    ("Videos", "视频"),
    ("Favorites", "收藏夹"),
    ("OneDrive", "OneDrive"),
    ("AppData", "AppData"),
    ("Roaming", "Roaming"),
    ("Local", "Local"),
    ("Public", "Public"),
]


def is_slow_path(path):
    """Classify potentially blocking paths without probing the filesystem."""
    if not path:
        return False
    if path.startswith(('\\\\', '//')) or 'onedrive' in path.lower():
        return True
    if os.name == 'nt':
        import ctypes
        drive = os.path.splitdrive(path)[0]
        if drive:
            return ctypes.windll.kernel32.GetDriveTypeW(drive + '\\') == 4
    return False


def translate_common_path(path):
    """
    尝试将路径中的中英文常用目录名互相转换，递归尝试所有组合，返回第一个存在的路径。
    若找不到存在的路径，则返回原始 path。
    """
    if not path or is_slow_path(path) or os.path.exists(path):
        return path
    # 只处理本地绝对路径
    norm_path = os.path.normpath(path)
    # 分割为各级目录
    parts = norm_path.split(os.sep)
    # 记录所有可替换的目录位置
    replace_indices = []
    replace_options = []
    for i, part in enumerate(parts):
        opts = [part]
        for en, zh in _COMMON_PATH_PAIRS:
            if part == en:
                opts.append(zh)
            elif part == zh:
                opts.append(en)
        if len(opts) > 1:
            replace_indices.append(i)
            replace_options.append(opts)
    # 如果没有可替换的，直接返回原始
    if not replace_indices:
        return path
    # 递归生成所有组合
    from itertools import product
    for combo in product(*replace_options):
        new_parts = parts[:]
        for idx, val in zip(replace_indices, combo):
            new_parts[idx] = val
        candidate = os.sep.join(new_parts)
        if os.path.exists(candidate):
            return candidate
    # 没找到存在的路径，返回原始
    return path

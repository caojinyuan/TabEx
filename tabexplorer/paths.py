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


def translate_common_path(path):
    """
    尝试将路径中的中英文常用目录名互相转换，递归尝试所有组合，返回第一个存在的路径。
    若找不到存在的路径，则返回原始 path。
    """
    if not path or os.path.exists(path):
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

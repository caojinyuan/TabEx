"""测试辅助：按名称读取 tabexplorer 包内对象；patch_all 在所有导入了该名称的模块中同时替换。"""

import importlib
from pathlib import Path
from unittest.mock import DEFAULT, MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
# 由底层到上层排列：同名对象以最先出现的模块（即定义处）为准
MODULE_NAMES = (
    'paths', 'i18n', 'constants', 'debuglog', 'diagnostics', 'system', 'hotkeys', 'title_shortcuts', 'widgets',
    'updates', 'workers', 'search', 'fileops', 'pathbar', 'bookmarks', 'shellview', 'explorer_tab', 'tabbar',
    'chat', 'settings', 'mainwindow', 'app',
)
MODULE_FILES = [ROOT / 'tabexplorer' / f'{name}.py' for name in MODULE_NAMES]
MODULES = [importlib.import_module(f'tabexplorer.{name}') for name in MODULE_NAMES]


class _App:
    """只读汇总视图；替换对象必须用 patch_all，直接赋值只会改到其中一个模块。"""

    def __getattr__(self, name):
        for module in MODULES:
            if name in vars(module):
                return vars(module)[name]
        raise AttributeError(name)

    def __setattr__(self, name, value):
        raise AttributeError(f'use patch_all({name!r}) instead of assigning on the aggregate view')


app = _App()


class patch_all:
    """与 patch.object 用法相同，但同时替换所有绑定了该名称的模块，保证各模块看到同一个替身。"""

    def __init__(self, name, new=DEFAULT, **kwargs):
        targets = [module for module in MODULES if name in vars(module)]
        if not targets:
            raise AttributeError(name)
        self.replacement = MagicMock(**kwargs) if new is DEFAULT else new
        self._patchers = [patch.object(module, name, self.replacement) for module in targets]

    def start(self):
        for patcher in self._patchers:
            patcher.start()
        return self.replacement

    def stop(self):
        for patcher in reversed(self._patchers):
            patcher.stop()

    def __enter__(self):
        return self.start()

    def __exit__(self, *exc_info):
        self.stop()
        return False

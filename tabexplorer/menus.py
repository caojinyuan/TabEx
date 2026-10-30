"""临时菜单的生命周期辅助：避免带父对象的 QMenu 在长时间运行中累积 GDI/USER 句柄。"""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMenu


def exec_menu(menu, pos, action=None):
    """同步弹出临时菜单并在返回后释放；每个残留菜单约占 2 个 GDI + 2 个 USER 对象。"""
    try:
        return menu.exec_(pos, action) if action is not None else menu.exec_(pos)
    finally:
        menu.deleteLater()


def clear_menu(menu):
    """重建菜单前调用：clear() 只删除 action，不会删除 addMenu() 创建的子菜单对象。"""
    for child in menu.findChildren(QMenu, options=Qt.FindDirectChildrenOnly):
        child.deleteLater()
    menu.clear()

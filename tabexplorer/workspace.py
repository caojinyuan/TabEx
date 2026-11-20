"""命名工作区管理窗口。"""

import time

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QDialogButtonBox, QHeaderView, QInputDialog, QLabel, QMessageBox, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout,
)

from .i18n import tr


class WorkspaceManagerDialog(QDialog):
    def __init__(self, owner, parent=None):
        super().__init__(parent)
        self.owner = owner
        self.setWindowTitle(tr("工作区管理"))
        self.resize(760, 500)

        layout = QVBoxLayout(self)
        self.empty_label = QLabel(tr("尚未保存工作区"), self)
        layout.addWidget(self.empty_label)

        self.tree = QTreeWidget(self)
        self.tree.setHeaderLabels([tr("工作区 / 分组 / 标签"), tr("保存时间")])
        self.tree.setRootIsDecorated(True)
        self.tree.setUniformRowHeights(True)
        self.tree.setAlternatingRowColors(True)
        self.tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tree.setColumnWidth(1, 160)
        self.tree.itemDoubleClicked.connect(self._on_double_click)
        self.tree.itemSelectionChanged.connect(self._update_actions)
        layout.addWidget(self.tree)

        buttons = QDialogButtonBox(QDialogButtonBox.Close, parent=self)
        buttons.button(QDialogButtonBox.Close).setText(tr("关闭"))
        self.save_button = buttons.addButton(tr("保存当前工作区"), QDialogButtonBox.ActionRole)
        self.open_button = buttons.addButton(tr("打开工作区"), QDialogButtonBox.ActionRole)
        self.switch_button = buttons.addButton(tr("切换工作区"), QDialogButtonBox.AcceptRole)
        self.rename_button = buttons.addButton(tr("重命名"), QDialogButtonBox.ActionRole)
        self.ungroup_button = buttons.addButton(tr("取消分组"), QDialogButtonBox.ActionRole)
        self.delete_button = buttons.addButton(tr("删除"), QDialogButtonBox.ActionRole)
        self.save_button.clicked.connect(self.save_current)
        self.open_button.clicked.connect(self.open_selected)
        self.switch_button.clicked.connect(self.switch_selected)
        self.rename_button.clicked.connect(self.rename_selected)
        self.ungroup_button.clicked.connect(self.ungroup_selected)
        self.delete_button.clicked.connect(self.delete_selected)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.populate()

    @staticmethod
    def _item_data(item):
        value = item.data(0, Qt.UserRole) if item is not None else None
        return value if isinstance(value, dict) else {}

    @staticmethod
    def _group_color(entry):
        if entry.get("is_pinned"):
            return ""
        return str(entry.get("bookmark_group_color", "") or "").strip()

    @staticmethod
    def _format_saved_at(state):
        try:
            return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(float(state["saved_at"])))
        except (KeyError, TypeError, ValueError, OverflowError, OSError):
            return tr("未知")

    def _append_tab_item(self, parent, pane_index, tab_index, entry):
        pane_label = tr("主视图") if pane_index == 0 else tr("右侧分屏")
        path = str(entry.get("path", "") or "")
        item = QTreeWidgetItem([f"{pane_label} · {path}", ""])
        item.setData(0, Qt.UserRole, {
            "type": "tab", "workspace": self._item_data(parent)["workspace"],
            "pane_index": pane_index, "tab_index": tab_index,
        })
        item.setToolTip(0, path)
        parent.addChild(item)

    def _append_workspace_contents(self, item, state):
        workspace_name = self._item_data(item)["workspace"]
        groups = state.get("groups", [])
        if not isinstance(groups, list):
            return
        for pane_index, pane in enumerate(groups[:2]):
            if not isinstance(pane, dict) or not isinstance(pane.get("tabs"), list):
                continue
            tabs = pane["tabs"]
            index = 0
            while index < len(tabs):
                entry = tabs[index]
                if not isinstance(entry, dict):
                    index += 1
                    continue
                color = self._group_color(entry)
                if not color:
                    self._append_tab_item(item, pane_index, index, entry)
                    index += 1
                    continue

                end = index
                while end + 1 < len(tabs):
                    next_entry = tabs[end + 1]
                    if not isinstance(next_entry, dict) or self._group_color(next_entry) != color:
                        break
                    end += 1
                group_name = str(entry.get("tab_group_name", "") or "").strip() or tr("未命名分组")
                pane_label = tr("主视图") if pane_index == 0 else tr("右侧分屏")
                group_item = QTreeWidgetItem([
                    f"{pane_label} · {group_name} ({end - index + 1})", ""])
                group_item.setData(0, Qt.UserRole, {
                    "type": "group", "workspace": workspace_name,
                    "pane_index": pane_index, "start": index, "end": end, "color": color,
                })
                item.addChild(group_item)
                for tab_index in range(index, end + 1):
                    grouped_entry = tabs[tab_index]
                    if isinstance(grouped_entry, dict):
                        self._append_tab_item(group_item, pane_index, tab_index, grouped_entry)
                index = end + 1

    def populate(self, selected_workspace=None):
        self.tree.clear()
        workspaces = self.owner.config.get("named_workspaces", {})
        for name, state in sorted(workspaces.items(), key=lambda entry: entry[0].casefold()):
            if not isinstance(state, dict):
                continue
            item = QTreeWidgetItem([name, self._format_saved_at(state)])
            item.setData(0, Qt.UserRole, {"type": "workspace", "workspace": name})
            self.tree.addTopLevelItem(item)
            self._append_workspace_contents(item, state)
            item.setExpanded(True)
            for child_index in range(item.childCount()):
                item.child(child_index).setExpanded(True)

        has_workspaces = self.tree.topLevelItemCount() > 0
        self.empty_label.setVisible(not has_workspaces)
        target = None
        if selected_workspace is not None:
            for index in range(self.tree.topLevelItemCount()):
                item = self.tree.topLevelItem(index)
                if item.text(0) == selected_workspace:
                    target = item
                    break
        if target is None and has_workspaces:
            target = self.tree.topLevelItem(0)
        self.tree.setCurrentItem(target)
        self._update_actions()

    def _selected_data(self):
        return self._item_data(self.tree.currentItem())

    def _update_actions(self):
        item_type = self._selected_data().get("type")
        self.open_button.setEnabled(item_type == "workspace")
        self.switch_button.setEnabled(item_type == "workspace")
        self.rename_button.setEnabled(item_type in ("workspace", "group"))
        self.ungroup_button.setEnabled(item_type == "group")
        self.delete_button.setEnabled(item_type in ("workspace", "group", "tab"))

    def _on_double_click(self, item, _column):
        if self._item_data(item).get("type") == "workspace":
            self.open_selected()

    def save_current(self):
        name = self.owner.save_named_workspace()
        self.populate(selected_workspace=name)

    def open_selected(self):
        data = self._selected_data()
        if data.get("type") != "workspace":
            return
        name = data["workspace"]
        self.accept()
        self.owner.open_named_workspace(name)

    def switch_selected(self):
        data = self._selected_data()
        if data.get("type") != "workspace":
            return
        name = data["workspace"]
        self.accept()
        self.owner.switch_named_workspace(name)

    def rename_selected(self):
        data = self._selected_data()
        workspaces = self.owner.config.get("named_workspaces", {})
        name = data.get("workspace")
        if data.get("type") == "workspace":
            new_name, accepted = QInputDialog.getText(self, tr("重命名工作区"), tr("名称"), text=name)
            new_name = new_name.strip()
            if not accepted or not new_name or new_name == name:
                return
            if new_name in workspaces:
                QMessageBox.warning(self, tr("重命名工作区"), tr("名称已存在"))
                return
            workspaces[new_name] = workspaces.pop(name)
            workspaces[new_name]["saved_at"] = time.time()
            self.owner.save_config(immediate=True)
            self.populate(selected_workspace=new_name)
            return

        if data.get("type") != "group":
            return
        pane = workspaces[name]["groups"][data["pane_index"]]
        tabs = pane["tabs"]
        current_name = str(tabs[data["start"]].get("tab_group_name", "") or "")
        new_name, accepted = QInputDialog.getText(
            self, tr("重命名分组"), tr("分组名称"), text=current_name)
        if not accepted:
            return
        for entry in tabs[data["start"]:data["end"] + 1]:
            entry["tab_group_name"] = new_name.strip()
        self._save_workspace_edit(name)

    def ungroup_selected(self):
        data = self._selected_data()
        if data.get("type") != "group":
            return
        state = self.owner.config["named_workspaces"][data["workspace"]]
        tabs = state["groups"][data["pane_index"]]["tabs"]
        for entry in tabs[data["start"]:data["end"] + 1]:
            entry["bookmark_group_color"] = ""
            entry["tab_group_name"] = ""
            entry["tab_group_collapsed"] = False
            entry["tab_group_separator_after"] = False
            entry["tab_group_separator_color"] = ""
            entry["tab_group_separator_name"] = ""
        self._save_workspace_edit(data["workspace"])

    def delete_selected(self):
        data = self._selected_data()
        item_type = data.get("type")
        name = data.get("workspace")
        if item_type == "workspace":
            self.owner.delete_named_workspace(name)
            self.populate()
            return
        if item_type not in ("group", "tab"):
            return

        item = self.tree.currentItem()
        prompt = (tr("将删除该分组中的全部标签：{}").format(item.text(0))
                  if item_type == "group" else tr("删除标签：{}").format(item.text(0)))
        if QMessageBox.question(
                self, tr("确认删除"), prompt, QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No) != QMessageBox.Yes:
            return

        state = self.owner.config["named_workspaces"][name]
        pane_index = data["pane_index"]
        pane = state["groups"][pane_index]
        tabs = pane["tabs"]
        start = data["start"] if item_type == "group" else data["tab_index"]
        end = data["end"] if item_type == "group" else data["tab_index"]
        del tabs[start:end + 1]
        active_index = int(pane.get("active_index", 0) or 0)
        if active_index > end:
            active_index -= end - start + 1
        elif active_index >= start:
            active_index = min(start, max(0, len(tabs) - 1))
        pane["active_index"] = active_index if tabs else 0
        if pane_index == 1 and not tabs:
            state["groups"].pop(pane_index)
        self._save_workspace_edit(name)

    def _save_workspace_edit(self, name):
        state = self.owner.config["named_workspaces"][name]
        state["saved_at"] = time.time()
        self.owner.save_config(immediate=True)
        self.populate(selected_workspace=name)

import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QDialog, QMainWindow, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QMessageBox, QVBoxLayout, QDialogButtonBox, QTableWidget,
    QHBoxLayout, QPushButton
)
from PyQt5 import uic
from PyQt5.QtGui import QColor
from PyQt5.QtCore import QTimer, QDateTime

DATA_FILE = "table_data.json"


class EditTableDialog(QDialog):
    def __init__(self, parent, table_data, headers):
        super().__init__(parent)
        self.setWindowTitle("Edit Table")
        self.resize(800, 500)

        self.table = QTableWidget()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(table_data))

        for row_index, row in enumerate(table_data):
            for col_index, value in enumerate(row):
                self.table.setItem(row_index, col_index, QTableWidgetItem(value))

        self.add_row_button = QPushButton("Add Row")
        self.remove_row_button = QPushButton("Remove Selected Row")
        self.move_up_button = QPushButton("Move Up")
        self.move_down_button = QPushButton("Move Down")

        self.add_row_button.clicked.connect(self.add_row)
        self.remove_row_button.clicked.connect(self.remove_row)
        self.move_up_button.clicked.connect(self.move_row_up)
        self.move_down_button.clicked.connect(self.move_row_down)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_row_button)
        button_layout.addWidget(self.remove_row_button)
        button_layout.addWidget(self.move_up_button)
        button_layout.addWidget(self.move_down_button)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addLayout(button_layout)
        layout.addWidget(self.buttons)
        self.setLayout(layout)

    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def remove_row(self):
        selected = self.table.currentRow()
        if selected >= 0:
            self.table.removeRow(selected)
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to remove.")

    def move_row_up(self):
        current = self.table.currentRow()
        if current > 0:
            self.swap_rows(current, current - 1)
            self.table.selectRow(current - 1)

    def move_row_down(self):
        current = self.table.currentRow()
        if current < self.table.rowCount() - 1:
            self.swap_rows(current, current + 1)
            self.table.selectRow(current + 1)

    def swap_rows(self, row1, row2):
        for col in range(self.table.columnCount()):
            item1 = self.table.item(row1, col)
            item2 = self.table.item(row2, col)
            text1 = item1.text() if item1 else ""
            text2 = item2.text() if item2 else ""
            self.table.setItem(row1, col, QTableWidgetItem(text2))
            self.table.setItem(row2, col, QTableWidgetItem(text1))

    def get_updated_data(self):
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        return [[self.table.item(i, j).text() if self.table.item(i, j) else "" for j in range(cols)] for i in range(rows)]


class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("WindowsSecurityUpdateMainWindow.ui", self)

        self.reload_done = False
        self.Reload.clicked.connect(self.reload)
        self.UpdateAll.clicked.connect(self.on_update_all_clicked)
        self.UpdateSelected.clicked.connect(self.on_update_selected_clicked)
        self.ClearSelections.clicked.connect(self.Table.clearSelection)
        self.Quit.clicked.connect(self.close)
        self.Test.clicked.connect(self.second_window)
        if hasattr(self, 'actionEdit_Table'):
            self.actionEdit_Table.triggered.connect(self.open_edit_table_dialog)

        self.set_button_disabled_appearance(self.UpdateAll, True)
        self.set_button_disabled_appearance(self.UpdateSelected, True)

        self.Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.Table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.Table.setSelectionMode(QAbstractItemView.MultiSelection)

        self.load_table_data()
        self.update_row_headers()

    def load_table_data(self):
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r") as f:
                saved = json.load(f)
                headers = saved["headers"]
                data = saved["data"]
                self.Table.setColumnCount(len(headers))
                self.Table.setRowCount(len(data))
                self.Table.setHorizontalHeaderLabels(headers)
                for row, row_data in enumerate(data):
                    for col, val in enumerate(row_data):
                        self.Table.setItem(row, col, QTableWidgetItem(val))

    def save_table_data(self):
        headers = [self.Table.horizontalHeaderItem(i).text() for i in range(self.Table.columnCount())]
        data = []
        for row in range(self.Table.rowCount()):
            row_data = []
            for col in range(self.Table.columnCount()):
                item = self.Table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        with open(DATA_FILE, "w") as f:
            json.dump({"headers": headers, "data": data}, f, indent=2)

    def set_button_disabled_appearance(self, button, disabled=True):
        button.setDisabled(disabled)
        button.setStyleSheet("""
            QPushButton {
                color: gray;
                background-color: lightgray;
                border: 1px solid #aaa;
            }
        """ if disabled else "")

    def update_row_headers(self):
        for row in range(self.Table.rowCount()):
            self.Table.setVerticalHeaderItem(row, QTableWidgetItem(str(row)))

    def show_warning(self):
        QMessageBox.warning(self, "Reload Required", "Please press the Reload button before updating.")

    def get_machine_list_from_table(self):
        return [self.Table.item(row, 0).text() for row in range(self.Table.rowCount()) if self.Table.item(row, 0)]

    def on_update_all_clicked(self):
        if not self.reload_done:
            self.show_warning()
            return
        self.apply_updates(self.get_machine_list_from_table())

    def on_update_selected_clicked(self):
        if not self.reload_done:
            self.show_warning()
            return
        selected_rows = {index.row() for index in self.Table.selectedIndexes()}
        selected_machines = [self.Table.item(i, 0).text() for i in selected_rows if self.Table.item(i, 0)]
        self.apply_updates(selected_machines)

    def apply_updates(self, selected_machines):
        if not selected_machines:
            print("No machines selected")
            return
        print(f"Applying updates to: {', '.join(selected_machines)}")

    def reload(self):
        for row in range(self.Table.rowCount()):
            self.Table.setItem(row, 4, QTableWidgetItem("Pending"))
            self.Table.setItem(row, 5, QTableWidgetItem("Pending"))
            self.Table.setItem(row, 6, QTableWidgetItem(""))
        self.reload_done = True
        self.set_button_disabled_appearance(self.UpdateAll, False)
        self.set_button_disabled_appearance(self.UpdateSelected, False)

    def second_window(self):
        self.Timing = QDialog(self)
        uic.loadUi('TimingUpdateWindow.ui', self.Timing)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()
        self.Timing.setWindowTitle("Update Timing")
        self.Timing.setModal(True)
        self.Timing.show()

    def update_time(self):
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        self.Timing.TimeAndDateLabel.setText(current_time)

    def open_edit_table_dialog(self):
        headers = [self.Table.horizontalHeaderItem(i).text() for i in range(self.Table.columnCount())]
        table_data = [
            [self.Table.item(row, col).text() if self.Table.item(row, col) else "" for col in range(self.Table.columnCount())]
            for row in range(self.Table.rowCount())
        ]

        dialog = EditTableDialog(self, table_data, headers)
        if dialog.exec_() == QDialog.Accepted:
            reply = QMessageBox.question(
                self, "Confirm Changes",
                "Are you sure you want to keep these changes?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                updated_data = dialog.get_updated_data()
                self.Table.setRowCount(len(updated_data))
                self.Table.setColumnCount(len(headers))
                self.Table.setHorizontalHeaderLabels(headers)
                for row_index, row_data in enumerate(updated_data):
                    for col_index, cell in enumerate(row_data):
                        self.Table.setItem(row_index, col_index, QTableWidgetItem(cell))
                self.update_row_headers()
                self.save_table_data()
                self.reload_done = False
                self.set_button_disabled_appearance(self.UpdateAll, True)
                self.set_button_disabled_appearance(self.UpdateSelected, True)


def main():
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

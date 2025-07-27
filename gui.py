import sys
import xml.etree.ElementTree as ET
from PyQt5.QtWidgets import (
    QApplication, QDialog, QMainWindow, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QMessageBox
)
from PyQt5 import uic
from PyQt5.QtGui import QColor
from PyQt5.QtCore import QTimer, QDateTime

class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui_file_path = 'WindowsSecurityUpdateMainWindow.ui'
        uic.loadUi(self.ui_file_path, self)

        self.machine_list = self.extract_machines_from_ui_file(self.ui_file_path)
        self.reload_done = False  # Flag to lock update buttons before reload

        # Connect buttons
        self.Reload.clicked.connect(self.reload)
        self.UpdateAll.clicked.connect(self.on_update_all_clicked)
        self.UpdateSelected.clicked.connect(self.on_update_selected_clicked)
        self.ClearSelections.clicked.connect(self.Table.clearSelection)
        self.Quit.clicked.connect(self.close)
        self.Test.clicked.connect(self.ScheduleUpdateWindow)

        # Make buttons look disabled initially
        self.set_button_disabled_appearance(self.UpdateAll, True)
        self.set_button_disabled_appearance(self.UpdateSelected, True)

        # Table setup
        self.Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.Table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.Table.setSelectionMode(QAbstractItemView.MultiSelection)

    def extract_machines_from_ui_file(self, file_path):
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()
            machines = []

            for row in root.iter("row"):
                prop = row.find('property[@name="text"]/string')
                if prop is not None and prop.text:
                    machines.append(prop.text.strip())

            return machines
        except Exception as e:
            print(f"Failed to read machine names from .ui file: {e}")
            return []

    def set_button_disabled_appearance(self, button, disabled=True):
        if disabled:
            button.setStyleSheet("""
                QPushButton {
                    color: gray;
                    background-color: lightgray;
                    border: 1px solid #aaa;
                }
            """)
        else:
            button.setStyleSheet("")

    def show_warning(self):
        QMessageBox.warning(self, "Reload Required", "Please press the Reload button before updating.")

    def on_update_all_clicked(self):
        if not self.reload_done:
            self.show_warning()
            return
        self.apply_updates(self.machine_list)

    def on_update_selected_clicked(self):
        if not self.reload_done:
            self.show_warning()
            return
        self.get_selected_items()

    def ScheduleUpdateWindow(self):
        self.Timing = QDialog(self)
        uic.loadUi('ScheduleUpdate.ui', self.Timing)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()

        self.Timing.setWindowTitle("Schedule Update")
        self.Timing.setModal(True)
        self.Timing.show()

    def update_time(self):
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        self.Timing.TimeAndDateLabel.setText(current_time)

    def reload(self):
        # Make sure the table has as many rows as machines
        self.Table.setRowCount(len(self.machine_list))

        # Enable update functionality and reset button styles
        self.reload_done = True
        self.set_button_disabled_appearance(self.UpdateAll, False)
        self.set_button_disabled_appearance(self.UpdateSelected, False)


    def apply_updates(self, selected_machines):
        if not selected_machines:
            print("No machines selected")
            return
        print(f"Applying updates to: {', '.join(selected_machines)}")
        # TODO: Add update logic here

    def get_selected_items(self):
        selected_rows = {index.row() for index in self.Table.selectedIndexes()}
        selected_machines = []
        for row_index in selected_rows:
            item = self.Table.item(row_index, 0)  # Column 0 = machine name
            if item:
                selected_machines.append(item.text())
        self.apply_updates(selected_machines)

def main():
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

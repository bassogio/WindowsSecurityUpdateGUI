import sys
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
        uic.loadUi('WindowsSecurityUpdateMainWindow.ui', self)

        self.machine_dict = {
            "71AW01": {"index": 0, "os": 2016},
            "71AW02": {"index": 1, "os": 2016},
            "71AW03": {"index": 2, "os": 2016},
            "71AW04": {"index": 3, "os": 2016},
            "71AW05": {"index": 4, "os": 2016},
            "71WS01": {"index": 5, "os": 10_1607},
            "71WS02": {"index": 6, "os": 10_1607},
            "71WS03": {"index": 7, "os": 10_1607},
            "71WS04": {"index": 8, "os": 10_1607},
            "71WS05": {"index": 9, "os": 10_1607},
            "71WS06": {"index": 10, "os": 10_1607},
            "TCLI01": {"index": 11, "os": 10_1607},
            "TCLI02": {"index": 12, "os": 10_1607},
            "TCLI03": {"index": 13, "os": 10_1607},
            "TCLI04": {"index": 14, "os": 10_1607},
            "TCLI05": {"index": 15, "os": 10_1089},
            "TCLI06": {"index": 16, "os": 10_1089},
            "71ES02": {"index": 17, "os": 10_1607},
            "71ES03": {"index": 18, "os": 10_1607},
            "VIRT03": {"index": 19, "os": 2019},
            "HIST1A": {"index": 20, "os": 2016},
            "NAS_SRV": {"index": 21, "os": 2019},
            "EPOSRV": {"index": 22, "os": 2016},
            "WSUS": {"index": 23, "os": 2016},
            "NETMON": {"index": 24, "os": 2016},
            "PI_MAG": {"index": 25, "os": 2016},
            "PIMAG_IN": {"index": 26, "os": 2016},
            "SYSADV": {"index": 27, "os": 2016},
            "AUTOSAVE": {"index": 28, "os": 2019},
            "PRIMDC": {"index": 29, "os": 2016},
            "SECNDC": {"index": 30, "os": 2016}
        }

        self.index_to_machine = {v["index"]: k for k, v in self.machine_dict.items()}
        self.reload_done = False  # Flag to lock update buttons before reload

        # Connect buttons
        self.Reload.clicked.connect(self.reload)
        self.UpdateAll.clicked.connect(self.on_update_all_clicked)
        self.UpdateSelected.clicked.connect(self.on_update_selected_clicked)
        self.ClearSelections.clicked.connect(self.Table.clearSelection)
        self.Quit.clicked.connect(self.close)
        self.Test.clicked.connect(self.second_window)

        # Make buttons look disabled initially
        self.set_button_disabled_appearance(self.UpdateAll, True)
        self.set_button_disabled_appearance(self.UpdateSelected, True)

        # Table setup
        self.Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.Table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.Table.setSelectionMode(QAbstractItemView.MultiSelection)

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
        self.apply_updates(list(self.machine_dict.keys()))

    def on_update_selected_clicked(self):
        if not self.reload_done:
            self.show_warning()
            return
        self.get_selected_items()

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

    def reload(self):
        self.Table.setItem(0, 0, QTableWidgetItem("test"))
        self.Table.setItem(0, 1, QTableWidgetItem("teeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeest"))

        self.success = QTableWidgetItem("Success")
        self.success.setBackground(QColor(0, 255, 0))
        self.fail = QTableWidgetItem("Failed")
        self.fail.setBackground(QColor(255, 0, 0))
        self.warning = QTableWidgetItem("Warning")
        self.warning.setBackground(QColor(255, 255, 0))

        self.Table.setItem(0, 2, self.success)
        self.Table.setItem(0, 3, self.fail)
        self.Table.setItem(0, 4, self.warning)

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
            machine = self.index_to_machine.get(row_index)
            if machine:
                selected_machines.append(machine)
        self.apply_updates(selected_machines)

def main():
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

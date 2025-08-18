import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QDialog, QMainWindow, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QMessageBox, QVBoxLayout, QDialogButtonBox, QTableWidget,
    QHBoxLayout, QPushButton,  QDockWidget, QTextEdit, QProgressBar
)
from PyQt5 import uic
from PyQt5.QtGui import QColor
from PyQt5.QtCore import QTimer, QDateTime, Qt

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

        self.load_done = False
        self.Load.clicked.connect(self.load) 
        self.UpdateAll.clicked.connect(self.on_update_all_clicked)
        self.UpdateSelected.clicked.connect(self.on_update_selected_clicked)
        self.ClearSelections.clicked.connect(self.Table.clearSelection)
        self.ScheduleTheUpdates.clicked.connect(self.ScheduleUpdatesWindow)
        self.Quit.clicked.connect(self.close)

        if hasattr(self, 'actionEdit_Table'):
            self.actionEdit_Table.triggered.connect(self.open_edit_table_dialog)

        if hasattr(self, 'actionChange_Path_to_Installation_Files'):
            self.actionChange_Path_to_Installation_Files.triggered.connect(self.open_edit_path_dialog)

        self.UpdateAll.setEnabled(True)
        self.UpdateSelected.setEnabled(True)
        self.ScheduleTheUpdates.setEnabled(True)
        self.set_button_visual_state(self.UpdateAll, inactive=True)
        self.set_button_visual_state(self.UpdateSelected, inactive=True)
        self.set_button_visual_state(self.ScheduleTheUpdates, inactive=True)

        self.Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.Table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.Table.setSelectionMode(QAbstractItemView.MultiSelection)

        self.load_table_data()
        self.update_row_headers()
        self.load_schedule()

        if hasattr(self, 'actionHelp_Center'):
            self.actionHelp_Center.triggered.connect(self.toggle_help_panel)
        self.init_help_panel()

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

    def init_help_panel(self):
        """Create a dockable Help Center panel but don't show it yet."""
        self.help_dock = QDockWidget("Help Center", self)
        self.help_dock.setObjectName("HelpDock")
        self.help_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)

        self.help_text = QTextEdit()
        self.help_text.setReadOnly(True)
        self.help_text.setHtml("""
            <h2>Help Center</h2>
            <p>Welcome to the Help Center. Here’s what you can do:</p>
            <ul>
                <li><b>Load</b>: Scans all listed machines and updates their info.</li>
                <li><b>Update All</b>: Installs updates on every machine in the list.</li>
                <li><b>Update Selected</b>: Only applies updates to selected rows.</li>
                <li><b>Clear Selections</b>: Deselects all rows.</li>
                <li><b>Edit Table</b>: Add, remove, or reorder machines and save changes.</li>
            </ul>
        """)
        self.help_dock.setWidget(self.help_text)
        self.help_dock.setVisible(False)
        self.addDockWidget(Qt.RightDockWidgetArea, self.help_dock)

    def toggle_help_panel(self):
        """Toggle visibility of the Help Center dock."""
        self.help_dock.setVisible(not self.help_dock.isVisible())

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

    def set_button_visual_state(self, button, inactive=False):
        if inactive:
            button.setStyleSheet("""
                QPushButton {
                    color: gray;
                    background-color: lightgray;
                    border: 1px solid #aaa;
                }
            """)
        else:
            button.setStyleSheet("")

    def update_row_headers(self):
        for row in range(self.Table.rowCount()):
            self.Table.setVerticalHeaderItem(row, QTableWidgetItem(str(row)))

    def show_warning(self):
        QMessageBox.warning(self, "Load Required", "Please press the Reload button before updating.")

    def get_machine_list_from_table(self):
        return [self.Table.item(row, 0).text() for row in range(self.Table.rowCount()) if self.Table.item(row, 0)]

    def on_update_all_clicked(self):
        if not self.load_done:
            self.show_warning()
            return
        self.apply_updates(self.get_machine_list_from_table())

    def on_update_selected_clicked(self):
        if not self.load_done:
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
    
    def get_simulated_os(self, hostname):
        # Simulated OS detection logic
        import random
        simulated_os_list = [
            "Windows 10 Enterprise", 
            "Windows Server 2016", 
            "Windows Server 2019", 
            "Windows 10 Pro", 
            "Windows 11", 
            "Unknown"
        ]
        return random.choice(simulated_os_list)

    def get_remote_os(self, hostname):
        """
        Returns the OS of a remote machine using native PowerShell.
        Requires:
            - Remote PowerShell enabled on the target
            - Your user has access rights
        """
        import subprocess
        try:
            ps_command = f"Invoke-Command -ComputerName {hostname} -ScriptBlock {{ (Get-CimInstance Win32_OperatingSystem).Caption }}"
            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
                capture_output=True, text=True, timeout=10
            )
            output = result.stdout.strip()
            error = result.stderr.strip()

            if result.returncode != 0 or error:
                return f"Error: {error}" if error else "Unknown error"
            return output if output else "Unavailable"
        except Exception as e:
            return f"Exception: {str(e)}"
        
    def get_simulated_disk_space(self, hostname):
        import random
        total = random.randint(100, 500)
        free = round(random.uniform(0.1, 0.9) * total, 1)
        return free, total  # return as tuple of numbers

    def set_disk_space_progress(self, row, col, free_gb, total_gb):
        if total_gb == 0:
            percent = 0
        else:
            percent = int((free_gb / total_gb) * 100)

        # Set the text value into the cell (so it gets saved)
        text = f"{free_gb:.1f} GB / {total_gb:.1f} GB"
        self.Table.setItem(row, col, QTableWidgetItem(text))

        # Show progress visually on top of that cell
        progress = QProgressBar()
        progress.setRange(0, 100)
        progress.setValue(percent)
        progress.setAlignment(Qt.AlignCenter)
        progress.setTextVisible(True)
        progress.setFormat(text)

        progress.setStyleSheet(f"""
            QProgressBar {{
                border: 1px solid gray;
                border-radius: 3px;
                text-align: center;
            }}
            QProgressBar::chunk {{
                background-color: {"#4caf50" if percent > 50 else "#ff9800" if percent > 20 else "#f44336"};
                width: 1px;
            }}
        """)

        self.Table.setCellWidget(row, col, progress)

    def get_disk_space(self, hostname):
        """
        Returns (free_gb, total_gb) of the primary disk on a remote machine using PowerShell.
        Requires:
            - Remote PowerShell enabled
            - Firewall rules for PowerShell remoting
            - Access rights to the remote host
        """
        import subprocess
        import re
        try:
            ps_command = (
                f"Invoke-Command -ComputerName {hostname} -ScriptBlock {{ "
                f"$d = Get-CimInstance Win32_LogicalDisk -Filter \"DeviceID='C:'\"; "
                f"'Free=' + $d.FreeSpace + ';Total=' + $d.Size "
                f"}}"
            )

            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
                capture_output=True, text=True, timeout=10
            )
            output = result.stdout.strip()
            error = result.stderr.strip()

            if result.returncode != 0 or error:
                raise RuntimeError(error or "Unknown PowerShell error")

            # Parse the output (expected: Free=1234567890;Total=9876543210)
            match = re.search(r'Free=(\d+);Total=(\d+)', output)
            if not match:
                raise ValueError("Unexpected format from remote command")

            free_bytes = int(match.group(1))
            total_bytes = int(match.group(2))

            free_gb = round(free_bytes / (1024**3), 1)
            total_gb = round(total_bytes / (1024**3), 1)
            return free_gb, total_gb

        except Exception as e:
            print(f"[Disk Error] {hostname}: {e}")
            return 0.0, 0.0  # fallback

    def load(self):
        header_index = {self.Table.horizontalHeaderItem(i).text(): i for i in range(self.Table.columnCount())}

        # Load OS paths
        if os.path.exists("os_paths.json"):
            with open("os_paths.json", "r") as f:
                os_paths = json.load(f)
        else:
            os_paths = {}

        for row in range(self.Table.rowCount()):
            machine_item = self.Table.item(row, 0)
            if machine_item:
                hostname = machine_item.text()

                os_name = self.get_simulated_os(hostname)
                # os_name = self.get_remote_os(hostname)
                self.Table.setItem(row, header_index["OS"], QTableWidgetItem(os_name))

                free_gb, total_gb = self.get_simulated_disk_space(hostname)
                # free_gb, total_gb = self.get_disk_space(hostname)
                self.set_disk_space_progress(row, header_index["Disk Space"], free_gb, total_gb)

            
                os_path = os_paths.get(os_name, "")
                cumulative_kb = ""
                servicing_stack_kb = ""
                if os_path and os.path.isdir(os_path):
                    for fname in os.listdir(os_path):
                        # Cumulative KB
                        if "Cumulative" in fname:
                            kb_match = self.extract_kb_from_filename(fname)
                            if kb_match:
                                cumulative_kb = kb_match
                        # Servicing Stack KB
                        if "Servicing Stack" in fname:
                            kb_match = self.extract_kb_from_filename(fname)
                            if kb_match:
                                servicing_stack_kb = kb_match
                self.Table.setItem(row, header_index.get("Cumulative", -1), QTableWidgetItem(cumulative_kb))
                self.Table.setItem(row, header_index.get("Servicing Stack", -1), QTableWidgetItem(servicing_stack_kb))

        header = self.Table.horizontalHeader()
        for i in range(self.Table.columnCount()):
            col_name = self.Table.horizontalHeaderItem(i).text()
            if col_name in ["OS", "Machine", "Servicing Stack", "Cumulative", "Install Status", "Disk Space"]:
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
            else:
                header.setSectionResizeMode(i, QHeaderView.Stretch)

        self.load_done = True
        self.set_button_visual_state(self.UpdateAll, inactive=False)
        self.set_button_visual_state(self.UpdateSelected, inactive=False)
        self.set_button_visual_state(self.ScheduleTheUpdates, inactive=False)

    def extract_kb_from_filename(self, fname):
        import re
        match = re.search(r'KB(\d+)', fname)
        if match:
            return f"KB{match.group(1)}"
        return ""

    def ScheduleUpdatesWindow(self):
        if not self.load_done:
            self.show_warning()
            return
        
        self.Timing = QDialog(self)
        uic.loadUi('ScheduleUpdates.ui', self.Timing)

        self.Timing.OkCancle.clicked.connect(self.schedule_updates_task)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()
        self.Timing.setWindowTitle("Schedule Updates")
        self.Timing.setModal(True)
        self.Timing.show()

    def schedule_updates_task(self):
        date = self.Timing.SetDate.selectedDate()
        time = self.Timing.SetTime.time()
        scheduled_datetime = QDateTime(date, time)

        if scheduled_datetime <= QDateTime.currentDateTime():
            QMessageBox.warning(self, "Invalid Time", "Please select a future time.")
            return 

        self.scheduled_time = scheduled_datetime
        self.repeat_mode = {
            "weekly": self.Timing.RepeatWeek.isChecked(),
            "monthly": self.Timing.RepeatMonth.isChecked(),
            "yearly": self.Timing.RepeatYear.isChecked()
        }
        # Get email address from the input field
        self.notification_email = self.Timing.EmailInput.text()

        self.save_schedule()

        self.monitor_timer = QTimer(self)
        self.monitor_timer.timeout.connect(self.check_if_time_reached)
        self.monitor_timer.start(10_000)

        self.Timing.accept()

    def check_if_time_reached(self):
        now = QDateTime.currentDateTime()
        if now >= self.scheduled_time:
            self.monitor_timer.stop()
            self.run_scheduled_updates()

    def run_scheduled_updates(self):        
        machine_list = self.get_machine_list_from_table()
        self.apply_updates(machine_list)
        QMessageBox.information(self, "Updates Started", f"Scheduled update started at {QDateTime.currentDateTime().toString()}")
        
    def update_time(self):
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        self.Timing.TimeAndDateLabel.setText(current_time)

    def save_schedule(self):
        data = {
            "timestamp": self.scheduled_time.toSecsSinceEpoch(),
            "date": self.scheduled_time.toString("yyyy-MM-dd"),
            "time": self.scheduled_time.toString("HH:mm:ss"),
            "repeat": self.repeat_mode,
            "email": getattr(self, "notification_email", "")
        }
        with open("schedule.json", "w") as f:
            json.dump(data, f, indent=2)

    def load_schedule(self):
        if os.path.exists("schedule.json"):
            with open("schedule.json", "r") as f:
                data = json.load(f)
                ts = data.get("timestamp")
                if ts:
                    self.scheduled_time = QDateTime.fromSecsSinceEpoch(ts)
                    self.repeat_mode = data.get("repeat", {
                        "weekly": False,
                        "monthly": False,
                        "yearly": False
                    })
                    self.monitor_timer = QTimer(self)
                    self.monitor_timer.timeout.connect(self.check_if_time_reached)
                    self.monitor_timer.start(10_000)

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
                self.load_done = False
                self.set_button_visual_state(self.UpdateAll, inactive=True)
                self.set_button_visual_state(self.UpdateSelected, inactive=True)

    def open_edit_path_dialog(self):
        from PyQt5.QtWidgets import QWidget, QLabel, QLineEdit, QFileDialog, QGridLayout

        class PathDialog(QDialog):
            def __init__(self, parent, os_list, os_paths):
                super().__init__(parent)
                self.setWindowTitle("Edit Installation Paths by OS")
                self.resize(600, 300)
                self.os_list = os_list
                self.os_paths = os_paths.copy()
                self.edits = {}

                layout = QGridLayout()
                for i, os_name in enumerate(os_list):
                    label = QLabel(os_name)
                    edit = QLineEdit(os_paths.get(os_name, ""))
                    browse_btn = QPushButton("Browse")
                    browse_btn.clicked.connect(lambda _, e=edit: self.browse_folder(e))
                    layout.addWidget(label, i, 0)
                    layout.addWidget(edit, i, 1)
                    layout.addWidget(browse_btn, i, 2)
                    self.edits[os_name] = edit

                self.update_btn = QPushButton("Update")
                self.close_btn = QPushButton("Close")
                layout.addWidget(self.update_btn, len(os_list), 1)
                layout.addWidget(self.close_btn, len(os_list), 2)
                self.setLayout(layout)

                self.update_btn.clicked.connect(self.update_paths)
                self.close_btn.clicked.connect(self.confirm_close)
                self.changes_saved = False

            def browse_folder(self, edit):
                folder = QFileDialog.getExistingDirectory(self, "Select Folder")
                if folder:
                    edit.setText(folder)

            def update_paths(self):
                for os_name, edit in self.edits.items():
                    self.os_paths[os_name] = edit.text()
                with open("os_paths.json", "w") as f:
                    json.dump(self.os_paths, f, indent=2)
                self.changes_saved = True
                QMessageBox.information(self, "Saved", "Paths updated successfully.")

            def confirm_close(self):
                if not self.changes_saved:
                    reply = QMessageBox.question(
                        self,
                        "Exit Without Saving",
                        "Are you sure you want to exit without updating the changes?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No
                    )
                    if reply == QMessageBox.Yes:
                        self.reject()
                        return
                    else:
                        return
                self.accept()

        # Gather OSs from table (excluding "Unknown")
        os_set = set()
        header_index = {self.Table.horizontalHeaderItem(i).text(): i for i in range(self.Table.columnCount())}
        os_col = header_index.get("OS")
        if os_col is None:
            QMessageBox.warning(self, "Error", "No OS column found.")
            return
        for row in range(self.Table.rowCount()):
            item = self.Table.item(row, os_col)
            if item:
                os_name = item.text()
                if os_name and os_name != "Unknown":
                    os_set.add(os_name)
        os_list = sorted(os_set)

        # Load current paths
        if os.path.exists("os_paths.json"):
            with open("os_paths.json", "r") as f:
                os_paths = json.load(f)
        else:
            os_paths = {}

        dialog = PathDialog(self, os_list, os_paths)
        dialog.exec_()

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self,
            "Save Changes",
            "Do you want to save the updated table before exiting?",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes
        )

        if reply == QMessageBox.Yes:
            self.save_table_data()
            event.accept()
        elif reply == QMessageBox.No:
            event.accept()
        else:
            event.ignore()

def main():
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
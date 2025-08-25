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
from PyQt5.QtGui import QTextDocument, QPageSize, QPageLayout
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtCore import QMarginsF, QRectF
import subprocess

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

        self.Table.setTextElideMode(Qt.ElideNone)
        self.Table.setWordWrap(True)
        self.Table.resizeRowsToContents()

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
        
        self.scheduled_dates = getattr(self, "scheduled_dates", [])

        self.scheduled_update_executed = False 

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
        QMessageBox.warning(self, "Load Required", "Please press the Load button before updating.")

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

        # Determine mode by comparing selection vs. all rows
        all_machines = set(self.get_machine_list_from_table())
        mode = "Update All" if set(selected_machines) == all_machines else "Update Selected"

        started_at = QDateTime.currentDateTime()
        print(f"[{mode}] Applying updates to: {', '.join(selected_machines)}")

        finished_at = QDateTime.currentDateTime()

        # Always generate a report for EVERY machine (as requested)
        self.generate_pdf_report(mode=mode, started_at=started_at, finished_at=finished_at)

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

        # Work on a temp list while the dialog is open
        self.temp_scheduled_dates = [dict(d) for d in self.scheduled_dates]

        # Ensure the button box doesn't auto-close the dialog
        try:
            self.Timing.OkCancle.accepted.disconnect()
            self.Timing.OkCancle.rejected.disconnect()
        except TypeError:
            pass  # nothing was connected by the .ui

        ok_btn = self.Timing.OkCancle.button(QDialogButtonBox.Ok)
        cancel_btn = self.Timing.OkCancle.button(QDialogButtonBox.Cancel)
        ok_btn.clicked.connect(self.schedule_updates_task)
        cancel_btn.clicked.connect(self.cancel_schedule_window)

        # Buttons that edit dates operate on the temp list
        self.Timing.AddDateButton.clicked.connect(self.add_date)
        self.Timing.RemoveDateButton.clicked.connect(self.remove_selected_date)

        self.populate_dates_listwidget()  # show temp list

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()
        self.Timing.setWindowTitle("Schedule Updates")
        self.Timing.setModal(True)
        self.Timing.show()

    def cancel_schedule_window(self):
        # Discard temp changes and just close
        self.Timing.reject()

    def add_date(self):
        date_str = self.Timing.SetDate.selectedDate().toString("yyyy-MM-dd")
        time_str = self.Timing.SetTime.time().toString("HH:mm:ss")
        entry = {"date": date_str, "time": time_str}
        if entry not in self.temp_scheduled_dates:
            self.temp_scheduled_dates.append(entry)
            self.Timing.DatesListWidget.addItem(f"{date_str} {time_str}")

    def remove_selected_date(self):
        selected_items = self.Timing.DatesListWidget.selectedItems()
        for item in selected_items:
            date_part, time_part = item.text().split(" ")
            self.temp_scheduled_dates = [
                d for d in self.temp_scheduled_dates
                if not (d.get("date") == date_part and d.get("time") == time_part)
            ]
            self.Timing.DatesListWidget.takeItem(self.Timing.DatesListWidget.row(item))

    def schedule_updates_task(self):
        date = self.Timing.SetDate.selectedDate()
        time = self.Timing.SetTime.time()
        scheduled_datetime = QDateTime(date, time)
        now = QDateTime.currentDateTime()

        # 1) Validate the currently selected date/time
        if scheduled_datetime <= now:
            sel_str = scheduled_datetime.toString("yyyy-MM-dd HH:mm:ss")
            now_str = now.toString("yyyy-MM-dd HH:mm:ss")
            QMessageBox.warning(
                self.Timing,
                "Invalid Time",
                f"The selected time ({sel_str}) is not in the future.\n"
                f"Current time is {now_str}.\nPlease pick a future date/time."
            )
            self.Timing.SetTime.setFocus()
            return  # keep dialog open

        # 2) Validate all entries in the list and name the bad ones
        invalid_entries = []
        for d in self.temp_scheduled_dates:
            dt = QDateTime.fromString(f"{d.get('date','')} {d.get('time','')}", "yyyy-MM-dd HH:mm:ss")
            if dt.isValid() and dt <= now:
                invalid_entries.append(dt.toString("yyyy-MM-dd HH:mm:ss"))

        if invalid_entries:
            bad = "\n • ".join(invalid_entries)
            QMessageBox.warning(
                self.Timing,
                "Invalid Time(s)",
                "These scheduled times are not in the future:\n"
                f" • {bad}\n\nPlease adjust or remove them."
            )
            self.Timing.DatesListWidget.setFocus()
            return  # keep dialog open

        # Commit only after all validations pass
        self.scheduled_dates = [dict(d) for d in self.temp_scheduled_dates]
        self.scheduled_time = scheduled_datetime
        self.Server = self.Timing.EmailServer.text()
        self.FromEmail = self.Timing.FromEmail.text()
        self.ToEmail = self.Timing.ToEmail.text()

        self.save_schedule()
        self.load_schedule()   # reload schedule (UI + future dates detection)

        # REPLACE the manual timer setup with:
        self.start_monitor()

        self.Timing.accept()  # close only after success

    def check_if_time_reached(self):
        now = QDateTime.currentDateTime()
        due_found = False
        remaining = []

        for entry in list(self.scheduled_dates):
            if not isinstance(entry, dict):
                continue
            dt = QDateTime.fromString(f"{entry.get('date','')} {entry.get('time','')}", "yyyy-MM-dd HH:mm:ss")
            if not dt.isValid():
                continue

            if dt <= now:
                due_found = True
            else:
                remaining.append(entry)

        if due_found:
            # commit list changes first to avoid re-entrancy
            self.scheduled_dates = remaining
            self.save_schedule()
            # run exactly once per tick even if multiple entries were due
            self.run_scheduled_updates()

        # Stop monitor if nothing left; otherwise keep ticking
        if not self.scheduled_dates:
            try:
                self.monitor_timer.stop()
            except AttributeError:
                pass

    def run_scheduled_updates(self):        
        machine_list = self.get_machine_list_from_table()
        self.apply_updates(machine_list)
        QMessageBox.information(self, "Updates Started", f"Scheduled update started at {QDateTime.currentDateTime().toString()}")

    def update_time(self):
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        self.Timing.TimeAndDateLabel.setText(current_time)

    def save_schedule(self):
        data = {
            "dates": self.scheduled_dates,  # List of dicts: {"date": "...", "time": "..."}
            "Email server": getattr(self, "Server", ""),
            "From email": getattr(self, "FromEmail", ""),
            "To email": getattr(self, "ToEmail", "")
        }
        with open("schedule.json", "w") as f:
            json.dump(data, f, indent=2)
    
    def start_monitor(self):
        # Ensure a single timer instance
        if hasattr(self, "monitor_timer"):
            try:
                self.monitor_timer.stop()
            except Exception:
                pass
        else:
            self.monitor_timer = QTimer(self)
            self.monitor_timer.timeout.connect(self.check_if_time_reached)
        self.monitor_timer.start(1000)  # 1s granularity

    def load_schedule(self):
        self.scheduled_dates = []

        if os.path.exists("schedule.json"):
            with open("schedule.json", "r") as f:
                data = json.load(f)
            self.scheduled_dates = data.get("dates", [])
            self.Server = data.get("Email server", "")
            self.FromEmail = data.get("From email", "")
            self.ToEmail = data.get("To email", "")

        # If any date/time is still in the future, start the monitor
        now = QDateTime.currentDateTime()
        has_future = False
        for d in self.scheduled_dates:
            dt = QDateTime.fromString(f"{d.get('date','')} {d.get('time','')}", "yyyy-MM-dd HH:mm:ss")
            if dt.isValid() and dt > now:
                has_future = True
                break

        if has_future:
            self.start_monitor()

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

    def populate_dates_listwidget(self):
        self.Timing.DatesListWidget.clear()
        # Use temp list if present (dialog open), else the committed list
        src = getattr(self, "temp_scheduled_dates", self.scheduled_dates)
        for entry in src:
            if isinstance(entry, dict):
                self.Timing.DatesListWidget.addItem(f'{entry.get("date","")} {entry.get("time","")}')

    def _header_index_map(self):
        return {
            self.Table.horizontalHeaderItem(i).text().strip(): i
            for i in range(self.Table.columnCount())
        }

    def _get_cell_text(self, row, col):
        """Prefer item text; if there's a QProgressBar (Disk Space), use its visible text."""
        item = self.Table.item(row, col)
        if item and item.text():
            return item.text()
        w = self.Table.cellWidget(row, col)
        if isinstance(w, QProgressBar):
            return w.text() or f"{w.value()}%"
        return ""

    def row_to_dict(self, row_index):
        cols = self._header_index_map()
        want = ["Machine", "OS", "Disk Space", "Servicing Stack", "Cumulative", "KB", "Install Status", "Notes"]
        out = {}
        for name in want:
            ci = cols.get(name, -1)
            out[name] = self._get_cell_text(row_index, ci) if ci >= 0 else ""
        return out

    def collect_all_rows(self):
        return [self.row_to_dict(r) for r in range(self.Table.rowCount())]

    def build_report_html(self, meta, rows):
        # derive success/failure counts
        succ_keys = {"success", "succeeded", "installed"}
        fail_keys = {"fail", "failed", "error"}
        success = sum(1 for r in rows if any(k in r["Install Status"].lower() for k in succ_keys))
        failed  = sum(1 for r in rows if any(k in r["Install Status"].lower() for k in fail_keys))

        started   = meta.get("started_at_str", "")
        finished  = meta.get("finished_at_str", "")
        elapsed   = meta.get("elapsed_str", "")
        mode      = meta.get("mode", "")

        # Simple, print-friendly CSS (A4, no external fonts)
        css = """
        <style>
          * { box-sizing: border-box; }
          body { font-family: Arial, Helvetica, sans-serif; font-size: 11pt; margin: 0; }
          h1 { font-size: 16pt; margin: 6pt 0 2pt 0; }
          .meta { margin: 0 0 10pt 0; }
          .meta div { margin: 2pt 0; }
          table { width: 100%; border-collapse: collapse; page-break-inside: auto; }
          thead { display: table-header-group; }
          tfoot { display: table-row-group; }
          tr { page-break-inside: avoid; page-break-after: auto; }
          th, td { border: 1px solid #999; padding: 6pt 5pt; vertical-align: top; }
          th { background: #f2f2f2; text-align: left; }
          .status-success { color: #0a7d2a; font-weight: bold; }
          .status-failed  { color: #b00020; font-weight: bold; }
          .small { color: #555; font-size: 9pt; }
        </style>
        """

        # Build table rows with conditional class on Install Status
        def status_class(s):
            s_low = s.lower()
            if any(k in s_low for k in succ_keys):
                return "status-success"
            if any(k in s_low for k in fail_keys):
                return "status-failed"
            return ""

        row_html = []
        for r in rows:
            row_html.append(f"""
              <tr>
                <td>{r['Machine']}</td>
                <td>{r['OS']}</td>
                <td>{r['Disk Space']}</td>
                <td>{r['Servicing Stack']}</td>
                <td>{r['Cumulative']}</td>
                <td>{r['KB']}</td>
                <td class="{status_class(r['Install Status'])}">{r['Install Status']}</td>
                <td>{r['Notes']}</td>
              </tr>
            """)

        html = f"""
        <html>
          <head>{css}</head>
          <body>
            <h1>Windows Update Report</h1>
            <div class="meta">
              <div><b>Mode:</b> {mode}</div>
              <div><b>Started:</b> {started}</div>
              <div><b>Finished:</b> {finished}</div>
              <div><b>Elapsed:</b> {elapsed}</div>
              <div><b>Summary:</b> Success: {success} &nbsp;&nbsp; Failed: {failed} &nbsp;&nbsp; Total: {len(rows)}</div>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Machine</th>
                  <th>OS</th>
                  <th>Disk Space</th>
                  <th>Servicing Stack</th>
                  <th>Cumulative</th>
                  <th>KB</th>
                  <th>Install Status</th>
                  <th>Notes</th>
                </tr>
              </thead>
              <tbody>
                {''.join(row_html)}
              </tbody>
            </table>
            <div class="small" style="margin-top:8pt;">Generated on {finished}</div>
          </body>
        </html>
        """
        return html

    def save_pdf(self, html, out_path):
        # Prepare printer
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(out_path)
        layout = QPageLayout(QPageSize(QPageSize.A4), QPageLayout.Portrait, QMarginsF(15, 15, 18, 18))  # mm
        printer.setPageLayout(layout)

        # Lay out the document at page size (points)
        doc = QTextDocument()
        doc.setHtml(html)
        doc.setPageSize(printer.pageRect(QPrinter.Point).size())

        # Manual paginate so we can draw "Page X of Y"
        from PyQt5.QtGui import QPainter
        painter = QPainter(printer)

        # Metrics in points
        page_rect_pt = QRectF(printer.pageRect(QPrinter.Point))
        page_h = page_rect_pt.height()
        page_w = page_rect_pt.width()
        footer_h = 20.0  # points reserved for footer

        # Count pages after layout
        page_count = doc.documentLayout().pageCount()
        if page_count < 1:
            page_count = 1

        for i in range(page_count):
            if i > 0:
                printer.newPage()

            # Draw document slice for this page with a clip (leave room for footer)
            painter.save()
            clip = QRectF(page_rect_pt.left(), page_rect_pt.top(), page_w, page_h - footer_h)
            painter.setClipRect(clip)
            painter.translate(0, -i * page_h)  # show the slice for page i
            doc.drawContents(painter, QRectF(page_rect_pt.left(), page_rect_pt.top(), page_w, page_h))
            painter.restore()

            # Footer: "Page X of Y"
            footer_rect = QRectF(page_rect_pt.left(), page_rect_pt.bottom() - footer_h, page_w, footer_h)
            painter.drawText(footer_rect, Qt.AlignCenter, f"Page {i+1} of {page_count}")

        painter.end()
        return True

    def generate_pdf_report(self, mode, started_at, finished_at):
        # Build meta
        started_str  = started_at.toString("yyyy-MM-dd HH:mm:ss")
        finished_str = finished_at.toString("yyyy-MM-dd HH:mm:ss")
        elapsed_secs = max(0, finished_at.toSecsSinceEpoch() - started_at.toSecsSinceEpoch())
        mm, ss = divmod(elapsed_secs, 60)
        hh, mm = divmod(mm, 60)
        elapsed_str = f"{hh:02d}:{mm:02d}:{ss:02d}"

        rows = self.collect_all_rows()  # REQUIREMENT: include every machine

        html = self.build_report_html(
            {
                "mode": mode,
                "started_at_str": started_str,
                "finished_at_str": finished_str,
                "elapsed_str": elapsed_str,
            },
            rows
        )

        # Ensure reports/ exists
        os.makedirs("reports", exist_ok=True)
        ts = QDateTime.currentDateTime().toString("yyyy-MM-dd_HH-mm-ss")
        out_path = os.path.join("reports", f"UpdateReport_{ts}.pdf")

        ok = self.save_pdf(html, out_path)
        if ok:
            QMessageBox.information(self, "Report saved", f"PDF report saved to:\n{out_path}")
            send_report_email(out_path)  # <--- Send email after saving PDF
        else:
            QMessageBox.warning(self, "Report error", "Failed to save the PDF report.")

def send_report_email(pdf_path):
    # Load email settings from schedule.json
    try:
        with open("schedule.json", "r") as f:
            config = json.load(f)
        smtp_server = config.get("Email server", "")
        from_addr = config.get("From email", "")
        to_addr = config.get("To email", "")
    except Exception as e:
        print(f"Failed to load email config: {e}")
        return

    subject = "Windows Security Update Report"
    body = "Please find the attached update report."

    if not (smtp_server and from_addr and to_addr):
        print("Email settings are incomplete in schedule.json.")
        return

    ps_command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-Command",
        f'Send-MailMessage -From "{from_addr}" -To "{to_addr}" -Subject "{subject}" -Body "{body}" -SmtpServer "{smtp_server}" -Attachments "{os.path.abspath(pdf_path)}"'
    ]
    result = subprocess.run(ps_command, capture_output=True, text=True)
    if result.returncode == 0:
        print("Email sent successfully.")
    else:
        print(f"Email send failed: {result.stderr}")

def main():
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
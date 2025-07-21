import sys
from PyQt5.QtWidgets import QApplication, QDialog , QTableWidgetItem, QHeaderView, QAbstractItemView
from PyQt5 import uic
from PyQt5.QtGui import QColor

class MyApp(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('WindowsSecurityUpdate.ui', self)

        self.machine_dict = {
                            "71AW01"  : {"index": 0, "os": 'I_dont_know'},
                            "71AW02"  : {"index": 1, "os": 2016},
                            "71AW03"  : {"index": 2, "os": 2016},
                            "71AW04"  : {"index": 3, "os": 2016},
                            "71AW05"  : {"index": 4, "os": 2016},
                            "71WS01"  : {"index": 5, "os": 10_1607},
                            "71WS02"  : {"index": 6, "os": 10_1607},
                            "71WS03"  : {"index": 7, "os": 10_1607},
                            "71WS04"  : {"index": 8, "os": 10_1607},
                            "71WS05"  : {"index": 9, "os": 10_1607},
                            "71WS06"  : {"index": 10, "os": 10_1607},
                            "TCLI01"  : {"index": 11, "os": 10_1607},
                            "TCLI02"  : {"index": 12, "os": 10_1607},
                            "TCLI03"  : {"index": 13, "os": 10_1607},
                            "TCLI04"  : {"index": 14, "os": 10_1607},
                            "TCLI05"  : {"index": 15, "os": 10_1089},
                            "TCLI06"  : {"index": 16, "os": 10_1089},
                            "71ES02"  : {"index": 17, "os": 10_1607},
                            "71ES03"  : {"index": 18, "os": 10_1607},
                            "VIRT03"  : {"index": 19, "os": 2019},
                            "HIST1A"  : {"index": 20, "os": 2016},
                            "NAS_SRV" : {"index": 21, "os": 2019},
                            "EPOSRV"  : {"index": 22, "os": 2016},
                            "WSUS"    : {"index": 23, "os": 2016},
                            "NETMON"  : {"index": 24, "os": 2016},
                            "PI_MAG"  : {"index": 25, "os": 2016},
                            "PIMAG_IN": {"index": 26, "os": 2016},
                            "SYSADV"  : {"index": 27, "os": 2016},
                            "AUTOSAVE": {"index": 28, "os": 2019},
                            "PRIMDC"  : {"index": 29, "os": 'I_dont_know'},
                            "SECNDC"  : {"index": 30, "os": 'I_dont_know'}
                            }
        
        #TODO: revert it manualy, but i am too lazy to do it now.
        self.index_to_machine = {v["index"]: k for k,v in self.machine_dict.items()} # Creates a new reversed dictionary.

        self.Reload.clicked.connect(self.reload) #TODO: Do the reload action
        self.UpdateAll.clicked.connect(self.apply_updates) #TODO: Do the update action - should be break to two parts - first servicing then comulative
        self.UpdateSelected.clicked.connect(self.get_selected_items) #TODO: Add update selected machines action - maybe in the general update I should update all or....
        self.ClearSelections.clicked.connect(self.Table.clearSelection)
        self.Quit.clicked.connect(self.close)

        # Set up the table widget
        header = self.Table.horizontalHeader() # Get the horizontal header of the table
        self.Table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) # Set the horizontal header to stretch
        self.Table.setSelectionBehavior(QAbstractItemView.SelectRows) # Set the selection behavior to select rows
        self.Table.setSelectionMode(QAbstractItemView.MultiSelection) # Set the selection mode to multiple selection

    def reload(self):
        # Set the table to have one row and one column
        self.Table.setItem(0,0, 
                          QTableWidgetItem("test")) 
        self.Table.setItem(0,1, 
                  QTableWidgetItem("teeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeest"))
        
        self.success = QTableWidgetItem("Success")
        self.success.setBackground(QColor(0, 255, 0))
        self.fail = QTableWidgetItem("Failed")
        self.fail.setBackground(QColor(255, 0, 0))
        self.warning = QTableWidgetItem("Warning")
        self.warning.setBackground(QColor(255, 255, 0))

        self.Table.setItem(0, 2, self.success)
        self.Table.setItem(0, 3, self.fail)
        self.Table.setItem(0, 4, self.warning)
        

    def apply_updates(self, selected_machines):
        if not selected_machines:
            print("No machines selected")
            return
        print(f"Applying updates to: {', '.join(selected_machines)}")
        # Here you would add the logic to apply updates to the selected machines 

    def get_selected_items(self):
        selected_rows = {index.row() for index in self.Table.selectedIndexes()}
        selected_machines = []
        for row_index in selected_rows:
                machine = self.index_to_machine.get(row_index)
                if machine:
                    selected_machines.append(machine)
        self.apply_updates(selected_machines)
        
app = QApplication(sys.argv)
window = MyApp()
window.show()
sys.exit(app.exec_())





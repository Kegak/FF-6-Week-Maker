from PyQt5 import QtWidgets, uic, sip
import sys
from FF_6_Week import *


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('FF_6_week.ui', self)
        self.show()
        self.month_reset = QtWidgets.QListWidgetItem("--")
        self.date_reset = QtWidgets.QListWidgetItem("0")
        self.button = self.findChild(QtWidgets.QPushButton, 'pushButton')  # Find the button
        self.button.clicked.connect(self.ButtonPressed)

        self.first = self.findChild(QtWidgets.QLineEdit, 'First_name_text')

        self.last = self.findChild(QtWidgets.QLineEdit, 'Last_name_list')

        self.month = self.findChild(QtWidgets.QListWidget, 'month_start_list')

        self.vaca_start_month = self.findChild(QtWidgets.QListWidget, 'month_vaca_start_list')

        self.vaca_start_month_2 = self.findChild(QtWidgets.QListWidget, 'month_vaca_start_list_2')

        self.vaca_start_month_3 = self.findChild(QtWidgets.QListWidget, 'month_vaca_start_list_3')

        self.vaca_end_month = self.findChild(QtWidgets.QListWidget, 'month_vaca_end_list')

        self.vaca_end_month_2 = self.findChild(QtWidgets.QListWidget, 'month_vaca_end_list_2')

        self.vaca_end_month_3 = self.findChild(QtWidgets.QListWidget, 'month_vaca_end_list_3')

        self.date = self.findChild(QtWidgets.QListWidget, 'date_start_list')

        self.vaca_date_start = self.findChild(QtWidgets.QListWidget, 'date_vaca_start_list')

        self.vaca_date_start_2 = self.findChild(QtWidgets.QListWidget, 'date_vaca_start_list_2')

        self.vaca_date_start_3 = self.findChild(QtWidgets.QListWidget, 'date_vaca_start_list_3')

        self.vaca_date_end = self.findChild(QtWidgets.QListWidget, 'date_vaca_end_list')

        self.vaca_date_end_2 = self.findChild(QtWidgets.QListWidget, 'date_vaca_end_list_2')

        self.vaca_date_end_3 = self.findChild(QtWidgets.QListWidget, 'date_vaca_end_list_3')

        self.height = self.findChild(QtWidgets.QLineEdit, 'height_text')

        self.which_fast = self.findChild(QtWidgets.QListWidget, 'listWidget')

        self.label = self.findChild(QtWidgets.QLabel, 'success_label')

    def ButtonPressed(self):

        height = int(self.height.text())

        first = self.first.text()

        last = self.last.text()

        which_fast = self.which_fast.currentRow() + 1

        month = str(self.month.currentItem().text())

        vaca_start_month = str(self.vaca_start_month.currentItem().text())

        if vaca_start_month == '--':
            vaca_start_month = 'Skadoosh'

        vaca_end_month = str(self.vaca_end_month.currentItem().text())

        if vaca_end_month == '--':
            vaca_end_month = 'Skadoosh'

        vaca_start_date = int(str(self.vaca_date_start.currentItem().text()))

        vaca_end_date = int(str(self.vaca_date_end.currentItem().text()))

        vaca_start_month_2 = str(self.vaca_start_month_2.currentItem().text())

        if vaca_start_month_2 == '--':
            vaca_start_month_2 = 'Skadoosh'

        vaca_end_month_2 = str(self.vaca_end_month_2.currentItem().text())

        if vaca_end_month_2 == '--':
            vaca_end_month_2 = 'Skadoosh'

        vaca_start_date_2 = int(str(self.vaca_date_start_2.currentItem().text()))

        vaca_end_date_2 = int(str(self.vaca_date_end_2.currentItem().text()))

        vaca_start_month_3 = str(self.vaca_start_month_3.currentItem().text())

        if vaca_start_month_3 == '--':
            vaca_start_month_3 = 'Skadoosh'

        vaca_end_month_3 = str(self.vaca_end_month.currentItem().text())

        if vaca_end_month_3 == '--':
            vaca_end_month_3 = 'Skadoosh'

        vaca_start_date_3 = int(str(self.vaca_date_start_3.currentItem().text()))

        vaca_end_date_3 = int(str(self.vaca_date_end_3.currentItem().text()))

        date = int(str(self.date.currentItem().text()))



        six_week(first, last, month, date, height, which_fast, vaca_start_month, vaca_start_date, vaca_end_month, vaca_end_date, vaca_start_month_2, vaca_start_date_2, vaca_end_month_2, vaca_end_date_2, vaca_start_month_3, vaca_start_date_3, vaca_end_month_3, vaca_end_date_3)


        self.vaca_start_month.setCurrentRow(0)
        self.vaca_end_month.setCurrentRow(0)
        self.vaca_date_start.setCurrentRow(0)
        self.vaca_date_end.setCurrentRow(0)
        self.vaca_start_month_2.setCurrentRow(0)
        self.vaca_end_month_2.setCurrentRow(0)
        self.vaca_date_start_2.setCurrentRow(0)
        self.vaca_date_end_2.setCurrentRow(0)
        self.vaca_start_month_3.setCurrentRow(0)
        self.vaca_end_month_3.setCurrentRow(0)
        self.vaca_date_start_3.setCurrentRow(0)
        self.vaca_date_end_3.setCurrentRow(0)
        self.which_fast.setCurrentRow(0)
        self.first.clear()
        self.last.clear()
        self.height.clear()
        self.label.setText(f'{first}-{last} 6 week plan.xlsx has been created.')

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    app.exec()
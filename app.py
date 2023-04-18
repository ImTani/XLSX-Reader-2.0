import sys
import webbrowser
import openpyxl
from main2 import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QApplication


class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.browseButton.clicked.connect(self.browseFiles)
        self.actionOpen_.triggered.connect(self.browseFiles)
        self.actionCheck_For_Updates.triggered.connect(self.openGithub)
        # Add your application logic here

    def browseFiles(self):
        browsedFilePath = QFileDialog.getOpenFileName(self.centralwidget, "Open File", filter="Excel Files (*xlsx)") # noqa E501
        self.filePath = browsedFilePath[0]
        if self.filePath:
            self.pathLineEdit.setText(self.filePath)
            self.ReadFile()
        else:
            self.pathLineEdit.setText("No file selected.")

    def openGithub(self):
        webbrowser.open_new_tab("https://github.com/ImTani/XLSX-Reader-2.0")

    def ReadFile(self):

        wb = openpyxl.load_workbook(self.filePath)
        sheet = wb.active

        num_rows = sheet.max_row # noqa F841
        num_columns = sheet.max_column # noqa F841

        first_row_values = [cell.value for cell in sheet[1]]

        marks_column_index = first_row_values.index("Marks") + 1

        mark_values_set = set()
        # Initialize a variable to store the previous row number
        prev_row_number = 0
        # Loop through all the rows and retrieve values from the "Marks" column
        for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2): # noqa E501
            mark_value = row[marks_column_index - 1]
            # Skip empty cells
            if mark_value is None:
                continue
            # Skip if the current row has the same row number as the previous row # noqa E501
            if i == prev_row_number + 1:
                continue
            # Add non-empty mark values to the set
            mark_values_set.add(mark_value)
            # Update the previous row number
            prev_row_number = i
        # Convert the set to a list
        mark_values_list = list(mark_values_set)
        mark_values_list = [str(i) for i in mark_values_list]

        self.comboBox.clear()
        self.comboBox.addItems(mark_values_list)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

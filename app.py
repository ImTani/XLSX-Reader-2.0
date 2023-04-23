import sys
import os
import webbrowser
import openpyxl
from main import Ui_MainWindow
from PyQt5 import QtWidgets
from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QApplication, QTableWidgetItem, QMessageBox  # noqa E501


class MyApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.updateRecentFilesMenu()
        self.sheetDropbox.currentIndexChanged.connect(self.loadSheet)
        self.browseButton.clicked.connect(self.browseFiles)
        self.checkButton.clicked.connect(self.getInfo)
        self.saveButton.clicked.connect(self.saveFile)
        self.actionOpen_.triggered.connect(self.browseFiles)
        self.actionCheck_For_Updates.triggered.connect(self.openGithub)
        self.actionSave_2.triggered.connect(self.saveFile)
        self.actionClear_Recent_Files.triggered.connect(self.clearRecentFilesMenu)  # noqa E501
        # Add your application logic here

    def browseFiles(self):
        browsedFilePath = QFileDialog.getOpenFileName(self.centralwidget, "Open File", filter="Excel Files (*.xlsx)") # noqa E501
        self.filePath = browsedFilePath[0]
        if self.filePath:
            self.tableWidget.setRowCount(0)
            self.getSheets(self.filePath)
        else:
            self.pathLineEdit.setText("No file selected.")

    def openGithub(self):
        webbrowser.open_new_tab("https://github.com/ImTani/XLSX-Reader-2.0")

    def ReadFile(self, filePath):

        self.filePath = filePath

        wb = openpyxl.load_workbook(filePath)
        try:
            sheet = wb[self.sheet]
        except KeyError:
            return
        num_rows = sheet.max_row # noqa F841
        num_columns = sheet.max_column # noqa F841

        first_row_values = [cell.value for cell in sheet[1]]

        try:
            marks_column_index = first_row_values.index("Marks") + 1
        except ValueError:
            error_dialog = QMessageBox()
            error_dialog.setIcon(QMessageBox.Critical)
            error_dialog.setWindowTitle("Error")
            error_dialog.setWindowIcon(self.icon)
            error_dialog.setText("The selected sheet does not contain the 'Marks' column.")  # noqa E501
            error_dialog.exec_()
            self.tableWidget.setRowCount(0)
            return

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

    def addRecentFile(self, file_path):
        settings = QSettings('TaniDev', 'XL_Reader')
        recent_files = settings.value('recent_files', [])
        if file_path in recent_files:
            recent_files.remove(file_path)
        recent_files.insert(0, file_path)
        if len(recent_files) > 5:
            recent_files.pop()
        settings.setValue('recent_files', recent_files)
        self.updateRecentFilesMenu()

    def updateRecentFilesMenu(self):
        settings = QSettings('TaniDev', 'XL_Reader')
        recent_files = settings.value('recent_files', [])
        self.menuRecent_Files.clear()
        for i, file_path in enumerate(recent_files):
            action = QtWidgets.QAction(f"{i+1}. {file_path}", self)
            action.triggered.connect(lambda _, fp=file_path: self.getSheets(fp))
            self.menuRecent_Files.addAction(action)

    def clearRecentFilesMenu(self):
        self.menuRecent_Files.clear()
        settings = QSettings('TaniDev', 'XL_Reader')
        settings.setValue('recent_files', [])

    def getSheets(self, filePath):
        self.filePath = filePath

        self.sheetDropbox.clear()
        self.addRecentFile(filePath)
        self.pathLineEdit.setText(self.filePath)

        workbook = openpyxl.load_workbook(filePath)
        # Get the sheet names
        sheet_names = workbook.sheetnames
        for i in sheet_names:
            self.sheetDropbox.addItem(i)

        self.sheet = sheet_names[0]

        self.ReadFile(filePath)

    def loadSheet(self):
        self.sheet = self.sheetDropbox.currentText()
        self.comboBox.clear()
        self.maxScoreValue.setText('')
        self.minScoreValue.setText('')
        self.maxScoreValue_2.setText('')
        self.minScoreValue_2.setText('')
        self.totalStudentValue_2.setText('')
        self.totalStudentValue.setText('')
        try:
            self.ReadFile(self.filePath)
        except openpyxl.utils.exceptions.InvalidFileException:
            return

    def getInfo(self):
        # Load the workbook and active sheet
        try:
            self.subCode = int(self.comboBox.currentText())
        except ValueError:
            return
        self.organiseData()

    def organiseData(self):
        try:
            wb = openpyxl.load_workbook(window.filePath)
        except openpyxl.utils.exceptions.InvalidFileException:
            error_dialog = QMessageBox()
            error_dialog.setIcon(QMessageBox.Warning)
            error_dialog.setWindowTitle("Error")
            error_dialog.setWindowIcon(self.icon)
            error_dialog.setText("No file is selected.")  # noqa E501
            error_dialog.exec_()
            self.tableWidget.setRowCount(0)
            self.sheetDropbox.clear()
            self.comboBox.clear()
            return
        sheet = wb[self.sheet]

        max_row = sheet.max_row

        # Define the selected subject code (example: 370)
        selected_subject_code = self.subCode

        # Create a list to store the filtered data
        self.filtered_data = []

        # Loop through each row in the worksheet
        for row in range(2, max_row+1):
            # Get the value of the cell in the second column of the current row
            cell_value = sheet.cell(row=row, column=2).value

            # Check if the cell value contains the selected subject code
            if str(selected_subject_code) in str(cell_value):
                # If there is a match, get the name, marks, and grade of the student
                name = sheet.cell(row=row, column=1).value

                # Get the marks from the cell immediately below the current cell
                marks = sheet.cell(row=row+1, column=2).value

                grade = sheet.cell(row=row+1, column=3).value

                # Add the filtered data to the list
                self.filtered_data.append([name, marks, grade])

        # Print the filtered data
        self.makeTable()

    def makeTable(self):
        dataLen = len(self.filtered_data)
        self.tableWidget.setRowCount(dataLen)
        self.tableWidget.setColumnCount(3)
        self.totalStudentValue.setText(str(dataLen))
        students = [i for i in self.filtered_data]
        name = []
        marks = []
        grades = []

        for i in students:
            name.append(i[0])
            marks.append(str(i[1]))
            grades.append(i[2])

        for i in range(dataLen):
            self.tableWidget.setItem(i, 0, QTableWidgetItem(name[i]))
            self.tableWidget.setItem(i, 1, QTableWidgetItem(marks[i]))
            self.tableWidget.setItem(i, 2, QTableWidgetItem(grades[i]))

        sum = 0
        for i in marks:
            try:
                sum += int(i)
            except ValueError:
                error_dialog = QMessageBox()
                error_dialog.setIcon(QMessageBox.Warning)
                error_dialog.setWindowTitle("Error")
                error_dialog.setWindowIcon(self.icon)
                error_dialog.setText("The given sheet is not a marksheet.")  # noqa E501
                error_dialog.exec_()
                self.tableWidget.setRowCount(0)
                return
            avg = round(sum/len(marks), 2)

        self.maxScoreValue.setText(str(max(marks)))
        self.minScoreValue.setText(str(min(marks)))
        self.maxScoreValue_2.setText(name[marks.index(max(marks))])
        self.minScoreValue_2.setText(name[marks.index(min(marks))])
        self.totalStudentValue_2.setText(str(avg))

    def saveFile(self):
        # Create a new workbook and sheet
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # copy headers
        for column in range(self.tableWidget.columnCount()):
            header_item = (self.tableWidget.horizontalHeaderItem(column))
            if header_item is not None:
                sheet.cell(row=1, column=column+1, value=header_item.text())

        # copy data
        for row in range(self.tableWidget.rowCount()):
            for column in range(self.tableWidget.columnCount()):
                item = (self.tableWidget.item(row, column))
                if item is not None:
                    sheet.cell(row=row+2, column=column+1, value=item.text())

        savePath, _ = QFileDialog.getSaveFileName(self.centralwidget, "Save File", filter="Excel Files (*.xlsx)")  # noqa

        if savePath:
            dirPath, fileNameExt = os.path.split(savePath)
            fileName, fileExt = os.path.splitext(fileNameExt)

            if fileExt != ".xlsx":
                fileExt = ".xlsx"

            savePath = os.path.join(dirPath, fileName + fileExt)

            workbook.save(savePath)
        else:
            return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())

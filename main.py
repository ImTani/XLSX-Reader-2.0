# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(650, 470)
        MainWindow.setMinimumSize(QtCore.QSize(650, 470))
        MainWindow.setMaximumSize(QtCore.QSize(650, 470))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("windowIcon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pathLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.pathLineEdit.setEnabled(True)
        self.pathLineEdit.setGeometry(QtCore.QRect(120, 40, 421, 23))
        self.pathLineEdit.setReadOnly(True)
        self.pathLineEdit.setObjectName("pathLineEdit")
        self.browseButton = QtWidgets.QPushButton(self.centralwidget)
        self.browseButton.setGeometry(QtCore.QRect(550, 40, 91, 23))
        self.browseButton.setObjectName("browseButton")
        self.selectedFile = QtWidgets.QLabel(self.centralwidget)
        self.selectedFile.setGeometry(QtCore.QRect(10, 40, 111, 21))
        font = QtGui.QFont()
        font.setFamily("Bahnschrift")
        font.setPointSize(14)
        self.selectedFile.setFont(font)
        self.selectedFile.setScaledContents(False)
        self.selectedFile.setWordWrap(True)
        self.selectedFile.setObjectName("selectedFile")
        self.selectedFile_2 = QtWidgets.QLabel(self.centralwidget)
        self.selectedFile_2.setGeometry(QtCore.QRect(10, 100, 161, 21))
        font = QtGui.QFont()
        font.setFamily("Bahnschrift")
        font.setPointSize(14)
        self.selectedFile_2.setFont(font)
        self.selectedFile_2.setScaledContents(False)
        self.selectedFile_2.setWordWrap(True)
        self.selectedFile_2.setObjectName("selectedFile_2")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(180, 100, 69, 22))
        self.comboBox.setObjectName("comboBox")
        self.checkButton = QtWidgets.QPushButton(self.centralwidget)
        self.checkButton.setGeometry(QtCore.QRect(260, 100, 61, 23))
        self.checkButton.setObjectName("checkButton")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(10, 170, 316, 260))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.tableWidget.setFont(font)
        self.tableWidget.setMouseTracking(True)
        self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setDragDropOverwriteMode(False)
        self.tableWidget.setGridStyle(QtCore.Qt.DashLine)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        item.setBackground(QtGui.QColor(255, 255, 255))
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(4, item)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(False)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(100)
        self.saveButton = QtWidgets.QPushButton(self.centralwidget)
        self.saveButton.setGeometry(QtCore.QRect(346, 400, 291, 30))
        self.saveButton.setObjectName("saveButton")
        self.infoLabel = QtWidgets.QLabel(self.centralwidget)
        self.infoLabel.setGeometry(QtCore.QRect(426, 180, 98, 20))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.infoLabel.setFont(font)
        self.infoLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.infoLabel.setObjectName("infoLabel")
        self.maxScoreText = QtWidgets.QLabel(self.centralwidget)
        self.maxScoreText.setGeometry(QtCore.QRect(350, 270, 71, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.maxScoreText.setFont(font)
        self.maxScoreText.setWordWrap(True)
        self.maxScoreText.setObjectName("maxScoreText")
        self.minScoreText = QtWidgets.QLabel(self.centralwidget)
        self.minScoreText.setGeometry(QtCore.QRect(350, 324, 71, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.minScoreText.setFont(font)
        self.minScoreText.setWordWrap(True)
        self.minScoreText.setObjectName("minScoreText")
        self.totalStudentText = QtWidgets.QLabel(self.centralwidget)
        self.totalStudentText.setGeometry(QtCore.QRect(350, 220, 71, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.totalStudentText.setFont(font)
        self.totalStudentText.setWordWrap(True)
        self.totalStudentText.setObjectName("totalStudentText")
        self.maxScoreValue = QtWidgets.QLabel(self.centralwidget)
        self.maxScoreValue.setGeometry(QtCore.QRect(410, 285, 71, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.maxScoreValue.setFont(font)
        self.maxScoreValue.setText("")
        self.maxScoreValue.setAlignment(QtCore.Qt.AlignCenter)
        self.maxScoreValue.setObjectName("maxScoreValue")
        self.minScoreValue = QtWidgets.QLabel(self.centralwidget)
        self.minScoreValue.setGeometry(QtCore.QRect(410, 335, 71, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.minScoreValue.setFont(font)
        self.minScoreValue.setText("")
        self.minScoreValue.setAlignment(QtCore.Qt.AlignCenter)
        self.minScoreValue.setObjectName("minScoreValue")
        self.totalStudentValue = QtWidgets.QLabel(self.centralwidget)
        self.totalStudentValue.setGeometry(QtCore.QRect(410, 235, 71, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.totalStudentValue.setFont(font)
        self.totalStudentValue.setText("")
        self.totalStudentValue.setAlignment(QtCore.Qt.AlignCenter)
        self.totalStudentValue.setObjectName("totalStudentValue")
        self.selectedFile_3 = QtWidgets.QLabel(self.centralwidget)
        self.selectedFile_3.setGeometry(QtCore.QRect(400, 100, 161, 21))
        font = QtGui.QFont()
        font.setFamily("Bahnschrift")
        font.setPointSize(14)
        self.selectedFile_3.setFont(font)
        self.selectedFile_3.setScaledContents(False)
        self.selectedFile_3.setWordWrap(True)
        self.selectedFile_3.setObjectName("selectedFile_3")
        self.sheetDropbox = QtWidgets.QComboBox(self.centralwidget)
        self.sheetDropbox.setGeometry(QtCore.QRect(510, 100, 69, 22))
        self.sheetDropbox.setObjectName("sheetDropbox")
        self.maxScoreText_2 = QtWidgets.QLabel(self.centralwidget)
        self.maxScoreText_2.setGeometry(QtCore.QRect(500, 271, 71, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.maxScoreText_2.setFont(font)
        self.maxScoreText_2.setWordWrap(True)
        self.maxScoreText_2.setObjectName("maxScoreText_2")
        self.maxScoreValue_2 = QtWidgets.QLabel(self.centralwidget)
        self.maxScoreValue_2.setGeometry(QtCore.QRect(560, 286, 71, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.maxScoreValue_2.setFont(font)
        self.maxScoreValue_2.setText("")
        self.maxScoreValue_2.setAlignment(QtCore.Qt.AlignCenter)
        self.maxScoreValue_2.setObjectName("maxScoreValue_2")
        self.minScoreValue_2 = QtWidgets.QLabel(self.centralwidget)
        self.minScoreValue_2.setGeometry(QtCore.QRect(560, 336, 71, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.minScoreValue_2.setFont(font)
        self.minScoreValue_2.setText("")
        self.minScoreValue_2.setAlignment(QtCore.Qt.AlignCenter)
        self.minScoreValue_2.setObjectName("minScoreValue_2")
        self.totalStudentText_2 = QtWidgets.QLabel(self.centralwidget)
        self.totalStudentText_2.setGeometry(QtCore.QRect(500, 221, 71, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.totalStudentText_2.setFont(font)
        self.totalStudentText_2.setWordWrap(True)
        self.totalStudentText_2.setObjectName("totalStudentText_2")
        self.minScoreText_2 = QtWidgets.QLabel(self.centralwidget)
        self.minScoreText_2.setGeometry(QtCore.QRect(500, 325, 71, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.minScoreText_2.setFont(font)
        self.minScoreText_2.setWordWrap(True)
        self.minScoreText_2.setObjectName("minScoreText_2")
        self.totalStudentValue_2 = QtWidgets.QLabel(self.centralwidget)
        self.totalStudentValue_2.setGeometry(QtCore.QRect(560, 236, 71, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.totalStudentValue_2.setFont(font)
        self.totalStudentValue_2.setText("")
        self.totalStudentValue_2.setAlignment(QtCore.Qt.AlignCenter)
        self.totalStudentValue_2.setObjectName("totalStudentValue_2")
        self.pathLineEdit.raise_()
        self.browseButton.raise_()
        self.selectedFile.raise_()
        self.selectedFile_2.raise_()
        self.comboBox.raise_()
        self.checkButton.raise_()
        self.saveButton.raise_()
        self.infoLabel.raise_()
        self.maxScoreText.raise_()
        self.minScoreText.raise_()
        self.totalStudentText.raise_()
        self.maxScoreValue.raise_()
        self.minScoreValue.raise_()
        self.totalStudentValue.raise_()
        self.selectedFile_3.raise_()
        self.sheetDropbox.raise_()
        self.tableWidget.raise_()
        self.maxScoreText_2.raise_()
        self.maxScoreValue_2.raise_()
        self.minScoreValue_2.raise_()
        self.totalStudentText_2.raise_()
        self.minScoreText_2.raise_()
        self.totalStudentValue_2.raise_()
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 650, 21))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuRecent_Files = QtWidgets.QMenu(self.menuFile)
        self.menuRecent_Files.setObjectName("menuRecent_Files")
        self.menuSettings = QtWidgets.QMenu(self.menubar)
        self.menuSettings.setObjectName("menuSettings")
        self.menuSettings_2 = QtWidgets.QMenu(self.menuSettings)
        self.menuSettings_2.setObjectName("menuSettings_2")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.actionNew = QtWidgets.QAction(MainWindow)
        self.actionNew.setShortcutVisibleInContextMenu(True)
        self.actionNew.setObjectName("actionNew")
        self.actionSave_As = QtWidgets.QAction(MainWindow)
        self.actionSave_As.setShortcutVisibleInContextMenu(True)
        self.actionSave_As.setObjectName("actionSave_As")
        self.actionCtrl_O = QtWidgets.QAction(MainWindow)
        self.actionCtrl_O.setObjectName("actionCtrl_O")
        self.actionOpen_ = QtWidgets.QAction(MainWindow)
        self.actionOpen_.setShortcutVisibleInContextMenu(True)
        self.actionOpen_.setObjectName("actionOpen_")
        self.actionTest1 = QtWidgets.QAction(MainWindow)
        self.actionTest1.setObjectName("actionTest1")
        self.actionTest2 = QtWidgets.QAction(MainWindow)
        self.actionTest2.setObjectName("actionTest2")
        self.actionSave_2 = QtWidgets.QAction(MainWindow)
        self.actionSave_2.setShortcutVisibleInContextMenu(True)
        self.actionSave_2.setObjectName("actionSave_2")
        self.actionCheck_For_Updates = QtWidgets.QAction(MainWindow)
        self.actionCheck_For_Updates.setObjectName("actionCheck_For_Updates")
        self.actionClear_Recent = QtWidgets.QAction(MainWindow)
        self.actionClear_Recent.setObjectName("actionClear_Recent")
        self.actionClear_Recent_Files = QtWidgets.QAction(MainWindow)
        self.actionClear_Recent_Files.setObjectName("actionClear_Recent_Files")
        self.menuFile.addAction(self.actionOpen_)
        self.menuFile.addAction(self.menuRecent_Files.menuAction())
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionSave_2)
        self.menuSettings_2.addAction(self.actionClear_Recent_Files)
        self.menuSettings.addAction(self.menuSettings_2.menuAction())
        self.menuHelp.addAction(self.actionCheck_For_Updates)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuSettings.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "XLSX Reader"))
        self.browseButton.setText(_translate("MainWindow", "Browse"))
        self.selectedFile.setText(_translate("MainWindow", "Selected File :"))
        self.selectedFile_2.setText(_translate("MainWindow", "Select Subject Code :"))
        self.checkButton.setText(_translate("MainWindow", "Check"))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "Roll. No."))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "Gender"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "Name"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "Marks"))
        item = self.tableWidget.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "Grade"))
        self.saveButton.setText(_translate("MainWindow", "Save"))
        self.infoLabel.setText(_translate("MainWindow", "Information"))
        self.maxScoreText.setText(_translate("MainWindow", "Highest Score"))
        self.minScoreText.setText(_translate("MainWindow", "Lowest Score"))
        self.totalStudentText.setText(_translate("MainWindow", "No. Of Students"))
        self.selectedFile_3.setText(_translate("MainWindow", "Select Sheet :"))
        self.maxScoreText_2.setText(_translate("MainWindow", "Highest Scorer"))
        self.totalStudentText_2.setText(_translate("MainWindow", "Class Average"))
        self.minScoreText_2.setText(_translate("MainWindow", "Lowest Scorer"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuRecent_Files.setTitle(_translate("MainWindow", "Recent Files"))
        self.menuSettings.setTitle(_translate("MainWindow", "Edit"))
        self.menuSettings_2.setTitle(_translate("MainWindow", "Settings"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.actionNew.setText(_translate("MainWindow", "New"))
        self.actionNew.setShortcut(_translate("MainWindow", "Ctrl+N"))
        self.actionSave_As.setText(_translate("MainWindow", "Save As"))
        self.actionSave_As.setShortcut(_translate("MainWindow", "Ctrl+Shift+S"))
        self.actionCtrl_O.setText(_translate("MainWindow", "Ctrl+O"))
        self.actionOpen_.setText(_translate("MainWindow", "Open"))
        self.actionOpen_.setShortcut(_translate("MainWindow", "Ctrl+O"))
        self.actionTest1.setText(_translate("MainWindow", "Test1"))
        self.actionTest2.setText(_translate("MainWindow", "Test2"))
        self.actionSave_2.setText(_translate("MainWindow", "Save"))
        self.actionSave_2.setShortcut(_translate("MainWindow", "Ctrl+S"))
        self.actionCheck_For_Updates.setText(_translate("MainWindow", "Check For Updates"))
        self.actionClear_Recent.setText(_translate("MainWindow", "Clear Recent Files"))
        self.actionClear_Recent_Files.setText(_translate("MainWindow", "Clear Recent Files"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

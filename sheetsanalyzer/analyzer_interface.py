# !/usr/bin/python3.7
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QComboBox, QCheckBox, QRadioButton, QDesktopWidget, QMainWindow, QGridLayout, QHBoxLayout, QFileDialog, QProgressBar, QMessageBox, QAction, QFrame, QVBoxLayout, QPushButton
import sys
from analyzer_manager import FileManager
from analyzer_run import Analyzer
from re import findall, split, compile
import datetime


class sheetsanalyzer(QMainWindow):

    def __init__(self):
        super().__init__()
        self.title = "SheetsAnalyzer"
        self.width = 400
        self.height = 320
        self._is_resetet = False
        self.setWindowTitle(self.title)
        self.setFixedSize(self.width, self.height)
        self.initUI()

    def initUI(self):
        self.CenterWindow()
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        grid_layout = QGridLayout()
        grid_layout.setContentsMargins(20, 5, 20, 5)
        central_widget.setLayout(grid_layout)

        self.create_toolbar()
        self.create_statusbar()

        folder_head = QLabel("Path")
        folder_head.setMaximumWidth(80)
        self.linee_folder = QLineEdit()
        self.linee_folder.editingFinished.connect(self.set_input_folderpath)
        self.push_folder = QPushButton("Browse...")
        self.push_folder.clicked.connect(self.set_folderpath)
        self.check_subfolder = QCheckBox("subfolder")
        self.check_subfolder.clicked.connect(self.set_subfolder)
        self.combo_filetyp = QComboBox()
        self.combo_filetyp.addItems(["", "csv", "txt", "xls", "xlsx", "xlsm"])
        self.combo_filetyp.currentTextChanged.connect(self.set_filetyp)
        keyword_head = QLabel("Keywords")
        keyword_head.setMaximumWidth(80)
        self.linee_search = QLineEdit()
        self.linee_search.editingFinished.connect(self.set_keywords)
        range_head = QLabel("Cell / Range")
        range_head.setMaximumWidth(80)
        self.linee_range = QLineEdit()
        self.linee_range.editingFinished.connect(self.set_range)

        sheets_head = QLabel("Sheets")
        sheets_head.setMaximumWidth(80)
        h1_sheet_layout = QHBoxLayout()
        self.radio_all = QRadioButton("all")
        self.radio_all.setChecked(True)
        self.radio_all.clicked.connect(self.set_all_sheets)
        self.radio_ran = QRadioButton("range")
        self.radio_ran.clicked.connect(self.set_all_sheets)
        self.check_hidden = QCheckBox("hidden sheets")
        self.check_hidden.clicked.connect(self.set_hidden_sheets)
        h1_sheet_layout.addWidget(self.radio_all)
        h1_sheet_layout.addWidget(self.radio_ran)

        self.linee_sheets = QLineEdit()
        self.linee_sheets.setEnabled(False)
        self.linee_sheets.editingFinished.connect(self.set_sheets_range)
        warn_label = QLabel(
            "*To execute the program it is necessary that the files have the same structure.")
        warn_label.setWordWrap(True)

        bar_layout = QVBoxLayout()
        progress_layout = QHBoxLayout()

        self.progress_label = QLabel("File:")
        self.progress_time = QLabel("Runtime:")
        self.progress_timer = QTimer(self)
        self.progress_timer.timeout.connect(self.runTime)
        self.progress_timer.setInterval(1)
        self.mscounter = 0
        self.setRunTime()

        progress_layout.addWidget(self.progress_label)
        progress_layout.addSpacing(165)
        progress_layout.addWidget(self.progress_time)

        self.progressBar = QProgressBar()
        self.progressBar.setRange(0, 100)
        self.progressBar.setAlignment(Qt.AlignCenter)
        bar_layout.addLayout(progress_layout)
        bar_layout.addWidget(self.progressBar)

        grid_layout.addWidget(folder_head, 0, 0)
        grid_layout.addWidget(self.linee_folder, 0, 1, 1, 2)
        grid_layout.addWidget(self.push_folder, 0, 3)
        grid_layout.addWidget(self.check_subfolder, 1, 2, Qt.AlignRight)
        grid_layout.addWidget(self.combo_filetyp, 1, 3)
        grid_layout.addWidget(keyword_head, 3, 0)
        grid_layout.addWidget(self.linee_search, 3, 1, 1, 3)
        grid_layout.addWidget(range_head, 5, 0)
        grid_layout.addWidget(self.linee_range, 5, 1, 1, 3)
        grid_layout.addWidget(sheets_head, 6, 0)
        grid_layout.addLayout(h1_sheet_layout, 6, 2)
        grid_layout.addWidget(self.check_hidden, 6, 3)
        grid_layout.addWidget(self.linee_sheets, 7, 2, 1, 2)
        grid_layout.addWidget(QHLine(), 8, 0, 1, 4)
        grid_layout.addWidget(warn_label, 9, 0, 1, 4)
        grid_layout.addWidget(QHLine(), 10, 0, 1, 4)
        grid_layout.addLayout(bar_layout, 11, 0, 1, 4, Qt.AlignCenter)

        self.set_file_settings(False)
        self.set_range_settings(False)
        self.set_sheet_settings(False)

        self.filemanager = FileManager()

    def CenterWindow(self):
        geo_frame = self.frameGeometry()
        c_origin = QDesktopWidget().availableGeometry().center()
        geo_frame.moveCenter(c_origin)
        self.move(geo_frame.topLeft())

    def create_toolbar(self):
        startAct = QAction(QIcon("ressources/img/btn_start.png"), "Start", self)
        startAct.setShortcut("Ctrl+A")
        startAct.triggered.connect(self.start_analysis)
        cancelAct = QAction(QIcon("ressources/img/btn_cancel.png"), "Cancel", self)
        cancelAct.setShortcut("Ctrl+C")
        cancelAct.triggered.connect(self.reset_ui)

        self.toolbar = self.addToolBar("Work")
        self.toolbar.addAction(startAct)
        self.toolbar.addAction(cancelAct)

    def create_statusbar(self):
        self.info_widget = QLabel()
        self.setFileInfo()

        self.range_widget = QLabel()
        self.range_widget.setAlignment(Qt.AlignRight)
        self.setRangeInfo()

        self.status_bar = self.statusBar()
        self.status_bar.setSizeGripEnabled(False)
        self.status_bar.setContentsMargins(20, 0, 20, 0)
        self.status_bar.addPermanentWidget(self.info_widget, 1)
        self.status_bar.addPermanentWidget(self.range_widget, 1)

    def UserCloseEvent(self):
        self.close()

    def set_file_settings(self, mode):

        self.check_subfolder.setEnabled(mode)
        self.combo_filetyp.setEnabled(mode)

    def set_range_settings(self, mode):
        self.linee_search.setEnabled(mode)
        self.linee_range.setEnabled(mode)

        if mode:
            self.linee_range.setPlaceholderText("For example: A1;A2:B4")
        else:
            self.linee_range.setPlaceholderText("")

    def set_sheet_settings(self, mode):
        self.radio_all.setEnabled(mode)
        self.radio_ran.setEnabled(mode)
        self.check_hidden.setEnabled(mode)

    def set_folderpath(self):
        folderpath = str(QFileDialog.getExistingDirectory(self, "Select Directory"))

        if folderpath:
            self.linee_folder.setText(folderpath)
            self.linee_folder.setCursorPosition(0)
            self.filemanager.setPath(folderpath)

            # self.set_files()
            self.filemanager.setFiles()
            self.setFileInfo(self.filemanager.count_files())
            self.check_subfolder.setEnabled(True)
            self.combo_filetyp.setEnabled(True)

    def set_input_folderpath(self):
        folderpath = self.linee_folder.text()

        if folderpath:
            self.filemanager.setPath(folderpath)
            self.linee_folder.setCursorPosition(0)

            self.filemanager.setFiles()
            self.setFileInfo(self.filemanager.count_files())
            self.check_subfolder.setEnabled(True)
            self.combo_filetyp.setEnabled(True)

    def set_files(self):
        if self.filemanager.getPath():
            self.setFileInfo(self.filemanager.count_files())
            self.setRangeInfo(self.filemanager.getRangeText())

    def set_subfolder(self):
        self.filemanager.setSubFolder(self.check_subfolder.isChecked())
        self.set_files()

    def set_filetyp(self, text):
        self.filemanager.setFileTyp(text)

        if self.filemanager.files_exist() and self.filemanager.getFileTyp():
            self.set_range_settings(True)
            if text in ["xls", "xlsx", "xlsm"]:
                self.set_sheet_settings(True)
            else:
                self.set_sheet_settings(False)
        else:
            self.set_sheet_settings(False)
            self.set_range_settings(False)
        self.set_files()

    def set_keywords(self):
        if not self._is_resetet:
            keywords = self.linee_search.text().split(";")
            self.filemanager.setKeywords(keywords)
            self.setFileInfo(self.filemanager.count_files())

    def set_range(self):
        if not self._is_resetet:
            cell_range = self.linee_range.text().upper()
            self.linee_range.setText(cell_range)

            wrong_chars = any(findall("[,@_!#$%^&*()<>?/|}{~ ]", cell_range))
            if wrong_chars:
                QMessageBox.warning(self, "Error",
                                    "Cell range contains one or more wrong characters.", QMessageBox.Ok)

            cells_patttern = compile("\w+\d+")
            wrong_pattern = all([True if not cells_patttern.match(
                Cell) else False for Cell in split(";|:", cell_range)])

            if wrong_pattern:
                QMessageBox.warning(self, "Error",
                                    "Entered cells have an incorrect format. Cells must have a format similar to Excel: A1 or BG67 (for example).", QMessageBox.Ok)
            if cell_range:
                self.filemanager.setUserRange(cell_range.split(";"))

    def set_hidden_sheets(self):
        self.filemanager.setHiddenSheetsSettings(self.check_hidden.isChecked())

    def set_all_sheets(self):
        self.filemanager.setAllSheets(self.radio_all.isChecked())

        if not self.radio_all.isChecked():
            self.linee_sheets.setEnabled(True)
            self.linee_sheets.setPlaceholderText("For Example: Table1;Table2;1:4")
        else:
            self.linee_sheets.setEnabled(False)
            self.linee_sheets.setPlaceholderText("")

    def set_sheets_range(self):
        sheets_range = self.linee_sheets.text().split(";")
        self.filemanager.setUserSheetNames(sheets_range)

    def start_analysis(self):
        print(self.filemanager.attributes())
        if self.filemanager.ready_to_run():
            self.progress_timer.start()
            self.progressBar.setTextVisible(True)
            self.analyzer = Analyzer(self.filemanager.getFiles(),
                                     "xlsx",
                                     ";",
                                     self.filemanager.getKeywords(),
                                     self.filemanager.getFileLength(),
                                     self.filemanager.getUserRange(),
                                     self.filemanager.getUserCellRange(),
                                     self.filemanager.getAllSheets(),
                                     self.filemanager.getHiddenSheetsSettings(),
                                     self.filemanager.getUserSheetIndexes())
            self.analyzer.countChanged.connect(self.onCountChanged)
            self.analyzer.finished.connect(self.onFinished)
            self.analyzer.start()
        else:
            QMessageBox.warning(self, "Error",
                                "No files could be found with the settings made.", QMessageBox.Ok)

    def reset_ui(self):
        self._is_resetet = True
        self.linee_folder.setText("")
        self.linee_search.setText("")
        self.linee_range.setText("")
        self.combo_filetyp.setCurrentIndex(-1)
        self.set_file_settings(False)
        self.set_sheet_settings(False)

        self.filemanager.basic_settings()
        self.setFileInfo()
        self.setRangeInfo()

        self.progress_label.setText("File:")
        self.progressBar.setValue(0)
        self.progressBar.setTextVisible(False)

        self.mscounter = 0
        self._is_resetet = False

    def onCountChanged(self, value_pro, value_file):
        maximum = self.filemanager.count_files()
        width = len(str(maximum))

        self.progress_label.setText("File: {:<{w}} / {:<{w}}".format(value_file, maximum, w=width))
        self.progressBar.setValue(value_pro)

    def onFinished(self):
        self.progress_label.setText(self.progress_label.text() + " --> Done!")
        self.progress_timer.stop()

    def setFileInfo(self, value=None):
        if not value:
            self.info_widget.setText("Files: {:>10}".format(" - not defined - "))
        else:
            self.info_widget.setText("Files: {:>10}".format(value, "."))

    def setRangeInfo(self, text=None):
        if not text:
            self.range_widget.setText("Range: {:<10}".format(" - not defined - "))
        else:
            self.range_widget.setText("Files: {:>10}".format(text))

    def setRunTime(self):
        tdelta = datetime.timedelta(milliseconds=self.mscounter)

        s = tdelta.seconds
        milliseconds = tdelta.microseconds / 1000
        hours, remainder = divmod(s, 3600)
        minutes, seconds = divmod(remainder, 60)

        self.progress_time.setText("Runtime: {:02}:{:02}:{:02}.{:03}".format(
            int(hours), int(minutes), int(seconds), int(milliseconds)))

    def runTime(self):
        self.mscounter += 1
        self.setRunTime()


class QHLine(QFrame):
    def __init__(self):
        super(QHLine, self).__init__()
        self.setFrameShape(QFrame.HLine)
        self.setFrameShadow(QFrame.Sunken)


if __name__ == "__main__":
    # Abfrage, ob Programm noch läuft oder nicht
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()

    # Assignment of gui design
    # app.setWindowIcon(SetWindowAppIcon())
    # with open("E:\\Python_Projekte\\SheetsAnalyzer_dev\\SheetsAnalyzer_v2_Stylesheet.css") as fh:  # , "r"
        # app.setStyleSheet(fh.read())

    # app.setStyleSheet(stylesheet)
    GUI = sheetsanalyzer()
    GUI.show()
    app.exec_()

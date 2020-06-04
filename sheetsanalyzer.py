from analyzer_interface import sheetsanalyzer
from PyQt5.QtWidgets import QApplication
import sys


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

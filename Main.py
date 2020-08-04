import sys

from PyQt5.QtWidgets import QApplication

from MainWindowMain import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainwindow = MainWindow()
    mainwindow.show()
    sys.exit(app.exec_())
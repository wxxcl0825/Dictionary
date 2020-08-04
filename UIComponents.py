import os

from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

import Fliter_mUI
import Fliter_sUI
import ModeUI
import NewUI
import PathUI
import UnionUI
from FileManager import File
from LibManager import Lib


class New(QDialog, NewUI.Ui_Dialog):
    def __init__(self, is_lemmas, gui):
        QDialog.__init__(self)
        NewUI.Ui_Dialog.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.f_btn)
        self.pushButton_2.clicked.connect(self.f_btn2)
        self.File = File()
        self.Lib = Lib()
        self.text = []
        self.is_lemmas = is_lemmas
        self.gui = gui

    def f_btn(self):
        if len(self.textEdit.toPlainText()) == 0:
            return
        self.pushButton.setEnabled(False)
        self.text = sorted(list((set(
            self.File.anafile(self.textEdit.toPlainText() + ' ', self.gui, is_lemmas=self.is_lemmas, title=self)))))
        self.setWindowTitle("新建词库")
        self.textEdit.setText(
            str(self.text)[1:-1].translate(str.maketrans(',', '\n', '\'')).translate(str.maketrans('', '', ' ')))
        self.pushButton.setEnabled(True)

    def f_btn2(self):
        savepath = QFileDialog.getSaveFileName(self, "另存为", self.File.getdesktop(), "Pickled File(*.pkl)")[0]
        if savepath != "":
            self.Lib.savelib(set(self.text), savepath)


class Path(QDialog, PathUI.Ui_Dialog):
    def __init__(self, settings):
        QDialog.__init__(self)
        PathUI.Ui_Dialog.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.f_btn)
        self.pushButton_2.clicked.connect(self.f_btn2)
        self.pushButton_3.clicked.connect(self.f_btn3)
        self.pushButton_4.clicked.connect(self.f_btn4)
        self.File = File()
        self.settings = settings
        self.default_settingspath = os.path.dirname(os.path.abspath(__file__)).translate(
            str.maketrans('\\', '/', '')) + "/Resources/settings.pkl"
        self.lineEdit.setText(self.settings[0])
        self.lineEdit_2.setText(self.settings[1])
        self.lineEdit_3.setText(self.settings[4])

    def f_btn(self):
        path = QFileDialog.getOpenFileName(self, "设置路径", self.File.getdesktop(), "Pickled File(*.pkl)")[0]
        if path != "":
            self.lineEdit.setText(path)

    def f_btn2(self):
        path = QFileDialog.getSaveFileName(self, "设置路径", self.File.getdesktop(), "Microsoft Word File(*.docx)")[0]
        if path != "":
            self.lineEdit_2.setText(path)

    def f_btn3(self):
        path = QFileDialog.getOpenFileName(self, "设置路径", self.File.getdesktop(), "Pickled File(*.pkl)")[0]
        if path != "":
            self.lineEdit_3.setText(path)

    def f_btn4(self):
        if self.lineEdit_3.text() != self.default_settingspath:
            try:
                self.settings = self.File.readsettings(self.settings[4])
            except:
                QMessageBox.critical(self, "错误", "加载失败！请检查配置路径是否正确，文件是否损坏！")
                return
        self.settings[0] = self.lineEdit.text()
        self.settings[1] = self.lineEdit_2.text()
        self.settings[4] = self.default_settingspath
        self.File.savesettings(self.settings[4], self.settings)
        self.hide()


class Union(QDialog, UnionUI.Ui_Dialog):
    def __init__(self):
        QDialog.__init__(self)
        UnionUI.Ui_Dialog.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.f_btn)
        self.pushButton_2.clicked.connect(self.f_btn2)
        self.pushButton_3.clicked.connect(self.f_btn3)
        self.Lib = Lib()
        self.File = File()

    def f_btn(self):
        lib1path = QFileDialog.getOpenFileName(self, "打开文件", self.File.getdesktop(), "Pickled File(*.pkl)")[0]
        self.lineEdit.setText(lib1path)

    def f_btn2(self):
        lib2path = QFileDialog.getOpenFileName(self, "打开文件", self.File.getdesktop(), "Pickled File(*.pkl)")[0]
        self.lineEdit_2.setText(lib2path)

    def f_btn3(self):
        text1 = self.lineEdit.text()
        text2 = self.lineEdit_2.text()
        if text1 != "" and text2 != "":
            savepath = QFileDialog.getSaveFileName(self, "另存为", self.File.getdesktop(), "Pickled File(*.pkl)")[0]
            if savepath != "":
                try:
                    self.Lib.unionlib(text1, text2, savepath)
                except:
                    QMessageBox.critical(self, "错误", "加载失败！请检查路径是否正确，文件是否损坏！")
        else:
            QMessageBox.information(self, "系统消息", "请将路径填写完整！")


class Fliter_s(QDialog, Fliter_sUI.Ui_Dialog):
    def __init__(self, word, result, title):
        QDialog.__init__(self)
        Fliter_sUI.Ui_Dialog.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.f_btn)
        self.pushButton_2.clicked.connect(self.f_btn2)
        self.setWindowTitle(title)
        self.label.setText(word)
        self.result = result

    def keyPressEvent(self, QKeyEvent):
        if QKeyEvent.key() == Qt.Key_Y:
            self.f_btn()
        elif QKeyEvent.key() == Qt.Key_N:
            self.f_btn2()

    def mousePressEvent(self, QMouseEvent):
        if QMouseEvent.buttons() == Qt.LeftButton:
            self.f_btn()
        elif QMouseEvent.buttons() == Qt.RightButton:
            self.f_btn2()

    def f_btn(self):
        self.result = True
        self.hide()

    def f_btn2(self):
        self.result = False
        self.hide()


class Fliter_m(QDialog, Fliter_mUI.Ui_Dialog):
    def __init__(self, wordlist, results):
        QDialog.__init__(self)
        Fliter_mUI.Ui_Dialog.__init__(self)
        self.setupUi(self)
        self.checkboxes = []
        self.results = results
        if len(wordlist) < 11:
            height = 39 * len(wordlist) + 71
            self.setMaximumSize(QSize(310, height))
            self.setMinimumSize(QSize(310, height))
            self.scrollArea.setMaximumSize(QSize(310, height))
            self.scrollArea.setMinimumSize(QSize(310, height))
            self.scrollAreaWidgetContents.setMaximumSize(QSize(310 - 2, height - 2))
            self.scrollAreaWidgetContents.setMinimumSize(QSize(310 - 2, height - 2))
        for word in wordlist:
            self.addword(word)

    def addword(self, word):
        self.frame = QtWidgets.QFrame(self.scrollAreaWidgetContents)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.frame)
        self.label = QtWidgets.QLabel(self.frame)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Consolas")
        font.setPointSize(11)
        self.label.setFont(font)
        self.horizontalLayout.addWidget(self.label)
        self.checkboxes.append(QtWidgets.QCheckBox(self.frame))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkboxes[-1].sizePolicy().hasHeightForWidth())
        self.checkboxes[-1].setSizePolicy(sizePolicy)
        self.checkboxes[-1].setText("")
        self.horizontalLayout.addWidget(self.checkboxes[-1])
        self.verticalLayout.addWidget(self.frame)
        self.label.setText(word)

    def closeEvent(self, event):
        for checkbox in self.checkboxes:
            self.results.append(checkbox.isChecked())


class Mode(QDialog, ModeUI.Ui_Dialog):
    def __init__(self, settings):
        QDialog.__init__(self)
        ModeUI.Ui_Dialog.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.f_btn)
        self.settings = settings
        self.default_settingspath = os.path.dirname(os.path.abspath(__file__)).translate(
            str.maketrans('\\', '/', '')) + "/Resources/settings.pkl"
        self.File = File()
        self.radioButton.setChecked(self.settings[2])
        self.radioButton_2.setChecked(not self.settings[2])
        self.radioButton_3.setChecked(self.settings[3] == "Single")
        self.radioButton_4.setChecked(self.settings[3] == "Multi")
        self.radioButton_5.setChecked(self.settings[6])
        self.radioButton_6.setChecked(not self.settings[6])
        self.radioButton_7.setChecked(self.settings[5])
        self.radioButton_8.setChecked(not self.settings[5])

    def f_btn(self):
        self.settings[2] = self.radioButton.isChecked()
        if self.radioButton_3.isChecked():
            self.settings[3] = "Single"
        else:
            self.settings[3] = "Multi"
        self.settings[6] = self.radioButton_5.isChecked()
        self.settings[5] = self.radioButton_7.isChecked()
        self.File.savesettings(self.default_settingspath, self.settings)
        self.hide()

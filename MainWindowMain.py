import os
from time import sleep

import win32com.client as wincl
from PyQt5 import QtGui
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox

from DicMaker import DicMaker
from FileManager import File
from LibManager import Lib
from MainWindowUI import Ui_MainWindow
from Searcher import Searcher
from UIComponents import New, Union, Path, Fliter_s, Fliter_m, Mode


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.actionhelp.triggered.connect(self.f_help)
        self.actionnew.triggered.connect(self.f_new)
        self.actionmode.triggered.connect(self.f_mode)
        self.actionunion.triggered.connect(self.f_union)
        self.actionsave.triggered.connect(self.f_save)
        self.actionsaveas.triggered.connect(self.f_saveas)
        self.actionopen.triggered.connect(self.f_open)
        self.actionpath.triggered.connect(self.f_path)
        self.pushButton.clicked.connect(self.f_btn)
        self.pushButton_2.clicked.connect(self.f_btn2)
        self.pushButton_3.clicked.connect(self.f_btn3)
        self.pushButton_4.clicked.connect(self.f_btn4)
        self.pushButton_5.clicked.connect(self.f_btn5)
        self.timer = QTimer(self.progressBar)
        self.timer2 = QTimer(self.progressBar_2)
        self.timer3 = QTimer(self.progressBar_3)
        self.timer4 = QTimer(self.progressBar_4)
        self.timer5 = QTimer(self.progressBar_5)
        self.timer.timeout.connect(lambda: self.upd_qpb(self.progressBar, self.timer))
        self.timer2.timeout.connect(lambda: self.upd_qpb(self.progressBar_2, self.timer2))
        self.timer3.timeout.connect(lambda: self.upd_qpb(self.progressBar_3, self.timer3))
        self.timer4.timeout.connect(lambda: self.upd_qpb(self.progressBar_4, self.timer4))
        self.timer5.timeout.connect(lambda: self.upd_qpb(self.progressBar_5, self.timer5))
        self.gui = QtGui.QGuiApplication.processEvents
        self.File = File()
        self.Lib = Lib()
        self.Searcher = Searcher()
        self.default_settingspath = os.path.dirname(os.path.abspath(__file__)).translate(
            str.maketrans('\\', '/', '')) + "/Resources/settings.pkl"
        self.Text = ""
        self.filepath = ""
        self.settings = ["", "", False, "Multi", self.default_settingspath, False, False]
        self.libpath, self.savepath, self.is_lemmas, self.mode, self.settingspath, self.is_header, self.is_check = self.settings
        self.lib = {}
        self.analyse_result, self.fliter_result = [], []
        self.missing_meanings, self.missing_pronounces = [], []
        self.flag_for_single = False
        self.show()
        try:
            self.settings = self.File.readsettings("Resources/settings.pkl")
            self.libpath, self.savepath, self.is_lemmas, self.mode, self.settingspath, self.is_header, self.is_check = self.settings
        except:
            QMessageBox.information(self, "系统消息", "配置文件丢失或损坏，请重新设置！")
            self.f_path()
            self.f_mode()
        if self.libpath != "":
            self.label_2.setText("等待加载")

    def keyPressEvent(self, QKeyEvent):
        if QKeyEvent.key() == Qt.Key_F5:
            self.flag_for_single = True

    def f_help(self):
        try:
            os.system("start {}/Resources/helper.html".format(os.path.dirname(os.path.abspath(__file__)).translate(
                str.maketrans('\\', '/', ''))))
        except:
            QMessageBox.information(self, "系统消息", "打开帮助文档失败！")

    def f_new(self):
        self.d_new = New(self.is_lemmas, self.gui)
        self.d_new.show()

    def f_mode(self):
        self.d_mode = Mode(self.settings)
        self.d_mode.show()
        while self.d_mode.isVisible():
            sleep(0.01)
            self.gui()
        self.libpath, self.savepath, self.is_lemmas, self.mode, self.settingspath, self.is_header, self.is_check = self.settings

    def f_union(self):
        self.d_union = Union()
        self.d_union.show()

    def f_save(self):
        try:
            self.File.savefile(self.filepath, self.textEdit.toPlainText())
        except:
            self.f_saveas()

    def f_saveas(self):
        self.filepath = QFileDialog.getSaveFileName(self, "另存为", self.File.getdesktop(), "Text File(*.txt)")[0]
        if self.filepath != "":
            self.File.savefile(self.filepath, self.textEdit.toPlainText())

    def f_open(self):
        self.filepath = QFileDialog.getOpenFileName(self, "打开文件", self.File.getdesktop(), "Text Files(*.txt)")[0]
        if self.filepath != "":
            self.Text = self.File.readfile(self.filepath)
            self.textEdit.setText(self.Text)

    def f_path(self):
        old_path = self.libpath
        self.d_path = Path(self.settings)
        self.d_path.show()
        while self.d_path.isVisible():
            sleep(0.01)
            self.gui()
        self.libpath, self.savepath, self.is_lemmas, self.mode, self.settingspath, self.is_header, self.is_check = self.settings
        if (self.lib != {} and old_path != self.libpath) or (old_path == "" and self.libpath != ""):
            self.label_2.setText("等待加载")
            self.progressBar_2.setValue(0)
            if self.label.text() == "分析完成":
                self.label_3.setText("请先加载词库")
            else:
                self.label_3.setText("请先分析原文")
            self.progressBar_3.setValue(0)
            self.label_4.setText("请先完成自动筛选")
            self.progressBar_4.setValue(0)
            self.label_5.setText("请先完成手动筛选")
            self.progressBar_5.setValue(0)
            if self.analyse_result != []:
                self.textEdit.setText(
                    str(self.analyse_result)[1:-1].translate(str.maketrans(',', ' ', '\'')))
        if self.label_5.text() == "请先设置输出路径" and self.savepath != "":
            if self.label_4.text() != "手动筛选完成":
                self.label_5.setText("请先完成手动筛选")
            else:
                self.label_5.setText("等待生成")
        if self.savepath == "":
            if self.label_4.text() != "手动筛选完成":
                self.label_5.setText("请先完成手动筛选")
            else:
                self.label_5.setText("请先设置输出路径")

    def f_btn(self):
        self.pushButton.setEnabled(False)
        self.progressBar.setValue(0)
        self.label.setText("分析中")
        self.gui()
        self.Text = self.textEdit.toPlainText() + ' '
        self.analyse_result = self.File.anafile(self.Text, self.gui, is_lemmas=self.is_lemmas, qpb=self.progressBar)
        self.textEdit.setText(
            str(self.analyse_result)[1:-1].translate(str.maketrans(',', ' ', '\'')))
        self.progressBar.setValue(100)
        self.label.setText("分析完成")
        if self.label_2.text() == "加载完成":
            self.label_3.setText("等待自动筛选")
            self.progressBar_3.setValue(0)
        else:
            self.label_3.setText("请先加载词库")
        self.label_4.setText("请先完成自动筛选")
        self.progressBar_4.setValue(0)
        self.label_5.setText("请先完成手动筛选")
        self.progressBar_5.setValue(0)
        self.pushButton.setEnabled(True)

    def f_btn2(self):
        if self.label_2.text() == "等待加载":
            try:
                self.lib = self.Lib.readlib(self.libpath)
            except:
                QMessageBox.critical(self, "错误", "加载失败！请检查路径是否正确，文件是否损坏！")
                return
            self.label_2.setText("加载中")
            self.gui()
            self.timer2.start(1)
            self.label_2.setText("加载完成")
            if self.label.text() == "分析完成":
                self.label_3.setText("等待自动筛选")
                self.progressBar_3.setValue(0)
            self.label_4.setText("请先完成自动筛选")
            self.progressBar_4.setValue(0)
            self.label_5.setText("请先完成手动筛选")
            self.progressBar_5.setValue(0)

    def f_btn3(self):
        if self.label_3.text() == "等待自动筛选":
            self.pushButton_3.setEnabled(False)
            self.progressBar_3.setValue(0)
            self.label_3.setText("自动筛选中")
            self.gui()
            self.timer3.start(1)
            self.fliter_result = sorted(list(self.File.fliterfile(self.analyse_result, self.lib)))
            self.textEdit.setText(str(self.fliter_result)[1:-1].translate(str.maketrans(',', ' ', '\'')))
            self.label_3.setText("自动筛选完成")
            self.label_4.setText("等待手动筛选")
            self.progressBar_4.setValue(0)
            self.label_5.setText("请先完成手动筛选")
            self.progressBar_4.setValue(0)
            self.pushButton_3.setEnabled(True)

    def f_btn4(self):
        if self.label_4.text() != "手动筛选完成" and self.label_4.text() != "等待手动筛选":
            return
        self.pushButton_4.setEnabled(False)
        self.label_4.setText("手动筛选中")
        self.gui()
        self.progressBar_4.setValue(0)
        if len(self.fliter_result) != 0:
            if self.mode == "Multi":
                self.multi_fliter()
                self.timer4.start(1)
            elif self.mode == "Single":
                self.flag_for_single = False
                self.single_fliter()
        self.textEdit.setText(str(self.fliter_result)[1:-1].translate(str.maketrans(',', '\n', '\'')).translate(
            str.maketrans('', '', ' ')))
        self.progressBar_4.setValue(100)
        self.label_4.setText("手动筛选完成")
        if self.savepath == "":
            self.label_5.setText("请先设置输出路径")
            self.progressBar_5.setValue(0)
        else:
            self.label_5.setText("等待生成")
            self.progressBar_5.setValue(0)
        self.pushButton_4.setEnabled(True)

    def f_btn5(self):
        if self.label_5.text() != "等待生成" and self.label_5.text() != "制作完成":
            return
        self.progressBar_5.setValue(0)
        if self.fliter_result == []:
            QMessageBox.information(self, "系统消息", "并没有可以制作的单词！")
            return
        if self.is_check:
            self.missing_pronounces = []
            self.missing_meanings = []
            self.label_5.setText("检查中")
            dl = 1 / len(self.fliter_result) * 100
            exact = 0
            for word in self.fliter_result:
                exact += dl
                self.progressBar_5.setValue(int(exact))
                self.gui()
                state = self.Searcher.search(word)[2]
                if state == -1:
                    self.missing_meanings.append(word)
                elif state == 0:
                    self.missing_pronounces.append(word)
            self.progressBar_5.setValue(100)
            text = ""
            if self.missing_meanings != []:
                text += "以下单词查询失败:\n{}\n".format(
                    str(self.missing_meanings)[1:-1].translate(str.maketrans(',', '\n', '\'')).translate(
                        str.maketrans('', '', ' ')))
            if self.missing_pronounces != []:
                text += "以下单词未查询到音标:\n{}".format(
                    str(self.missing_pronounces)[1:-1].translate(str.maketrans(',', '\n', '\'')).translate(
                        str.maketrans('', '', ' ')))
            if self.missing_meanings == [] and self.missing_pronounces == []:
                self.textEdit.setText("未检查到异常单词")
            else:
                self.textEdit.setText(text)
            result = QMessageBox.information(self, "系统消息", "检查成功！点击确定继续制作，点击取消停止制作！", QMessageBox.Yes | QMessageBox.No,
                                             QMessageBox.Yes)
            if result == QMessageBox.No:
                self.label_5.setText("等待生成")
                self.progressBar_5.setValue(0)
                return
            self.progressBar_5.setValue(0)
            self.label_5.setText("检查完成")
        if self.is_header:
            QMessageBox.information(self, "系统消息", "制作过程中请不要点击屏幕！")
        self.label_5.setText("制作中")
        try:
            self.DicMaker = DicMaker(self.fliter_result, self.savepath, self.gui, self.progressBar_5,
                                     self.is_header, self.label_5, self.Searcher.pre_load)
        except Exception as e:
            QMessageBox.critical(self, "错误", "错误信息：\n{}".format(e))
            self.label_5.setText("等待生成")
            self.progressBar_5.setValue(0)
            return
        self.progressBar_5.setValue(100)
        self.label_5.setText("制作完成")

    def upd_qpb(self, qpb, timer):
        if qpb.value() < qpb.maximum():
            qpb.setValue(qpb.value() + 1)
        else:
            timer.stop()

    def single_fliter(self):
        reslen = len(self.fliter_result)
        dl = 1 / reslen * 100
        exact, result, cnt = 0, False, 0
        for word in self.fliter_result:
            cnt += 1
            exact += dl
            self.progressBar_4.setValue(int(exact))
            self.d_fliter_s = Fliter_s(word, result, "{}/{}".format(cnt, reslen))
            self.d_fliter_s.show()
            speaker = wincl.Dispatch("SAPI.SpVoice")
            speaker.rate = 5
            speaker.volume = 100
            speaker.Speak(word)
            while self.d_fliter_s.isVisible():
                sleep(0.01)
                self.gui()
                if self.flag_for_single:
                    self.d_fliter_s.hide()
                    return
            if result:
                self.fliter_result.remove(word)

    def multi_fliter(self):
        reslen = len(self.fliter_result)
        results = []
        self.d_fliter_m = Fliter_m(self.fliter_result, results)
        self.d_fliter_m.show()
        while self.d_fliter_m.isVisible():
            sleep(0.01)
            self.gui()
        cnt = 0
        for result in results:
            if result:
                del self.fliter_result[cnt - reslen]
            cnt += 1

    def closeEvent(self, event):
        os._exit(0)

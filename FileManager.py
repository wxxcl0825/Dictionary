# -*- coding: utf-8 -*-
import winreg
from pickle import load, dump

from Searcher import Searcher


class File:
    def __init__(self):
        self.Searcher = Searcher()

    def readfile(self, path):
        with open(file=path, mode='r', encoding='UTF-8') as f:
            text = f.read()
        return text

    def anafile(self, text, gui=None, is_lemmas=False, qpb=None, title=None):
        word, origin, textlen, dl, exact = "", [], 0, 0, 0
        textlen = len(text)
        dl = 1 / textlen * 100
        for s in text:
            exact += dl
            if qpb is not None:
                qpb.setValue(int(exact))
            if title is not None:
                title.setWindowTitle("分析中{}%".format(int(exact)))
            if gui is not None:
                gui()
            if s.isupper():
                s = s.lower()
            if s < 'a' or s > 'z':
                if word != "":
                    if is_lemmas:
                        word = self.Searcher.lemmas(word)
                    origin.append(word)
                    word = ""
            else:
                word += s
        return origin

    def fliterfile(self, origin, lib):
        s_origin = set(origin)
        return s_origin - lib

    def savefile(self, file, text):
        with open(file=file, mode="w", encoding='UTF-8') as f:
            f.write(text)

    def readsettings(self, file):
        with open(file=file, mode="rb") as f:
            return load(f)

    def savesettings(self, file, settings):
        with open(file=file, mode="wb") as f:
            dump(settings, f)

    def getdesktop(self):
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        return winreg.QueryValueEx(key, "Desktop")[0]

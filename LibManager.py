# -*- coding: utf-8 -*-
from pickle import dump, load


class Lib:
    def __init__(self):
        pass

    def readlib(self, path):
        with open(file=path, mode='rb') as f:
            lib = load(f)
            return lib

    def unionlib(self, path1, path2, savepath):
        with open(file=path1, mode='rb') as f:
            lib1 = load(f)
        with open(file=path2, mode='rb') as f:
            lib2 = load(f)
        self.savelib(lib1 | lib2, savepath)

    def savelib(self, lib, savepath):
        with open(file=savepath, mode='wb') as f:
            dump(lib, f)

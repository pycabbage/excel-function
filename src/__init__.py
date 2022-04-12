#!/usr/bin/env python3
#coding: utf-8

from .util import *


class Loader:
    def __init__(self, prog: str) -> None:
        self.prog = prog
        self.progs = {
            "position": "",

        }

    def load(self):
        progl = self.prog.splitlines()
        

class ExcelFunc:
    def __init__(self) -> None:
        pass

    def load(self, filepath: str) -> None:
        with open(filepath) as fp:
            self.__prog = fp.read()
        self.__load(self.__prog)
    
    def __load(self, prog: str) -> None:
        pass

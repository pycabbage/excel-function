#!/usr/bin/env python3
#coding: utf-8

from sys import argv
from typing import List


def console_wrapper(arg:List[str]=argv) -> None:
    print(arg[1:])


if __name__ == "__main__":
    console_wrapper(argv)

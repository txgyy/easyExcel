#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
from easyExcel.excelTools import groupBookstoOne
if __name__ == "__main__":
    argvs = sys.argv
    path = os.getcwd()
    destination = 'destination.xls'
    if len(argvs)==3:
        from_start = eval(argvs[1])
        from_end = eval(argvs[2])
        groupBookstoOne(path, destination, from_start,from_end)
    elif len(argvs)==2:
        pass

# -*- coding: utf-8 -*-
"""
Created on Fri May 25 10:51:24 2018

@author: 12600771
"""

import openpyxl

typeDict = {"bool":"BOOL", "double":"LREAL", "float": "REAL", "int":"DINT", "unsigned int":"DWORD"}

file = "g_pChannelData.xlsx"
workbook = openpyxl.load_workbook(file)
sheet1 = workbook["Sheet1"]
cell = sheet1.cell(row=3, column=4)

# -*- coding: utf-8 -*-
"""
Created on Wed Aug 18 17:55:13 2021

@author: VR787FC
"""

import camelot,xlwings,pandas as pd
tables = camelot.read_pdf(r'C:\Users\VR787FC\Desktop\SSG.pdf', pages='all',flavor='stream', edge_tol=1000)
# tables
xlapp = xlwings.App(visible=True)
wb = xlapp.books[0]
for i in tables:
    ws = wb.sheets.add()
    ws["A1"].options(pd.DataFrame, header=1, index=True, expand='table').value = i.df

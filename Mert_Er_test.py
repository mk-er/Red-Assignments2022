# -*- coding: utf-8 -*-
"""
Created on Fri Jun  3 22:38:32 2022

@author: mertk
"""

import pandas as pd  
import numpy as np
import IDEALib as ideaLib
from openpyxl import load_workbook
  

"""load data using openpyxl because IDEA doesn't support pandas.read_excel"""
wb = load_workbook(filename=r"C:\Users\mertk\OneDrive\Belgeler\My IDEA Documents\IDEA Projects\test_excel\Source Files.ILB\4- BOBİ FRS Nakit Akış Tablosu - Dolaylı Yöntem (Konsolide).xlsx")
sheet = wb["BOBİ FRS NAT Dolaylı Konsolide"] 
df = pd.DataFrame(sheet.values)

"""Take the relevant columns, i.e containing strings to form AÇIKLAMAR column"""
df = pd.DataFrame(df.values[3:,2:5])

"""In the excel file AÇIKLAMAR column is spread out in 3 different columns.
 Here we combine them into a single column"""
values = df.values

values[pd.isna(values[:,0]),0] = values[pd.isna(values[:,0]),1]
values[pd.isna(values[:,0]),0] = values[pd.isna(values[:,0]),2]

acıklamalar = values[:,0].reshape((-1,1))

"""Create the dataframe to be imported to IDEA"""
table = np.concatenate((acıklamalar, np.zeros((len(acıklamalar), 2))), axis = 1)
columns = ["ACIKLAMALAR", "CARI_DONEM", "ONCEKI_DONEM"]
df = pd.DataFrame(table, columns = columns)

"""Import the dataframe to IDEA"""
ideaLib.py2idea(df, "Mert_Er_test", client=None, createUniqueFile = True)


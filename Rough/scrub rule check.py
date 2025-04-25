import pandas as pd
from openpyxl import Workbook, load_workbook 
import _functions as cfx
import os


initialdir = r"C:\Users\karthikeyan.s\Documents\Sage Servicewise Report\SCRUBRULE - PYTHON\0Source Files"
data_path = cfx.ifile(title='OPDATA.CSV', initialdir=initialdir)
data_df = pd.read_csv(data_path, header=None)
col_value_list = []
col_value_list1 = []
chk = int(cfx.inputbox(title='Check', prompt='1. Column Names\n\n2. Values Columns\n\n3. Both'))

if chk == 1:   
    for col in data_df.columns:
        xx = data_df[col].unique()
        nn = len(xx)
        if nn == 1:
            col_value_list.append(f"{col} - {xx[0]}")

    for col_value in col_value_list:
        print(col_value)

elif chk == 2:
    for col in data_df.columns:
        xx = data_df[col].unique()
        nn = len(xx)
        if nn > 1:
            col_value_list.append(f"{col} - {xx[0]}")

    for col_value in col_value_list[:10]:
        print(col_value)

    print("......................")

    for col_value in col_value_list[-10:]:
        print(col_value)

elif chk == 3:
    for col in data_df.columns:
        xx = data_df[col].unique()
        nn = len(xx)
        if nn == 1:
            col_value_list.append(f"{col} - {xx[0]}")

    for col_value in col_value_list:
        print(col_value)

    print("========================")
    print("========================")

    for col1 in data_df.columns:
        xx1 = data_df[col1].unique()
        nn1 = len(xx1)
        if nn1 > 1:
            col_value_list1.append(f"{col1} - {xx1[0]}")

    for col_value1 in col_value_list1[:10]:
        print(col_value1)

    print("......................")

    for col_value1 in col_value_list1[-10:]:
        print(col_value1)
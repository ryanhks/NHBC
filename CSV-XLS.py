# Import packages
import pandas as pd
import numpy as py
import os
import sys
import csv


# Path to xlrd and openpyxl
sys.path.extend([r"C:/Users/Ryan/Anaconda3/envs/nhbc/lib/site-packages"])

# Import Openpyxl and load workbook

import openpyxl
from openpyxl import load_workbook

# Change directory
os.chdir("C:/Users/Ryan/Desktop/NHBC_Project")

# List all files and directories in current directory
os.listdir('.')

# csv to dataframe
pd.read_csv('Novel_FAC.txt', delimiter='\t')
df1 = pd.read_csv('Novel_FAC.txt', delimiter='\t')
print(df1)
df1.head()

# Designate Genelist as workbook for Openpyxl

book = load_workbook('Novel_GeneList.xlsx')

# Write df to Excel

writer = pd.ExcelWriter('Novel_GeneList.xlsx', engine='openpyxl')
writer.book = book

df1.to_excel(writer, sheet_name='Novel_FAC')
writer.save()

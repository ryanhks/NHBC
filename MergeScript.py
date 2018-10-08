# Import packages
import pandas as pd
import numpy as py
import os
import sys

# Path to xlrd
sys.path.extend([r"C:/Users/Ryan/Anaconda3/envs/nhbc/lib/site-packages"])

# Change directory
os.chdir("C:/Users/Ryan/Desktop/NHBC_Project")

# List all files and directories in current directory
os.listdir('.')

#Import data and concatenate Data Frames
xls = pd.ExcelFile('TCGA_GeneList.xlsx')
df1 = pd.read_excel(xls, 'Sheet1')
df2 = pd.read_excel(xls, 'Sheet2')
df3 = pd.read_excel(xls, 'Sheet3')
df4 = pd.read_excel(xls, 'Sheet4')
df5 = df1[['TCGA_GeneName','Expression_Normal','Expression_Tumor']]
df6 = df2[['UHCC_GeneName']]
df = pd.merge(left=df3, right=df4, on="GeneName", how="left")
df
df.head()

df7 = df1.append(df2)
df8 = df2.append(df1)

# Merge shared genes between TCGA and UHCC
df9 = (pd.merge(df3,df4, on=['GeneName']))

# Output df9 with expression data
df10 = df9[['GeneName','Expression_Normal','Expression_Tumor','Fold-Change(Tumor vs. Normal)']]

# Generate df for to Append for Novel Genes Lists Inputs
df15 = df3[['GeneName']]
df16 = df4[['Transcript ID','GeneName','Fold-Change(Tumor vs. Normal)','RefSeq','p-value(Tumor vs. Normal)']]

# Append and Drop Dupliates to Generate Novel Genes Lists
df11 = df15.append(df16)
df12 = df11.drop_duplicates(keep=False)

df13 = df16.append(df15)
df14 = df13.drop_duplicates(subset=['GeneName'], keep=False)

# Check out info of DataFrame
df.info()


# Print index of df
print(df.index)
print(df3)

# Write df to Excel
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
df10.to_excel(writer, sheet_name='TCGA-UHCC Shared Genes')
df14.to_excel(writer, sheet_name='Novel UHCC Genes')
df12.to_excel(writer, sheet_name='Novel TCGA Genes')
df1.to_excel(writer, sheet_name='Original TCGA Dataset')
df2.to_excel(writer, sheet_name='Original UHCC Dataset')
writer.save()

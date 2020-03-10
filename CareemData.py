import openpyxl
from openpyxl import workbook
import pandas as pd
import numpy as np

df1 = pd.read_excel("testingtheory.xlsx")
total_rows1=len(df1.axes[0])
filtered_df1 = df1[df1['text'].notnull()]
print(filtered_df1)
print(len(filtered_df1))
export_excel = filtered_df1.to_excel (r'D:\AI\SocialPrediction\testingtheory123.xlsx', index = None, header=True) #Don't forget to add '.xlsx' at the end of the path

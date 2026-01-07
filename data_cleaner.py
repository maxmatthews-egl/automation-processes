import numpy as np
import pandas as pd 

#########################Instructions#############################
#Change excel_path, sheet_name, usecols to match excel formatting#
##################################################################

no_sims = 101

excel_path = r"C:\Users\matthewsm\OneDrive - Enstargroup\2025 Own Drive Work\12. Dec\three_lens_test.xlsx"
sheet_name = '1Yr Reserve Risk by Class 5928'
class_cols =  "B:AG"

df = pd.read_excel(excel_path, sheet_name = sheet_name, usecols = class_cols)

classes = (df.columns.tolist())
no_classes = len(classes)
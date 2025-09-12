import pandas as pd

path = r"c:\Users\karthikeyans\Documents\Epicle\Audit File\Audit Raw Files\Audit Master File  - Aug'2025- Week-4 1.xlsx"

df = pd.read_excel(path, sheet_name='Prebilling Data')
dd = df.iloc[15,5]
print(dd)


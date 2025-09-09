import pandas as pd

df = pd.DataFrame({
    'Name': ['Alice', 'Bob'],
    'JoinDate': pd.to_datetime(['2023-09-08', '2023-09-09']),
    'EndDate' : pd.to_datetime(['2023-09-20', '2023-09-21']),
    'EndDate1' : pd.to_datetime(['2023-09-30', '2023-10-21'])
})

output_file = 'rough.xlsx'

with pd.ExcelWriter(output_file, engine='xlsxwriter', datetime_format='m/d/yyyy') as writer:
    df.to_excel(writer, index=False, sheet_name='Sheet1')

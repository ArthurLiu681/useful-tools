import pandas
import os
import datetime

print("begin: " + datetime.datetime.now().strftime("%H:%M:%S") + "\n")
df = pandas.DataFrame()
directory = 'C:\\Users\\0209105\\Desktop\\path'
desdir = 'C:\\Users\\0209105\\Desktop\\despath'
for filename in os.listdir(directory):
    file_path = os.path.join(directory, filename)
    file = pandas.ExcelFile(file_path)
    for sheet in file.sheet_names:
        if len(df.columns) == 0:
            df = pandas.read_excel(file_path, sheet_name=sheet)
        else:
            df = pandas.concat([df, pandas.read_excel(file_path, sheet_name=sheet)])

df.fillna("",inplace=True)
df = df.astype(str)
for company, data in df.groupby("公司"):
    data.to_excel(os.path.join(desdir, company.replace("?"," ")+"-2022收入成本大表.xlsx"), index=False)

print("end: " + datetime.datetime.now().strftime("%H:%M:%S") + "\n")
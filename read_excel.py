Continue: Open config.jsonimport pandas as pd
import openpyxl

# Load the Excel file
file_path = r"c:\Users\User\OneDrive\03=personal(香港)\31 badminton　対戦プログラム\チームリスト.xlsx"

wb = openpyxl.load_workbook(file_path, read_only=True)
sheet = wb.active
data = list(sheet.values)
df = pd.DataFrame(data[1:], columns=data[0])
print("チームリスト.xlsx:")
print(df.head(20))
print("Total teams:", len(df))
print("Level counts:")
print(df['レベル'].value_counts())
import pandas as pd
import openpyxl

# Load the Excel file
file_path = r"c:\Users\User\OneDrive\03=personal(香港)\31 badminton　対戦プログラム\251122　集計結果 final 36th 交流会 - vf.xlsm"

# Read the 'リスト' sheet
wb = openpyxl.load_workbook(file_path, read_only=True)
sheet = wb['リスト']
data = list(sheet.values)
df = pd.DataFrame(data[1:], columns=data[0])
valid_teams = df.dropna(subset=[df.columns[0]])

# Create team list Excel
wb_new = openpyxl.Workbook()
ws = wb_new.active
ws.title = "チームリスト"
ws.append(["ペア名", "氏名", "レベル", "グループ"])

for _, row in valid_teams.iterrows():
    name = row['ペア名 ↓値ばりで記入']
    members = row['氏名　↓値ばりで記入']
    level = [c for c in name if c in 'ABC'][0]  # Correctly extract level
    group = name.rstrip('0123456789')
    ws.append([name, members, level, group])

wb_new.save(r"c:\Users\User\OneDrive\03=personal(香港)\31 badminton　対戦プログラム\チームリスト.xlsx")
print("チームリスト.xlsx created")
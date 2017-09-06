from openpyxl import load_workbook

wb = load_workbook(filename = 'xlsx/BMW_Sales_Standards_2016_ME.xlsx', data_only=True)
sheets = wb.sheetnames[11:12]

for sheet in sheets:
  print sheet
  ws = wb[sheet]
  cell = ws['B6']
  for column in ws.columns:
    print column[5].value
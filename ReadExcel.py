from openpyxl import load_workbook

wb = load_workbook(filename = 'xlsx/libro.xlsx', data_only=True)
sheets = wb.sheetnames[0:1]

for sheet in sheets:
  ws = wb[sheet]
  cell = ws['B2']
  for column in ws.columns:
    print column[1].value
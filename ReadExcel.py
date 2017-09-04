from openpyxl import load_workbook
wb = load_workbook(filename = 'libro.xlsx', read_only=True)

print wb.sheetnames

for sheet in wb.sheetnames:
  ws = wb[sheet]
  print ws['A1'].value
  
# ws = wb['Datos']

# for row in ws.rows:
#   for cell in row:
#     print cell.value
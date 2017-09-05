from openpyxl import load_workbook
from openpyxl.worksheet import Worksheet

wb = load_workbook(filename = 'xlsx/BMW_Sales_Standards_2016_ME.xlsx', data_only=True)

sheets = wb.sheetnames[3:13]
count = 100
sizeCell = []



for sheet in sheets:
  ws = wb[sheet]
  cell_range = ws['X6':'X200']
  data = ws['X6':'X200']
  for i, rowOfCellObjects in enumerate(cell_range):
    for n, cellObj in enumerate(rowOfCellObjects):
      if cellObj.value is not None:
        sizeCell.append(cellObj.value)
        print sizeCell
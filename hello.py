from openpyxl import Workbook
wb = Workbook()

ws = wb.active

ws['A1'] = 42
ws['A2'] = 88
ws.append([1, 2, 3])

# Save the file
wb.save("sample.xlsx")
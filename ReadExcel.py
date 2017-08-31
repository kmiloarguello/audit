from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = 'Datos'

ws['A1'] = 'ID'
ws['B1'] = 'Nombre'
ws['C1'] = 'Apellido'

for i in range(2,5):
  for j in range(2,5):
    ws['A' + str(i)] = i
    ws['B' + str(i)] = 'Camilo'
    ws['C' + str(i)] = 'Arguello'
    
wb.save('libro.xlsx')
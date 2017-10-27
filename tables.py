import openpyxl


xfile = openpyxl.load_workbook('xlsx/libro.xlsx')

sheet = xfile.get_sheet_by_name('Datos')
sheet['A1'] = 'hello world'
xfile.save('xlsx/libroB.xlsx')
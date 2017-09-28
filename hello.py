# Tkinter lib to create user interface
import sys
from Tkinter import *
from tkFileDialog import askopenfilename
from tkintertable import TableCanvas, TableModel

# Openpyxl libs
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.writer.write_only import WriteOnlyCell

window = Tk()
window.title('K@PTA Excel Auditorias')
window.wm_iconbitmap('img/kapta_mex.ico')
window.geometry('{}x{}'.format(800, 600))
# window.resizable(0,0)


window.filename = askopenfilename( filetypes = (("Archivos de Auditorias", ".xlsx"), ("Todos los archivos", "*.*")))

numberCategory = []
zerovalue = []
index_number_categories = []
rowData = []
cleanedRowData = []
auditvalue = []
essential = []

wb = load_workbook(filename = window.filename, data_only=True)
sheets = wb.sheetnames[3:12]

archivo = Label(window, text='En el archivo ' + window.filename)
archivo.grid(row=1, column=1)
archivo.pack()

for sheet in sheets:
  ws = wb[sheet]

  hoja = Label(window, text='En la hoja ' + sheet)
  hoja.configure(foreground="red")
  hoja.pack()

  for row in ws.rows: 
    numberCategory.insert(0,row[23].value)  
    number_categories_without_filter = next(i for i in numberCategory if i is not None)
    index_number_categories.extend([number_categories_without_filter])

    zerovalue.insert(0,row[25].value)  
    zero_categories_without_filter = next(i for i in zerovalue if i is not None)

    auditvalue.insert(0,row[13].value)  
    audit_categories_without_filter = next(i for i in auditvalue if i is not None)

    essential.insert(0,row[15].value)  
    essential_without_filter = next(i for i in essential if i is not None)

    if(row[23].value == "N" and zero_categories_without_filter == 0 and audit_categories_without_filter == 'Audit' ):

      negativos = Label(window, text='Valor ' + str(row[23].value) +  ' en la celda ' +  str(row[23]),font=("Helvetica", 13))
      negativos.configure(foreground="blue")
      negativos.pack()

      result = Label(window, text='Valor ' + str(zero_categories_without_filter) + ' en la celda ' +  str(row[25]),font=("Helvetica", 10))
      result.pack()
      
      audit = Label(window, text='Valor ' + str(audit_categories_without_filter) + ' en la celda ' +  str(row[13]),font=("Helvetica", 9))
      audit.pack()

      # Essential,Contract, Optional
      essentials = Label(window, text='Valor ' + str(essential_without_filter) + ' en la celda ' +  str(row[15]),font=("Helvetica", 8))
      essentials.pack()

window.mainloop()

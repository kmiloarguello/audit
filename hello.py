# Tkinter lib to create user interface
import sys
from Tkinter import *
from tkFileDialog import askopenfilename
from tkintertable import TableCanvas, TableModel
import tkMessageBox


# Openpyxl libs
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.writer.write_only import WriteOnlyCell


# Functions

def myFunction():
  ne = Label(window, text='Hola',bg="red")
  ne.pack()

def editMenu():
  print 'Again'

def helpMenu():
  print 'Help'

def openExcel():
  print 'hola'

def acercaDe():
  myWindow = Toplevel(window)
  myWindow.title('K@PTA')
  myWindow.wm_iconbitmap('img/kapta_mex.ico')
  myWindow.geometry('200x100')
  acercaDeContent = Label(myWindow, text='Derechos Reservados K@PTA')
  acercaDeContent.pack()

def exitApp():
  exited = tkMessageBox.askyesno('Salir','Esta seguro?')
  if(exited == True):
    window.destroy()

# Initialization

window = Tk()
window.title('K@PTA Excel Auditorias')
window.wm_iconbitmap('img/kapta_mex.ico')
window.geometry('800x600')
window.configure(background='white')

# Menu

menu = Menu(window)
window.config(menu=menu)

subMenu = Menu(menu,tearoff=0,bg='white')
menu.add_cascade(label='Archivo', menu=subMenu)
subMenu.add_command(label='Nuevo proyecto', command=myFunction)
subMenu.add_command(label='Abrir Excel', command=openExcel)
subMenu.add_command(label='Guardar', command=myFunction)
subMenu.add_command(label='Exportar Excel', command=myFunction)
subMenu.add_separator()
subMenu.add_command(label='Acerca de K@PTA', command=acercaDe)
subMenu.add_command(label='Salir', command=exitApp)

editMenu = Menu(menu,tearoff=0,bg='white')
menu.add_cascade(label='Editar', menu=editMenu)
editMenu.add_command(label='Deshacer', command=myFunction)

helpMenu = Menu(menu,tearoff=0,bg='white')
menu.add_cascade(label='Ayuda', menu=helpMenu)

# Bottom
toolbar = Frame(window,bg='white')
myLabel = Label(toolbar, text='Derechos Reservados K@PTA', bg='white')
myLabel.pack(side=RIGHT)
toolbar.pack(side=BOTTOM, fill=X)
  
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


# myRows = []

# for ii in range(5):
#   myCols = []
#   for jj in range(4):
#     e = Entry(relief=RIDGE)
#     e.grid(row=ii,column=jj,sticky=NSEW, ipadx=30)
#     e.insert(END, 2)
#     myCols.append(e)
#   myRows.append(myCols)

auditInfo = {'rec1': {'col1': 99.88, 'col2': 108.79, 'label': 'rec1'},}

tframe = Frame(window)
tframe.pack()
model = TableModel()
table = TableCanvas(tframe,model=model,editable=False)
table.createTableFrame()

model = table.model
model.importDict(auditInfo)
table.redrawTable()


for sheet in sheets:
  ws = wb[sheet]
  hoja = Label(window, text='En la hoja ' + sheet,bg='white')
  hoja.configure(foreground="red")
  hoja.pack()

#   for row in ws.rows: 
#     numberCategory.insert(0,row[23].value)  
#     number_categories_without_filter = next(i for i in numberCategory if i is not None)
#     index_number_categories.extend([number_categories_without_filter])

#     zerovalue.insert(0,row[25].value)  
#     zero_categories_without_filter = next(i for i in zerovalue if i is not None)

#     auditvalue.insert(0,row[13].value)  
#     audit_categories_without_filter = next(i for i in auditvalue if i is not None)

#     essential.insert(0,row[15].value)  
#     essential_without_filter = next(i for i in essential if i is not None)

#     if(row[23].value == "N" and zero_categories_without_filter == 0 and audit_categories_without_filter == 'Audit' ):

#       negativos = Label(window, text='Valor ' + str(row[23].value) +  ' en la celda ' +  str(row[23]),font=("Helvetica", 13),bg='white')
#       negativos.configure(foreground="blue")
#       negativos.pack()

#       result = Label(window, text='Valor ' + str(zero_categories_without_filter) + ' en la celda ' +  str(row[25]),font=("Helvetica", 10),bg='white')
#       result.pack()
      
#       audit = Label(window, text='Valor ' + str(audit_categories_without_filter) + ' en la celda ' +  str(row[13]),font=("Helvetica", 9),bg='white')
#       audit.pack()

#       # Essential,Contract, Optional
#       essentials = Label(window, text='Valor ' + str(essential_without_filter) + ' en la celda ' +  str(row[15]),font=("Helvetica", 8),bg='white')
#       essentials.pack()

window.mainloop()

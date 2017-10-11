# Tkinter lib to create user interface
import sys
from Tkinter import *
from tkFileDialog import askopenfilename
from tkintertable import TableCanvas, TableModel
import tkMessageBox
from PIL import Image, ImageTk

# Openpyxl libs
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.writer.write_only import WriteOnlyCell

# Functions
def newProject():
  openExcel()

def myFunction():
  print 'hi'

def editMenu():
  print 'Again'

def helpMenu():
  print 'Help'

def openExcel():
  excelSection.delete('all')


  filename = askopenfilename( filetypes = (("EXCEL", ".xlsx"), ("Todos los archivos", "*.*")))

  numberCategory = []
  zerovalue = []
  index_number_categories = []
  rowData = []
  cleanedRowData = []
  auditvalue = []
  essential = []
  standard = []
  number = []
  requirement = []
  comments = []
  question = []
  observation = []
  suggested = []

  wb = load_workbook(filename = filename, data_only=True)

  sheets = wb.sheetnames[3:12]

  myHoja = []
  myStandard = []
  myNumber = []
  myRequeriment = []
  myComment = []
  myAudit = []
  myEssentials = []
  myAuditQuestion = []
  myObservation = []
  mySuggested = []
  myN = []
  myZero = []
  myAComments = []
  myPic = []
  auditcomments = []
  picture = []

  for sheet in sheets:
    ws = wb[sheet]

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

      standard.insert(0,row[0].value)  
      standard_categories_without_filter = next(i for i in standard if i is not None)

      number.insert(0,row[1].value)  
      number_categories_without_filter = next(i for i in number if i is not None)

      requirement.insert(0,row[2].value)  
      requirement_categories_without_filter = next(i for i in requirement if i is not None)

      comments.insert(0,row[4].value)  
      comments_categories_without_filter = next(i for i in comments if i is not None)

      question.insert(0,row[17].value)  
      question_categories_without_filter = next(i for i in question if i is not None)

      observation.insert(0,row[19].value)  
      observation_categories_without_filter = next(i for i in observation if i is not None)

      suggested.insert(0,row[21].value)  
      suggested_categories_without_filter = next(i for i in suggested if i is not None)

      auditcomments.insert(0,row[30].value)  
      auditcomments_categories_without_filter = next(i for i in auditcomments if i is not None)

      picture.insert(0,row[30].value)  
      picture_categories_without_filter = next(i for i in picture if i is not None)


      if(row[23].value == "N" and zero_categories_without_filter == 0 and audit_categories_without_filter == 'Audit' or audit_categories_without_filter == 'Audit / Regional Office' ):
        myHoja.extend([sheet])
        myN.extend([str(row[23].value)])
        myZero.extend([str(zero_categories_without_filter)])
        myAudit.extend([str(audit_categories_without_filter)])
        myEssentials.extend([str(essential_without_filter)])
        
        myStandard.extend([str(standard_categories_without_filter)])
        myNumber.extend([str(number_categories_without_filter)])
        myRequeriment.extend([requirement_categories_without_filter.encode('utf-8')])
        myComment.extend([str(comments_categories_without_filter)])
        myAuditQuestion.extend([str(question_categories_without_filter)])
        myObservation.extend([str(observation_categories_without_filter)])
        mySuggested.extend([str(suggested_categories_without_filter)])
        myAComments.extend([str(auditcomments_categories_without_filter)])
        myPic.extend([str(picture_categories_without_filter)])


  tframe = Frame(window)
  tframe.pack()
  model = TableModel()
  table = TableCanvas(tframe,model=model,editable=False,rowheaderwidth=50)
  table.createTableFrame()
  model = table.model

  dict = {}

  for i in range(len(myHoja)):
    dict[i] = {'ID': i}

  model.importDict(dict)

  table.addColumn('Hoja Excel')
  table.addColumn('Standard')
  table.addColumn('Number')
  table.addColumn('Requirement 2015')
  table.addColumn('Comments')
  table.addColumn('Type of Check')
  table.addColumn('Essentials')
  table.addColumn('Audit Question')
  table.addColumn('Observation / Evidence Required / Audit Remarks')
  table.addColumn('Suggested Person to ask')
  table.addColumn('Evaluation(0/1')
  table.addColumn('Result')
  table.addColumn('Audit Comments')
  table.addColumn('Picture / Statement / Proof')


  for i in range(len(myHoja)):
    table.model.data[i]['Hoja Excel'] = myHoja[i]
    table.model.data[i]['Standard'] = myStandard[i]
    table.model.data[i]['Number'] = myNumber[i]
    table.model.data[i]['Requirement 2015'] = myRequeriment[i]
    table.model.data[i]['Comments'] = myComment[i]
    table.model.data[i]['Type of Check'] = myAudit[i]
    table.model.data[i]['Essentials'] = myEssentials[i]
    table.model.data[i]['Audit Question'] = myAuditQuestion[i]
    table.model.data[i]['Observation / Evidence Required / Audit Remarks'] = myObservation[i]
    table.model.data[i]['Suggested Person to ask'] = mySuggested[i]
    table.model.data[i]['Evaluation (0/1)'] = myN[i]
    table.model.data[i]['Result'] = myZero[i]
    table.model.data[i]['Audit Comments'] = myAComments[i]
    table.model.data[i]['Picture / Statement / Proof'] = myPic[i]


  table.redrawTable()

  return tframe

def openImage():
  filename = askopenfilename( filetypes = (("Imagen de resultado", ".jpg"), ("Todos los archivos", "*.*")))
  toolbar2 = Frame(window,bg='red')
  myLabel2 = Label(toolbar2, text=filename, bg='red')
  myLabel2.pack(side=LEFT)
  toolbar2.pack(side=TOP, fill=X)
  return filename

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

# scrollbar = Scrollbar(window)
# scrollbar.pack(side=RIGHT, fill=Y)

# listbox = Listbox(window, yscrollcommand=scrollbar.set)
# for i in range(1000):
#   listbox.insert(END,str(i))

# listbox.pack(side=LEFT, fill=BOTH)

# scrollbar.config(command=listbox.yview)

# Menu

menu = Menu(window)
window.config(menu=menu)

subMenu = Menu(menu,tearoff=0,bg='white')
menu.add_cascade(label='Archivo', menu=subMenu)
subMenu.add_command(label='Nuevo proyecto', command=newProject)
subMenu.add_command(label='Abrir excel', command=openExcel)
subMenu.add_command(label='Abrir imagen', command=openImage)
subMenu.add_command(label='Guardar', command=myFunction)
subMenu.add_command(label='Exportar excel', command=myFunction)
subMenu.add_separator()
subMenu.add_command(label='Acerca de K@PTA', command=acercaDe)
subMenu.add_command(label='Salir', command=exitApp)

editMenu = Menu(menu,tearoff=0,bg='white')
menu.add_cascade(label='Editar', menu=editMenu)
editMenu.add_command(label='Deshacer', command=myFunction)

helpMenu = Menu(menu,tearoff=0,bg='white')
menu.add_cascade(label='Ayuda', menu=helpMenu)

excelSection = Canvas(window, width=800,height=600, bg='red')
excelSection.pack(fill=BOTH,expand=YES)

print excelSection

image = Image.open('img/bg.png')
photo = ImageTk.PhotoImage(image)
imgFile = excelSection.create_image(400,300, image=photo)

# Bottom
toolbar = Frame(window,bg='white')
myLabel = Label(toolbar, text='Derechos reservados K@PTA', bg='white')
myLabel.pack(side=RIGHT)
toolbar.pack(side=BOTTOM, fill=X)

window.mainloop()

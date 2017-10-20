# Tkinter lib to create user interface
from Tkinter import *
from tkFileDialog import askopenfilename
from tkintertable.Tables import TableCanvas
from tkintertable.TableModels import TableModel
import tkMessageBox
from PIL import Image, ImageTk

# Openpyxl libs
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.writer.write_only import WriteOnlyCell

entry_txt = StringVar

# # Functions
def newProject():
  openExcel()

def myFunction():
  print 'hi'

def editMenu():
  print 'Again'

def helpMenu():
  print 'Help'

def openExcel():
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

      final_audit = audit_categories_without_filter.encode('ascii','ignore')

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



      if(row[23].value == "N" and zero_categories_without_filter == 0 and 'Audit' in final_audit ):
        myHoja.extend([sheet])
        myN.extend([str(row[23].value)])
        myZero.extend([str(zero_categories_without_filter)])
        myAudit.extend([audit_categories_without_filter.encode('ascii', 'ignore').decode('ascii')])
        myEssentials.extend([str(essential_without_filter)])
        
        myStandard.extend([str(standard_categories_without_filter)])
        myNumber.extend([str(number_categories_without_filter)])
        myRequeriment.extend([requirement_categories_without_filter.encode('utf-8')])
        myComment.extend([comments_categories_without_filter.encode('ascii', 'ignore').decode('ascii')])
        myAuditQuestion.extend([question_categories_without_filter.encode('ascii', 'ignore').decode('ascii')])
        myObservation.extend([observation_categories_without_filter.encode('ascii', 'ignore').decode('ascii')])
        mySuggested.extend([suggested_categories_without_filter.encode('ascii', 'ignore').decode('ascii')])
        myAComments.extend([auditcomments_categories_without_filter.encode('ascii', 'ignore').decode('ascii')])
        myPic.extend([picture_categories_without_filter.encode('ascii', 'ignore').decode('ascii')])

  tframe = Frame(window)
  tframe.pack(fill=X)
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
  table.addColumn('Evaluation (0/1)')
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
  imgLayout(filename)
  return filename

def returnEntry(parent,entry):
  e = Entry(parent)
  e.insert(10,entry)

def imgLayout(filename):
  img_container = Frame(window)
  Label(img_container, text="  ").grid(row=1, column=0)
  Label(img_container, text="  ").grid(row=2, column=0)

  Label(img_container, text='Ruta de archivo origen', font=("Helvetica", 10), foreground='#303133').grid(row=1, column=1)
  Label(img_container, text=filename, font=("Helvetica", 8), foreground='#000').grid(row=2, column=1)

  Label(img_container, text='Ruta de archivo para guardar', font=("Helvetica", 10), foreground='#303133').grid(row=1, column=2)
  entry = Entry(img_container,textvariable=entry_txt)
  entry.grid(row=2, column=2)

  Label(img_container, text='Selecciona Hoja', font=("Helvetica", 10), foreground='#303133').grid(row=1, column=3)
  ex_sh_sel = StringVar(img_container)
  ex_sh_sel.set('one')
  option = OptionMenu(img_container, ex_sh_sel, 'one', 'two', 'three', 'four').grid(row=2, column=3)


  Label(img_container, text='Seleccione Item', font=("Helvetica", 10), foreground='#303133').grid(row=1, column=4)
  ex_sh_sel = StringVar(img_container)
  ex_sh_sel.set('one')
  option = OptionMenu(img_container, ex_sh_sel, 'one', 'two', 'three', 'four').grid(row=2, column=4)

  Label(img_container, text="   ").grid(row=1, column=5)
  Label(img_container, text="   ").grid(row=2, column=5)

  Button(img_container, text="Guardar en Excel", command="").grid(row=2, column=6)


  img_container.pack(side=TOP, fill=X)

  return img_container

def acercaDe():
  myWindow = Toplevel(window)
  myWindow.title('K@PTA')
  myWindow.wm_iconbitmap('kapta_mex.ico')
  myWindow.geometry('200x100')
  acercaDeContent = Label(myWindow, text='Derechos Reservados K@PTA')
  acercaDeContent.pack()

def exitApp():
  exited = tkMessageBox.askyesno('Salir','Esta seguro?')
  if(exited == True):
    window.destroy()


# # Initialization

window = Tk()
window.title('K@PTA Excel Auditorias')
window.wm_iconbitmap('kapta_mex.ico')
window.geometry('800x600')
window.configure(background='white')

# Menu
menu = Menu(window)
window.config(menu=menu)

subMenu = Menu(menu,tearoff=0,bg='white')
menu.add_cascade(label='Archivo', menu=subMenu)
subMenu.add_command(label='Nuevo proyecto', command=newProject)
subMenu.add_command(label='Abrir excel', command=openExcel)
subMenu.add_command(label='Abrir imagen', command=openImage)
subMenu.add_command(label='Exportar excel', command=myFunction)
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
myLabel = Label(toolbar, text='Derechos reservados K@PTA', bg='white')
myLabel.pack(side=RIGHT)
toolbar.pack(side=BOTTOM, fill=X)



window.mainloop()

from Tkinter import *
from PIL import Image, ImageTk
from tkFileDialog import askopenfilename, asksaveasfilename
import tkMessageBox
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.writer.write_only import WriteOnlyCell
import types
from pandastable import Table, TableModel
 
myColor = 'white'

class OtherFrame(Toplevel):
    """"""
    #----------------------------------------------------------------------
    def __init__(self, original):
      """Constructor"""
      self.original_frame = original
      Toplevel.__init__(self)
      self.wm_iconbitmap('kapta_mex.ico')

      myW = 800
      myH = 400

      myWs = root.winfo_screenwidth()
      myHs = root.winfo_screenheight()

      myX = (myWs/2) - (myW/2)
      myY = (myHs/2) - (myH/2)

      self.geometry('%dx%d+%d+%d' % (myW, myH, myX, myY))

      # self.geometry("800x400")
      self.title("K@PTA Auditorias")
      self.configure(background=myColor)

      self.loadMenu()
      self.loadExcel()

      self.frame = Frame(self)
      self.frame.pack(fill=X,expand=1)
      self.renderExcel()

      self.frame2 = Frame(self)
      btn = Button(self.frame2, text="Close", command=self.onClose)
      btn.pack()
      self.frame2.pack()

      self.toolbar = Frame(self,bg='white')
      self.myLabel = Label(self.toolbar, text='Derechos reservados K@PTA', bg='white')
      self.myLabel.pack(side=RIGHT)
      self.toolbar.pack(side=BOTTOM, fill=X)

    #----------------------------------------------------------------------
    def loadMenu(self):
      self.menu = Menu(self)
      self.config(menu=self.menu)

      self.subMenu = Menu(self.menu,tearoff=0,bg='white')
      self.menu.add_cascade(label='Archivo', menu=self.subMenu)
      self.subMenu.add_command(label='Abrir excel', command="openExcel")
      self.subMenu.add_command(label='Abrir imagen', command="openImage")
      self.subMenu.add_command(label='Exportar excel', command="")
      self.subMenu.add_separator()
      self.subMenu.add_command(label='Acerca de K@PTA', command=self.acercaDe)
      self.subMenu.add_command(label='Salir', command="exitApp")

      self.editMenu = Menu(self.menu,tearoff=0,bg='white')
      self.menu.add_cascade(label='Editar', menu=self.editMenu)
      self.editMenu.add_command(label='Deshacer', command="myFunction")

      self.helpMenu = Menu(self.menu,tearoff=0,bg='white')
      self.menu.add_cascade(label='Ayuda', menu=self.helpMenu)

    #----------------------------------------------------------------------
    def arraysInit(self):
      self.numberCategory = []
      self.zerovalue = []
      self.index_number_categories = []
      self.rowData = []
      self.cleanedRowData = []
      self.auditvalue = []
      self.essential = []
      self.standard = []
      self.number = []
      self.requirement = []
      self.comments = []
      self.question = []
      self.observation = []
      self.suggested = []
      self.myHoja = []
      self.myStandard = []
      self.myNumber = []
      self.myRequeriment = []
      self.myComment = []
      self.myAudit = []
      self.myEssentials = []
      self.myAuditQuestion = []
      self.myObservation = []
      self.mySuggested = []
      self.myN = []
      self.myZero = []
      self.myAComments = []
      self.myPic = []
      self.auditcomments = []
      self.picture = []        
    
    def withoutFilter(self,item,cell):
      """
        The NumberTypes variable stores all numeric data type from types library
        The main key is that if the type of variable is not numeric, encodes it into a  ascii
      """
      NumberTypes = (types.IntType, types.LongType, types.FloatType, types.ComplexType)
      item.insert(0,cell.value)  
      without_filter = next(i for i in item if i is not None)
      if not isinstance(without_filter, NumberTypes):
        return without_filter.encode('ascii', 'ignore')
      else:
        return without_filter

    def loadWorkbook(self):
      """
       Excel Funcionality to make an average of N and Audit Values and return it
       First Load the excel file
      """
      self.wb = load_workbook(filename=self.filename,data_only=True)
      self.sheets = self.wb.sheetnames[3:12]

      for sheet in self.sheets:
        self.ws = self.wb[sheet]
        for row in self.ws.rows:
          final_numberCat = self.withoutFilter(self.numberCategory,row[23])
          final_zero = self.withoutFilter(self.zerovalue,row[25])
          final_audit = self.withoutFilter(self.auditvalue,row[13])
          final_essential = self.withoutFilter(self.essential,row[15])
          final_standard = self.withoutFilter(self.standard,row[0])
          final_number = self.withoutFilter(self.number,row[1])
          final_requirement = self.withoutFilter(self.requirement,row[2])
          final_comments = self.withoutFilter(self.comments,row[4])
          final_question = self.withoutFilter(self.question,row[17])
          final_observation = self.withoutFilter(self.observation,row[19])
          final_suggested = self.withoutFilter(self.suggested,row[21])
          final_auditcomments = self.withoutFilter(self.auditcomments,row[30])
          final_picture = self.withoutFilter(self.picture,row[30])

          
          if(row[23].value == "N" and final_zero == 0 and 'Audit' in final_audit ):
            self.myHoja.extend([sheet])
            self.myN.extend([row[23]])
            self.myZero.extend([final_zero])
            self.myAudit.extend([final_audit])
            self.myEssentials.extend([final_essential])
            self.myStandard.extend([final_standard])
            self.myNumber.extend([final_number])
            self.myRequeriment.extend([final_requirement])
            self.myComment.extend([final_comments])
            self.myAuditQuestion.extend([final_question])
            self.myObservation.extend([final_observation])
            self.mySuggested.extend([final_suggested])
            self.myAComments.extend([final_auditcomments])
            self.myPic.extend([final_picture])
        
    def loadExcel(self):
      self.filename = askopenfilename( initialdir = "/KAPTA Camilo/python/xlsx",title = "Subir archivo de excel", filetypes = (("Excel Auditorias", ".xlsx"), ("Todos los archivos", "*.*")))  
      self.arraysInit()
      self.loadWorkbook()

    def renderExcel(self):
      self.table = Table(self.frame)
      self.table.show()

    #----------------------------------------------------------------------
    def acercaDe(self):
      self.about = Toplevel(self)
      self.about.title('K@PTA')
      self.about.wm_iconbitmap('kapta_mex.ico')
      self.about.geometry('200x100')
      Label(self.about, text='Derechos Reservados K@PTA \n Desarrollador Camilo Arguello \n Farrell, D 2016 DataExplore: An Application for General Data Analysis in Research and Education. Journal of Open Research Software, 4: e9, DOI: http://dx.doi.org/10.5334/jors.94', font=("Segoe UI", 9), justify=LEFT).pack()
    #----------------------------------------------------------------------
    def onClose(self):
      """"""
      self.destroy()
      self.quit()
      root.destroy()
      root.quit()
      exit()
      # self.original_frame.show()
 
########################################################################
class MyApp(object):
    """ Initial PAGE """
 
    #----------------------------------------------------------------------
    def __init__(self, parent):
        """Constructor"""
        self.root = parent

        myW = 600
        myH = 200

        myWs = root.winfo_screenwidth()
        myHs = root.winfo_screenheight()

        myX = (myWs/2) - (myW/2)
        myY = (myHs/2) - (myH/2)


        self.root.wm_iconbitmap('kapta_mex.ico')
        self.root.geometry('%dx%d+%d+%d' % (myW, myH, myX, myY))
        self.root.title("K@PTA Excel Auditorias")
        self.root.resizable(0,0)
        self.frame = Frame(parent)

        self.label = Label(self.frame, text='Para iniciar por favor carga el archivo de auditorias')
        self.button = Button(self.frame, justify=LEFT,command=self.openFrame)
        self.photo = ImageTk.PhotoImage(file='img/subir_excel.png')
        self.button.config(image=self.photo, width='300',height='50')
        self.label.pack(fill=X)
        self.button.pack()

        self.frame.pack(expand=1)
 
    #----------------------------------------------------------------------
    def hide(self):
        """"""
        self.root.withdraw()
 
    #----------------------------------------------------------------------
    def openFrame(self):
        """"""
        self.hide()
        subFrame = OtherFrame(self)
 
    #----------------------------------------------------------------------
    def show(self):
        """"""
        self.root.update()
        self.root.deiconify()
 

#----------------------------------------------------------------------
if __name__ == "__main__":
  root = Tk()
  app = MyApp(root)
  root.mainloop()
  exit()
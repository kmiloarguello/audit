# Tkinter lib to create user interface
import sys
from Tkinter import *
from tkFileDialog import askopenfilename

# Openpyxl libs
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.writer.write_only import WriteOnlyCell


class App(Frame):
  def __init__(self, master=None):
    Frame.__init__(self,master)
    self.pack()
    self.createWidgets()

  def createWidgets(self):
    self.QUIT = Button(self)
    self.QUIT['text'] = 'Salir'
    self.QUIT['fg'] = 'red'
    self.QUIT['command'] = self.quit
    self.QUIT.pack(side=LEFT, padx=10,pady=10)
  

window = Tk()
window.title('K@PTA Excel Auditorias')
window.wm_iconbitmap('img/kapta_mex.ico')
window.geometry('{}x{}'.format(800, 600))
window.resizable(0,0)

window.filename = askopenfilename( filetypes = (("Archivos de Auditorias", ".xlsx"), ("Todos los archivos", "*.*")))

print window.filename

wb = load_workbook(filename = window.filename, data_only=True)
sheets = wb.sheetnames[3:4] #Current Sheet

for sheet in sheets:
  print sheet

app = App(master=window)

app.mainloop()
window.destroy()
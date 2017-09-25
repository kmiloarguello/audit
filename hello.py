import sys
from Tkinter import *
from tkFileDialog import askopenfilename

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

window.fileName = filedialog.askopenfilename( filetypes = (("howCode files", ".hc"), ("All files", "*.*")))


# window.configure(background="#ffffff")

# bg = PhotoImage(file="img/bg.gif")
# img = Label(window, image=bg)
# img.place(x=0,y=0, relwidth=1,relheight=1)
# img.pack()

app = App(master=window)

app.mainloop()
window.destroy()
from Tkinter import *

master = Tk()
entryb1 = StringVar()

def print_content():
  global entryb1
  content = entryb1.get()
  print content

Label(master, text="Input: ").grid(row=0, sticky=W)

entry = Entry(master, textvariable=entryb1)
entry.grid(row=1, column=1)

b1 = Button(master, text="continue", command=print_content)
b1.grid(row=2, column=1)

master.mainloop()
from tkinter import *

root = Tk()

e = Entry(root, width=50)

def clik():
    hello ="Hello " + e.get()
    mylab = Label(root,text=hello)
    mylab.grid(row = 0,column=1)


myButton1 = Button(root, text="Button1")
myButton2 = Button(root, text="Button2",command=clik)

e.grid(row=0,column=0)
myButton1.grid(row = 1,column=0)
myButton2.grid(row = 1,column=1)


root.mainloop()

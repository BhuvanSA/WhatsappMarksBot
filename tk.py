from tkinter import *
from tkinter import filedialog


filepath = "empty string of some loong length"

window = Tk()
button = Button(text="Open",command=openFile)
button.pack()
window.mainloop()

print(filepath)


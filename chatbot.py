from tkinter import *
master = Tk()
master.geometry("400x200")


import openpyxl      
path="./GymExercisesDataset.xlsx"
wb_obj=openpyxl.load_workbook(path)

sheet_obj=wb_obj.active

cell_obj=sheet_obj.cell(row=1, column=1)
out=cell_obj.value

questionLabel=Label(master, text="Ask question: ")

question=StringVar()
e1=Entry(master, textvariable=question)

out=StringVar()
out.set("hello")
w=Label(master, textvariable=out)

def submitValue():
    out.set(e1.get())

submitButton=Button(master, text="Submit", fg='green', command=submitValue)

questionLabel.grid(row=0, column=0)
e1.grid(row=0, column=1)
submitButton.grid(row=2, column=1)
w.grid(row=5, column=1)

# Label(master, text=out).grid(row=1)
mainloop()
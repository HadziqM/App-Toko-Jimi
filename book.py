import openpyxl as xl
import tkinter as tk

#EXEL SETUP
workbook = xl.load_workbook(filename="shipment.xlsx")

sheet = workbook.active
name=[]
unit=[]
price=[]
mymoney = [0]

def load():
    name.clear()
    unit.clear()
    price.clear()
    mymoney[0]=sheet["C7"].value
    for i in range(4):
        z = "A" + str(i+2)
        x = "B" + str(i+2)
        y = "C" + str(i+2)
        unit.append(sheet[x].value)
        price.append(sheet[y].value)
        name.append(sheet[z].value)

def save():
    workbook.save(filename="shipment.xlsx")

load() #read spreadsheat

#TKINTER SETUP

root = tk.Tk()
canvas1 = tk.Canvas(root,width = 800,height = 500)
canvas1.pack()

#button command
def one ():
    unit[0] = unit[0]-1
    mymoney[0] += price[0]
    sheet["B2"] = unit [0]
    sheet["C7"] = mymoney[0]
    save()
    label1 = tk.Label(root,text=str(unit[0]) , fg='green',font=('helvetica',12,'bold'))
    canvas1.create_window(200,50,window=label1)
    label2 = tk.Label(root,text=str(mymoney[0]) , fg='blue',font=('helvetica',12,'bold'))
    canvas1.create_window(200,300,window=label2)
def two ():
    unit[1] = unit[1]-1
    mymoney[0] += price[1]
    sheet["B3"] = unit [1]
    sheet["C7"] = mymoney[0]
    save()
    label1 = tk.Label(root,text=str(unit[1]) , fg='green',font=('helvetica',12,'bold'))
    canvas1.create_window(200,100,window=label1)
    label2 = tk.Label(root,text=str(mymoney[0]) , fg='blue',font=('helvetica',12,'bold'))
    canvas1.create_window(200,300,window=label2)
def three ():
    unit[2] = unit[2]-1
    mymoney[0] += price[2]
    sheet["B4"] = unit [2]
    sheet["C7"] = mymoney[0]
    save()
    label1 = tk.Label(root,text=str(unit[2]) , fg='green',font=('helvetica',12,'bold'))
    canvas1.create_window(200,150,window=label1)
    label2 = tk.Label(root,text=str(mymoney[0]) , fg='blue',font=('helvetica',12,'bold'))
    canvas1.create_window(200,300,window=label2)
def four ():
    unit[3] = unit[3]-1
    mymoney[0] += price[3]
    sheet["B5"] = unit [3]
    sheet["C7"] = mymoney[0]
    save()
    label1 = tk.Label(root,text=str(unit[3]) , fg='green',font=('helvetica',12,'bold'))
    canvas1.create_window(200,200,window=label1)
    label2 = tk.Label(root,text=str(mymoney[0]) , fg='blue',font=('helvetica',12,'bold'))
    canvas1.create_window(200,300,window=label2)

#unit name
label1 = tk.Label(root,text=name[0], fg='green',font=('helvetica',12,'bold'))
label2 = tk.Label(root,text=name[1], fg='green',font=('helvetica',12,'bold'))
label3 = tk.Label(root,text=name[2], fg='green',font=('helvetica',12,'bold'))
label4 = tk.Label(root,text=name[3], fg='green',font=('helvetica',12,'bold'))
label5 = tk.Label(root,text='my money', fg='green',font=('helvetica',12,'bold'))
canvas1.create_window(100,50,window=label1)
canvas1.create_window(100,100,window=label2)
canvas1.create_window(100,150,window=label3)
canvas1.create_window(100,200,window=label4)
canvas1.create_window(100,300,window=label5)
#buy button
button1 = tk.Button(text='buy',command=one,bg='brown',fg='white')
button2 = tk.Button(text='buy',command=two,bg='brown',fg='white')
button3 = tk.Button(text='buy',command=three,bg='brown',fg='white')
button4 = tk.Button(text='buy',command=four,bg='brown',fg='white')
canvas1.create_window(150,50, window=button1)
canvas1.create_window(150,100, window=button2)
canvas1.create_window(150,150, window=button3)
canvas1.create_window(150,200, window=button4)

root.mainloop()

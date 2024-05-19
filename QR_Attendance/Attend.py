from tkinter.constants import GROOVE, RAISED, RIDGE
import cv2
import pyzbar.pyzbar as pyzbar
import time
from datetime import date, datetime
import tkinter as tk 
from tkinter import Frame, ttk, messagebox
from tkinter import *

window = tk.Tk()
window.title('Attendance- Uttistha Kaunteya:Roar of the Lion and the Lion Inside')
window.geometry('900x600') 
                          
                          
seminar= tk.StringVar()
section= tk.StringVar()

title = tk.Label(window,text="Attendance- Uttistha Kaunteya:Roar of the Lion and the Lion Inside",bd=10,relief=tk.GROOVE,font=("times new roman",23),bg="lavender",fg="black")
title.pack(side=tk.TOP,fill=tk.X)

Manage_Frame=Frame(window,bg="lavender")
Manage_Frame.place(x=0,y=80,width=480,height=530)

ttk.Label(window, text = "Seminar",background="lavender", foreground ="black",font = ("Times New Roman", 15)).place(x=100,y=200)
combo_search=ttk.Combobox(window,textvariable=seminar,width=15,font=("times new roman",13),state='readonly')
combo_search['values']=("Uttistha_Kaunteya: Roar_of the_Lion and_the Lion_Inside")
combo_search.place(x=250,y=200)

ttk.Label(window, text = "Section",background="lavender", foreground ="black",font = ("Times New Roman", 15)).place(x=100,y=300)
combo_search=ttk.Combobox(window,textvariable=section,width=15,font=("times new roman",13),state='readonly')
combo_search['values']=("Yuvak", "Yuvati")
combo_search.place(x=250,y=300)

def checkk():
    if(seminar.get() and section.get()):
        window.destroy()
    else:
        messagebox.showwarning("Warning", "All fields required!!")

exit_button = tk.Button(window,width=13, text="Submit",font=("Times New Roman", 15),command=checkk,bd=2,relief=RIDGE)
exit_button.place(x=300,y=380)



Manag_Frame=Frame(window,bg="lavender")
Manag_Frame.place(x=480,y=80,width=450,height=530)

canvas = Canvas(Manag_Frame, width = 300, height = 300,background="lavender")
canvas.pack()
img = PhotoImage(file="uttishthakaunteya_1.png")
canvas.create_image(50,50, anchor=NW, image=img)

window.mainloop()

if not section.get(): #if the section is not chosen, set the default to Yuva
    section.set("Yuva")

cap = cv2.VideoCapture(0)
names=[]
today=date.today()
current_time = datetime.now().strftime("%H-%M-%S")
d= today.strftime("%b-%d-%Y")

fob=open(f"{d}_{current_time}.xlsx",'w+')
fob.write("Reg No."+'\t')
fob.write("Full Name"+'\t')
fob.write("Mandal"+'\t')
fob.write("Yuvak/Yuvati"+'\t')
fob.write("Mobile"+'\t')
fob.write("In Time"+'\n')

def enterData(reg_no, full_name, mandal, mobile):

                it = datetime.now()
                names.append(reg_no)
                intime = it.strftime("%H:%M:%S")
                data = f"{reg_no}\t{full_name}\t{mandal}\t{mobile}\t{intime}\n"
                fob.write(reg_no+'\t'+full_name+'\t'+mandal+'\t'+section.get()+'\t'+mobile+'\t'+intime+'\n')

                return names

print('Reading...')

def checkData(reg_no, full_name, mandal, mobile):
    # data=str(data)
    if reg_no in names:
        print('Already Present')
    else:
        print('New entry detected...')
        enterData(reg_no, full_name, mandal, mobile)
        print('SAVED')


while True:
    _, frame = cap.read()
    decodedObjects = pyzbar.decode(frame)
    for obj in decodedObjects:
        data = obj.data.decode('utf-8').split()
        if len(data) == 5:
            if len(data) == 5 and data[0].startswith('EID'):
                reg_no = data[0][3:] # SKIPS THE EID FROM THE ID eg. EID777777
                full_name = data[1] + ' ' + data[2]
                mandal = data[3]
                mobile = data[4]
            checkData(reg_no, full_name, mandal, mobile)
        else:
            print("Invalid QR code")
        time.sleep(1)
       
    cv2.imshow("Frame", frame)

    if cv2.waitKey(1)&0xFF == ord('g'):
        cv2.destroyAllWindows()
        break
    
fob.close()

from tkinter import font 
from tkinter import *
from tkinter import messagebox
import pandas as pd
from xlwt import Workbook
from datetime import datetime

# Add users here
names = ["Leo Jiang", "Eric Chen", "Your Mom"]# Todo: add button to add names

df = pd.DataFrame({"Name": names, "HoursAttendence": [0] * len(names), "CheckedInTime" : [-1] * len(names)})

try:
    excelData = pd.read_excel(io='output.xlsx')

    for currentName in names:
        try:
            df.HoursAttendence[df.Name == currentName] = excelData.HoursAttendence[excelData.Name == currentName].item()
        except:
            pass

except FileNotFoundError:
    pass




frame = Tk()
frame.state('zoomed')



def check():
    
    isCheckedIn = df.CheckedInTime[df.Name == chosen_name.get()].item() != -1
    

    if isCheckedIn:
        btn_text.set("Check In")
        label.config(text=chosen_name.get() + " Checked Out")

        checkInLength = datetime.now() - df.CheckedInTime[df.Name == chosen_name.get()].item()
        df.HoursAttendence[df.Name == chosen_name.get()] = df.HoursAttendence[df.Name == chosen_name.get()].item() + (checkInLength.total_seconds() / 3600)


        df.CheckedInTime[df.Name == chosen_name.get()] = -1
    else:
        btn_text.set("Check Out")
        label.config(text=chosen_name.get() + " Checked In")

        df.CheckedInTime[df.Name == chosen_name.get()] = datetime.now()

def changeButtonText(*args):
    isCheckedIn = df.CheckedInTime[df.Name == chosen_name.get()].item() != -1

    if isCheckedIn:
        btn_text.set("Check Out")
    else:
        btn_text.set("Check In")



chosen_name = StringVar()
chosen_name.set(names[0])
chosen_name.trace("w", changeButtonText)

#Change Font Size Here
helv35=font.Font(family='Helvetica', size=36)

nameMenu = OptionMenu(frame, chosen_name, *names)
nameMenu.config(font=helv35)
frame.nametowidget(nameMenu.menuname).config(font=helv35)
nameMenu.pack()

btn_text = StringVar()
btn_text.set("Check In")
button = Button(frame, textvariable=btn_text, command=check)
button.place(relx=.5, rely=.5,anchor= CENTER)
button["font"] = helv35

label = Label(frame, text=" ", font=helv35)
label.pack()


def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        while True:
            ifSuceed = False
            if messagebox.askyesno("Save?", "Do you want to save?"):
                try:
                    df.drop(["CheckedInTime"], axis=1).to_excel("output.xlsx", sheet_name='attendence')
                    ifSuceed = True
                    break
                except:
                    pass
            else:
                ifSuceed = True
                
            if ifSuceed:
                break


        frame.destroy()

        

frame.protocol("WM_DELETE_WINDOW", on_closing)
frame.mainloop()

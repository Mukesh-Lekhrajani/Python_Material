#****************************************
#Telephone Diary (21 July 2019)
#Developed By - Mukesh Lekhrajani
#****************************************

from tkinter import *
from tkinter import messagebox
import win32com.client
import pyodbc

root = Tk()
root.title("My Personal Assistant -- Phone Diary")
w = root.winfo_screenwidth()
h = root.winfo_screenheight()
root.geometry("%dx%d+-8+0" % (w, h))
#root.minsize(root.winfo_screenwidth() , root.winfo_screenheight())

lbl_Header = Label(root, text = "My Personal Assistant", font="consolas 15", anchor="e", width=18).place(x = 55 , y = 55)



#Make DB Connection
daoEngine = win32com.client.Dispatch('DAO.DBEngine.120')
db = 'J:\Python\PythonTest.mdb'
daoDB = daoEngine.OpenDatabase(db)
query = 'Select * FROM PhoneDiary'
daoRS = daoDB.OpenRecordset(query)
daoRS.MoveFirst()


#Variable Declaration
varName = StringVar()
varLandLine = StringVar()
varMobile = StringVar()
varImportant = IntVar()
tkvar = StringVar(root)
choices = { 'Business','Sports','Relatives','Enemies','Social Media'}
checkvar1  = IntVar()

#Display Data
def Display_Data():
    varName.set(daoRS.Contact_Name)
    varLandLine.set(daoRS.Land_Line_No)
    varMobile.set(daoRS.Mobile_No)
    tkvar.set(daoRS.Contact_Category)
    varImportant.set(daoRS.Is_Important)

def do_validations():
    if checkvar1.get() == 1:
        messagebox.showinfo("-Importance-","Yes, Its Important Contact")

def cmd_MoveNext():
    daoRS.MoveNext()
    Display_Data()
		
def cmd_MovePrevious():
    daoRS.MovePrevious()
    Display_Data()

def cmd_MoveFirst():
    daoRS.MoveFirst()
    Display_Data()
	
def cmd_MoveLast():
    daoRS.MoveLast()
    Display_Data()
	
def cmd_Exit():
    daoRS.Close()
	exit()
	
#Set Screen Layout
lbl_Name  = Label(root, text = "Name : ", font="consolas 10", anchor="e", width=18).place(x = 30 , y = 100)
lbl_LL    = Label(root, text = "Land Line No. : ", font="consolas 10", anchor="e", width=18).place(x = 30,y = 130)
lbl_Cell  = Label(root, text = "Mobile No. : ", font="consolas 10", anchor="e", width=18).place(x = 30,y = 160)
lbl_Categ = Label(root, text = "Category : ", font="consolas 10", anchor="e", width=18).place(x = 30,y = 190)
lbl_Imp   = Label(root, text = "Important : ", font="consolas 10", anchor="e", width=18).place(x = 30,y = 220)

txt_Name   = Entry(root, textvariable=varName).place(x = 160, y = 102)
txt_LL     = Entry(root, textvariable=varLandLine).place(x = 160, y = 132)
txt_Cell   = Entry(root, textvariable=varMobile).place(x = 160, y = 162)
chk_imp    = Checkbutton(root, variable = varImportant, onvalue = 1, offvalue = 2).place(x = 155,y = 220)
popupMenu = OptionMenu(root, tkvar, *choices).place(x = 158, y = 188)

# Control Button
cmd_Save   = Button(root,text = "Save", width=10, command = do_validations).place(x = 100,y = 300)
cmd_Cancel = Button(root,text = "Cancel", width=10).place(x = 200,y = 300)


#Toolbar
cmdNew       = Button(root, text = "New"      , width=5, height=2).place(x = 3 , y = 3)
cmdEdit      = Button(root, text = "Edit"     , width=5, height=2).place(x = 47 , y = 3)
cmdDelete    = Button(root, text = "Delete"   , width=5, height=2).place(x = 91 , y = 3)
cmdFind      = Button(root, text = "Find"     , width=5, height=2).place(x = 135 , y = 3)
cmdFirst     = Button(root, text = "First"    , width=5, height=2 , command = cmd_MoveFirst).place(x = 185 , y = 3)
cmdPrevious  = Button(root, text = "Previous" , width=6, height=2 , command = cmd_MovePrevious).place(x = 229 , y = 3)
cmdNext      = Button(root, text = "Next"     , width=5, height=2 , command = cmd_MoveNext).place(x = 280 , y = 3)
cmdLast      = Button(root, text = "Last"     , width=5, height=2 , command = cmd_MoveLast).place(x = 324 , y = 3)
cmdReport    = Button(root, text = "Report"   , width=5, height=2).place(x = 374 , y = 3)
cmdHelp      = Button(root, text = "Help"     , width=5, height=2).place(x = 418 , y = 3)
cmdExit      = Button(root, text = "Exit"     , width=5, height=2 , command = cmd_Exit).place(x = 462 , y = 3)

Display_Data()


root.mainloop()



#Using Grid


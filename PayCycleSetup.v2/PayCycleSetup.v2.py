# Matt Potter
# Created May 31 2016 
# Last Edited June 27 2016
# Pay Cycle Setup v2.1.1

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pypyodbc
import re
import base64
import configparser

class MainApplication:
    def __init__(self, master, conn):
        self.master = master
        self.frame = ttk.Frame(self.master)
        self.master.wm_title("Pay Cycle Setup")
        self.master.geometry("%dx%d%+d%+d" % (370, 215, 250, 125))
        self.fields = ('Location', 'Pay Group', 'Tip Share', 'Pay Cycle', 'ADP Store Code')
        MainApplication.connection = conn
        MainApplication.cursor = self.connection.cursor()
        self.SQLCommand = ("")
        self.createWidgets()
        self.frame.grid(column=0, row=0)

    def createWidgets(self):
        for field in self.fields:
            if field == 'Location': 
                self.LocationVariable = tk.StringVar()
                self.lLab = tk.Label(self.frame, text="Location: ")
                self.lLab.grid(column=1, columnspan=2, row=1, padx=5, pady=5, sticky='W')
                self.SQLCommand = ("select [SiteName] from [POSLabor].[dbo].[NBO_Sites]")
                MainApplication.cursor.execute(self.SQLCommand)
                self.siteNames = MainApplication.cursor.fetchall()
                self.lEntry = ttk.OptionMenu(self.frame, self.LocationVariable, "Site Name", *self.siteNames)
                self.lEntry.grid(column=4, columnspan=4, row=1, padx=5, pady=5)

            elif field == 'Pay Group': 
                MainApplication.PayGroupVariable = tk.StringVar()
                self.gLab = tk.Label(self.frame, text="Pay Group: ")
                self.gLab.grid(column=1, columnspan=2, row=2, padx=5, pady=5, sticky='W')
                self.SQLCommand = ("select distinct [PayrollGroupName] from [POSLabor].[dbo].[NBO_PayGroup]")
                MainApplication.cursor.execute(self.SQLCommand)
                MainApplication.payGroups = MainApplication.cursor.fetchall()
                MainApplication.gEntry = ttk.OptionMenu(self.frame, MainApplication.PayGroupVariable, "Payroll Group Name", *MainApplication.payGroups)
                MainApplication.gEntry.grid(column=4, columnspan=4, row=2, padx=5, pady=5)

            elif field == 'Tip Share':
                self.tLab = tk.Label(self.frame, text="Tip Share: ")
                self.tLab.grid(column=1, columnspan=2, row=3, padx=5, pady=5, sticky='W')
                self.TipShareVariable = tk.BooleanVar()
                self.tCheck = tk.Checkbutton(self.frame, variable=self.TipShareVariable)
                self.tCheck.grid(column=4, columnspan=4, row=3, padx=5, pady=5)

            elif field == 'Pay Cycle': 
                self.cLab = tk.Label(self.frame, text="Pay Cycle: ")
                self.cLab.grid(column=1, columnspan=2, row=4, padx=5, pady=5, sticky='W')
                self.PayCycleVariable = tk.StringVar()
                self.cEntry = ttk.OptionMenu(self.frame, self.PayCycleVariable, "Pay Cycle", *range(1,5))
                self.cEntry.grid(column=4, columnspan=4, row=4, padx=5, pady=5)

            elif field == 'ADP Store Code': 
                self.sLab = tk.Label(self.frame, text="ADP Store Code: ")
                self.sLab.grid(column=1, columnspan=2, row=5, padx=5, pady=5, sticky='W')
                self.ADPStoreCodeVariable = tk.StringVar()
                self.sEntry = tk.Entry(self.frame, textvariable=self.ADPStoreCodeVariable, width=30)
                self.sEntry.grid(column=4, columnspan=4, row=5, padx=5, pady=5)

        self.addButton = tk.Button(self.frame, text="Add Pay Group", command=self.newAddWindow)
        self.addButton.grid(column=3, columnspan=2, row=6, padx=5, pady=5)
        self.editButton = tk.Button(self.frame, text="Edit Pay Group", command=self.newEditWindow)
        self.editButton.grid(column=5, columnspan=2, row=6, padx=5, pady=5)
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=7, columnspan=2, row=6, padx=5, pady=5)
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.master.destroy)
        self.cancelButton.grid(column=1, columnspan=2, row=6, padx=5, pady=5, sticky='W')

    def newAddWindow(self):
        self.newWindow = tk.Toplevel(self.master)
        self.app = addWindow(self.newWindow)

    def newEditWindow(self):
        self.newWindow = tk.Toplevel(self.master)
        self.app = editWindow(self.newWindow)

    def submit(self):        
        self.loc = self.LocationVariable.get()
        self.loc = self.loc.strip("(\",)")
        self.loc = str(self.loc).replace("'","%")
        self.SQLCommand = ("SELECT [SiteNumber] FROM [POSLabor].[dbo].[NBO_Sites] where [SiteName] like '" + self.loc + "';")
        MainApplication.cursor.execute(self.SQLCommand)
        self.loc = MainApplication.cursor.fetchone()
        self.loc = str(self.loc).strip("(,)")
               
        self.payg = MainApplication.PayGroupVariable.get()
        self.payg = self.payg.strip("(',)")
        self.SQLCommand = ("SELECT [PayGroupID] FROM [POSLabor].[dbo].[NBO_PayGroup] where [PayrollGroupName] like '" + self.payg + "';")
        MainApplication.cursor.execute(self.SQLCommand)
        self.payg = MainApplication.cursor.fetchone()
        self.payg = str(self.payg).strip("(,)")
        
        self.tip = int(self.TipShareVariable.get())
        self.payc = self.PayCycleVariable.get()
        self.adp = self.ADPStoreCodeVariable.get()
        self.insertSQL(self.loc, self.payg, self.tip, self.payc, self.adp)    


    def insertSQL(self, loc, payg, tip, payc, adp):
        try:
            self.SQLCommand = ("DECLARE @RT INT "\
                           "EXECUTE @RT = dbo.pr_NBO_PayCycleSetup_ADD " + self.loc + ", " + self.payg + ", " + str(self.tip) + ", '" + self.adp + "', 0, " + self.payc + " PRINT @RT") # command to add data
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()
            self.mBox = tk.messagebox.showinfo("Success!","Import Complete")

        except:
            self.mBox = tk.messagebox.showinfo("Error!","Import Failed")
            print("Failed Command: " + self.SQLCommand)
            MainApplication.connection.rollback()
            
        self.master.destroy()


class addWindow:
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Add Pay Group")
        self.master.geometry("%dx%d%+d%+d" % (250, 100, 250, 125))
        self.newPayGroupName = tk.StringVar()
        self.createWidgets()
        self.frame.grid(column=0, row=0)

    def createWidgets(self):        
        self.lab = tk.Label(self.frame, text="New Pay Group Name")
        self.lab.grid(column=1, columnspan = 4, row=1, padx=5, pady=5, sticky='W')
        self.entry = tk.Entry(self.frame, textvariable=self.newPayGroupName, width = 39) 
        self.entry.grid(column=1, columnspan=4, row=2, padx=5, pady=5, sticky='W') 
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=1, row=3, padx=5, pady=5, sticky='W')
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=4, row=3, padx=5, pady=5, sticky='E')

    def submit(self):
        self.flag = False
        for name in MainApplication.payGroups:
            if str(self.newPayGroupName.get()) in name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True

        if self.flag == False:
            MainApplication.payGroups.insert((len(MainApplication.payGroups) + 1), self.newPayGroupName.get())
            self.SQLCommand = ("INSERT INTO [POSLabor].[dbo].[NBO_PayGroup] (PayrollGroupName, PayGroupID) " \
                          "VALUES ('" + str(self.newPayGroupName.get()) + "', " + str(len(MainApplication.payGroups)) + " );")                      
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()

        MainApplication.payGroups.insert(0,self.newPayGroupName)
        MainApplication.gEntry.set_menu(*MainApplication.payGroups)
        MainApplication.PayGroupVariable.set(self.newPayGroupName.get())

        self.master.destroy()
        
    def close_windows(self):
        self.master.destroy()

    
class editWindow:
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Edit Pay Group")
        self.master.geometry("%dx%d%+d%+d" % (360, 210, 250, 125))
        self.createWidgets()
        self.frame.grid(column=0, row=0)

    def createWidgets(self):
        self.NewPayGroupName = tk.StringVar()

        self.lab1 = tk.Label(self.frame, text="Pay Group to Edit:")
        self.lab1.grid(column=1, columnspan=5, row=1, sticky='W', padx=5, pady=5,)

        self.nameList = tk.Listbox(self.frame)
        self.count = 0
        for name in MainApplication.payGroups:
            name = str(name).strip("(,)")
            self.nameList.insert(self.count, name)
            self.count = self.count + 1
        self.nameList.grid(row=2, rowspan=5, column=1, padx=5, pady=5)

        self.lab2 = tk.Label(self.frame, text="New Pay Group Name: ", anchor='w', width=30)
        self.lab2.grid(row=3, column=3, columnspan=3)
        self.entry = tk.Entry(self.frame, textvariable=self.NewPayGroupName, width=30)
        self.entry.grid(row=4, column=3, columnspan=3)
        self.entry.grid_columnconfigure(3, weight=1)
        self.submitButton = tk.Button(self.frame, text="Submit Edit", command=lambda: self.submit(self.nameList.get(tk.ACTIVE), self.NewPayGroupName))
        self.submitButton.grid(column=5, row=6, padx=5, pady=5, sticky='E') 
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=4, row=6, sticky='E')

    def submit(self, OldPayGroupName, NewPayGroupName):
        self.flag = False
        for name in MainApplication.payGroups:
            if str(NewPayGroupName.get()) in name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True 

        if self.flag == False:
            # use sql statement to UPDATE values to new values
            OldPayGroupName = str(OldPayGroupName).replace("'", "%")
            self.SQLCommand = ("UPDATE [POSLabor].[dbo].[NBO_PayGroup] " \
                          "SET [PayrollGroupName]='" + str(NewPayGroupName.get()) + "' " + \
                          "WHERE [PayrollGroupName] like '" + OldPayGroupName + "';")          
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()

            OldPayGroupName = OldPayGroupName.strip('%')
            for index, item in enumerate(MainApplication.payGroups):
                if (OldPayGroupName in item):
                    MainApplication.payGroups[index] = str(NewPayGroupName.get())            

            MainApplication.gEntry.set_menu(*MainApplication.payGroups)
            MainApplication.PayGroupVariable.set(NewPayGroupName.get())
            self.master.destroy()
        
    def close_windows(self):
        self.master.destroy()


def main():
    #'Driver={SQL Server};'     
    #'Server=VSQL15\NBOTest;'    
    #'Database=POSLabor;'       
    #'uid=LaborLoad;'            
    #'pwd=B8{uwBJ!ZxZ{'
    
    Config = configparser.ConfigParser()
    Config.read("config.ini")
    
    # reading base64 login from user.txt
    driver =    Config.get("Login","Driver")
    server =    Config.get("Login","Server")
    database =  Config.get("Login","Database")
    uid =       Config.get("Login","uid")
    pwd =       Config.get("Login","pwd")

    # decoding credentials
    driver =    str(base64.b64decode(driver)).strip('b\'')
    server =    str(base64.b64decode(server)).strip('b\'')
    ind = server.index("\\\\")
    server = server[:ind] + server[ind+1:]
    database =  str(base64.b64decode(database)).strip('b\'')
    uid =       str(base64.b64decode(uid)).strip('b\'')
    pwd =       str(base64.b64decode(pwd)).strip('b\'')

    login =     ("Driver=%s;Server=%s;Database=%s;uid=%s;pwd=%s" % (driver, server, database, uid, pwd))

    connection = pypyodbc.connect(login)
    root = tk.Tk()
    app = MainApplication(root, connection)
    root.mainloop()
    connection.close()


if __name__ == '__main__':
    main()

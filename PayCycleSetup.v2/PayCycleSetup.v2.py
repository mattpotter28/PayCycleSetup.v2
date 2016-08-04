# Matt Potter
# Created May 31 2016 
# Last Edited August 2 2016
# Pay Cycle Setup v3.1.0

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pypyodbc
import re
import base64
import configparser

class MainApplication: # main window
    def __init__(self, master, conn, serv): 
        # region constructor
        self.master = master
        self.frame = ttk.Frame(self.master)
        self.master.wm_title("Pay Cycle Setup")
        self.master.geometry("%dx%d%+d%+d" % (455, 225, 250, 125))
        self.fields = ('Location', 'Pay Group', 'Tip Share', 'Pay Cycle', 'ADP Store Code')
        MainApplication.connection = conn
        MainApplication.cursor = self.connection.cursor()
        self.SQLCommand = ("")
        self.serverName = serv
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        #endregion constructor
    
    def createWidgets(self):
        # region Label/Option Menu creation
        self.serverLab = tk.Label(self.frame, text="Connected to: "+self.serverName)
        self.serverLab.grid(column=1, columnspan=4, row=0, padx=5, pady=5, sticky='W')
        for field in self.fields:
            if field == 'Location': 
                # region Location Field
                self.LocationVariable = tk.StringVar()
                self.lLab = tk.Label(self.frame, text="Location: ")
                self.lLab.grid(column=1, columnspan=2, row=1, padx=5, pady=5, sticky='W')
                self.SQLCommand = ("select [SiteName] FROM dbo.NBO_Sites")
                MainApplication.cursor.execute(self.SQLCommand)
                self.siteNames = MainApplication.cursor.fetchall()
                self.siteNames = [str(s).strip('{(\"\',)}') for s in self.siteNames]
                self.siteNames.sort()
                MainApplication.lEntry = ttk.Combobox(self.frame, textvariable=self.LocationVariable, value = self.siteNames, state="readonly")      
                MainApplication.lEntry.grid(column=4, columnspan=6, row=1, padx=5, pady=5, sticky="ew")
                MainApplication.lEntry.bind('<<ComboboxSelected>>', self.locationSelect)
                # endregion Location Field

            elif field == 'Pay Group': 
                # region Pay Group Field
                MainApplication.PayGroupVariable = tk.StringVar()
                self.gLab = tk.Label(self.frame, text="Pay Group: ")
                self.gLab.grid(column=1, columnspan=2, row=2, padx=5, pady=5, sticky='W')
                self.SQLCommand = ("select distinct [PayrollGroupName] from [POSLabor].[dbo].[NBO_PayGroup]")
                MainApplication.cursor.execute(self.SQLCommand)
                MainApplication.payGroups = MainApplication.cursor.fetchall()
                MainApplication.payGroups = [str(s).strip('{(\"\',)}') for s in MainApplication.payGroups]
                MainApplication.gEntry = ttk.Combobox(self.frame, textvariable=MainApplication.PayGroupVariable, values=MainApplication.payGroups, state="readonly")
                MainApplication.gEntry.grid(column=4, columnspan=6, row=2, padx=5, pady=5, sticky="ew")
                # endregion Pay Group Field

            elif field == 'Tip Share':
                # region Tip Share Field
                self.tLab = tk.Label(self.frame, text="Tip Share: ")
                self.tLab.grid(column=1, columnspan=2, row=3, padx=5, pady=5, sticky='W')
                self.TipShareVariable = tk.BooleanVar()
                self.tCheck = tk.Checkbutton(self.frame, variable=self.TipShareVariable)
                self.tCheck.grid(column=4, columnspan=6, row=3, padx=5, pady=5, sticky='W')
                # endregion Tip Share Field

            elif field == 'Pay Cycle': 
                # region Pay Cycle Field
                self.cLab = tk.Label(self.frame, text="Pay Cycle: ")
                self.cLab.grid(column=1, columnspan=2, row=4, padx=5, pady=5, sticky='W')
                self.PayCycleVariable = tk.StringVar()
                self.cEntry = ttk.Combobox(self.frame, textvariable=self.PayCycleVariable, values=(1,2,3,4))
                self.cEntry.grid(column=4, columnspan=6, row=4, padx=5, pady=5, sticky='W')
                # endregion Pay Cycle Field

            elif field == 'ADP Store Code':
                # region ADP Code Field 
                self.sLab = tk.Label(self.frame, text="ADP Store Code: ")
                self.sLab.grid(column=1, columnspan=2, row=5, padx=5, pady=5, sticky='W')
                self.ADPStoreCodeVariable = tk.StringVar()
                self.sEntry = tk.Entry(self.frame, textvariable=self.ADPStoreCodeVariable, width=30)
                self.sEntry.grid(column=4, columnspan=6, row=5, padx=5, pady=5, sticky='W')
                # endregion ADP Code Field
        
        # endregion Label/Option Menu creation
        
        # region Button creation
        self.tableButton = tk.Button(self.frame, text="View Table", command=self.newTableWindow)
        self.tableButton.grid(column=3, columnspan=2, row=6, padx=5, pady=5)
        self.addButton = tk.Button(self.frame, text="Add Pay Group", command=self.newAddWindow)
        self.addButton.grid(column=5, columnspan=2, row=6, padx=5, pady=5)
        self.editButton = tk.Button(self.frame, text="Edit Pay Group", command=self.newEditWindow)
        self.editButton.grid(column=7, columnspan=2, row=6, padx=5, pady=5)
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=9, columnspan=2, row=6, padx=5, pady=5)
        self.cancelButton = tk.Button(self.frame, text="Close", command=self.master.destroy)
        self.cancelButton.grid(column=1, columnspan=2, row=6, padx=5, pady=5, sticky='W')
        # endregion Button creation

    def newTableWindow(self):
        # region opens Table Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = tableWindow(self.newWindow)
        # endregion opens Table Window
        
    def newAddWindow(self):
        # region opens Add Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = addWindow(self.newWindow)
        # endregion opens Add Window

    def newEditWindow(self):
        # region opens Edit Window
        self.newWindow = tk.Toplevel(self.master)
        self.app = editWindow(self.newWindow)
        # endregion region opens Edit Window

    def submit(self):        
        # region SQLStatement creation

        ## Location conversion ##
        self.loc = self.LocationVariable.get()
        self.loc = self.loc.strip("(\",)")
        self.loc = str(self.loc).replace("'","%")
        self.SQLCommand = ("SELECT [SiteNumber] FROM [POSLabor].[dbo].[NBO_Sites] where [SiteName] like '" + self.loc + "';")
        MainApplication.cursor.execute(self.SQLCommand)
        self.loc = MainApplication.cursor.fetchone()
        self.loc = str(self.loc).strip("(,)")
               
        ## Pay Group conversion ##
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

        # endregion SQLStatement creation

    def insertSQL(self, loc, payg, tip, payc, adp):
        try:
            # region Statement submittion
            self.SQLCommand = ("DECLARE @RT INT "\
                           "EXECUTE @RT = dbo.pr_NBO_PayCycleSetup_ADD " + self.loc + ", " + self.payg + ", " + str(self.tip) + ", '" + self.adp + "', 0, " + self.payc + " PRINT @RT") # command to add data
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()
            self.mBox = tk.messagebox.showinfo("Success!","Import Complete")
            # endregion Statement submittion

        except:
            # region submition error
            self.mBox = tk.messagebox.showinfo("Error!","Import Failed")
            print("Failed Command: " + self.SQLCommand)
            MainApplication.connection.rollback()
            # endregion submition error
            
        self.master.destroy()

    def locationSelect(self, oth):
        # check current location selection
        location = self.siteNames[MainApplication.lEntry.current()]
        location = str(location).replace("'", "%")

        # clear current values from entries
        self.gEntry.set('')
        self.tCheck.deselect()
        self.cEntry.set('')
        self.sEntry.delete(0,'end')

        # find corresponding site number to name
        self.SQLCommand = ("select [SiteNumber] from dbo.NBO_Sites where [SiteName] like \'%"+location+"%\'")
        MainApplication.cursor.execute(self.SQLCommand)
        siteNumber = MainApplication.cursor.fetchall()
        siteNumber = str(siteNumber).strip("[(\',)]")

        # take in values based on site number
        self.SQLCommand = ("SELECT * FROM [POSLabor].[dbo].[NBO_PayCycleSetup] where [SiteNumber] like \'%"+siteNumber+"%\'")
        MainApplication.cursor.execute(self.SQLCommand) 
                   
        # fill entries with values
        for row in MainApplication.cursor:
            # 1 - PayGroupID: translate ID to name and match with combobox option
            self.SQLCommand = ("SELECT [PayrollGroupName] FROM [POSLabor].[dbo].[NBO_PayGroup] where [PayGroupID] = \'"+str(row[1])+"\'")
            MainApplication.cursor.execute(self.SQLCommand)
            groupName = str(MainApplication.cursor.fetchall()).strip("[(\',)]")
            self.gEntry.set(groupName)

            # 2 - TipShareLocation: translate to boolean to apply to checkbutton
            if int(row[2]) == 1:
                self.tCheck.select()

            # 3 - ADPStoreCode: fill entry with string
            self.sEntry.insert(0, str(row[3]))

            # 5 - NBOCalendarNumber: translate to paycycle and match with combobox option
            if row[5] == 1693843:
                self.cEntry.set('1')
            elif row[5] == 1693844:
                self.cEntry.set('2')
            elif row[5] == 1693845:
                self.cEntry.set('3')
            elif row[5] == 1837816:
                self.cEntry.set('4')


class tableWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("View Table")
        self.master.geometry("%dx%d%+d%+d" % (855, 275, 250, 125))
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):
        self.table = ttk.Treeview(self.frame)
        self.table.grid(column=1, columnspan=6, row=1, rowspan=5, padx=(5,0), pady=5)
        self.table['columns'] = ('pgID', 'tipShare', 'adp', 'active', 'nbo')
        self.table.heading('#0', text="Site Name")
        self.table.column('#0', width=275)
        self.table.heading('pgID', text="Payroll Group Name")
        self.table.column('pgID', width=150)
        self.table.heading('tipShare', text="Tip Share")
        self.table.column('tipShare', width=100)
        self.table.heading('adp', text="ADP Store Code")
        self.table.column('adp', width=100)
        self.table.heading('active', text="Active Location")
        self.table.column('active', width=100)
        self.table.heading('nbo', text="Pay Cycle")
        self.table.column('nbo', width=100)

        self.SQLCommand = ("EXECUTE dbo.pr_NBO_PayCycleReport")
        MainApplication.cursor.execute(self.SQLCommand)
        for row in MainApplication.cursor:
            self.table.insert('', 'end', text=str(row[0]), values=(row[1], row[2], row[3], row[4], row[5]))
        sb = ttk.Scrollbar(self.frame, orient='vertical', command=self.table.yview)
        self.table.configure(yscroll=sb.set)
        sb.grid(row=1, column=9, rowspan=5, sticky="NS", pady=5)  
        self.closeButton = tk.Button(self.frame, text="Close", command=self.close_windows)
        self.closeButton.grid(column=1, columnspan=6, row=6, padx=5, pady=5)

    def close_windows(self):
        self.master.destroy()

class addWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Add Pay Group")
        self.master.geometry("%dx%d%+d%+d" % (250, 100, 250, 125))
        self.newPayGroupName = tk.StringVar()
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):        
        # region Label/Entry creation
        self.lab = tk.Label(self.frame, text="New Pay Group Name")
        self.lab.grid(column=1, columnspan = 4, row=1, padx=5, pady=5, sticky='W')
        self.entry = tk.Entry(self.frame, textvariable=self.newPayGroupName, width = 39) 
        self.entry.grid(column=1, columnspan=4, row=2, padx=5, pady=5, sticky='W')
        # endregion Label/Entry creation

        # region Button creation
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=1, row=3, padx=5, pady=5, sticky='W')
        self.submitButton = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submitButton.grid(column=4, row=3, padx=5, pady=5, sticky='E')
        # endregion Button creation

    def submit(self):
        # region name validation
        self.flag = False
        for name in MainApplication.payGroups:
            print(name+"\n"+self.newPayGroupName.get())
            if str(self.newPayGroupName.get()) == name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True
        # endregion name validation

        # region SQL submittion
        if self.flag == False:
            MainApplication.payGroups.insert((len(MainApplication.payGroups) + 1), self.newPayGroupName.get())
            self.SQLCommand = ("INSERT INTO [POSLabor].[dbo].[NBO_PayGroup] (PayrollGroupName, PayGroupID) " \
                          "VALUES ('" + str(self.newPayGroupName.get()) + "', " + str(len(MainApplication.payGroups)) + " );")                      
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()        

            # region resetting menus
            MainApplication.payGroups.insert(0,self.newPayGroupName)
            MainApplication.gEntry['values'] = MainApplication.payGroups
            MainApplication.PayGroupVariable.set(self.newPayGroupName.get())
            # endregion resetting menus

        # endregion SQL submittion

        self.master.destroy()
        
    def close_windows(self):
        self.master.destroy()

    
class editWindow:
    def __init__(self, master):
        # region constructor
        self.master = master
        self.frame = tk.Frame(self.master)
        self.master.wm_title("Edit Pay Group")
        self.master.geometry("%dx%d%+d%+d" % (411, 205, 250, 125))
        self.createWidgets()
        self.frame.grid(column=0, row=0)
        # endregion constructor

    def createWidgets(self):        
        # region Label/Entry creation
        
        self.lab1 = tk.Label(self.frame, text="Pay Group to Edit:")
        self.lab1.grid(column=1, columnspan=5, row=1, sticky='W', padx=5, pady=5)

        ## nameList
        self.nameList = tk.Listbox(self.frame, width=25)
        self.count = 0
        for name in MainApplication.payGroups:
            name = str(name).strip("(,)")
            self.nameList.insert(self.count, name)
            self.count = self.count + 1
        self.nameList.grid(row=2, rowspan=6, column=1, padx=(5,0), pady=5)
        sb = ttk.Scrollbar(self.frame, orient='vertical', command=self.nameList.yview)
        self.nameList.configure(yscroll=sb.set)
        sb.grid(row=2, column=2, rowspan=6, sticky="NS", pady=5)
        self.lab2 = tk.Label(self.frame, text="New Pay Group Name: ", anchor='w', width=30)
        self.lab2.grid(row=3, column=3, columnspan=4, padx=10, pady=5, sticky='S')
        self.NewPayGroupName = tk.StringVar()
        self.entry = tk.Entry(self.frame, textvariable=self.NewPayGroupName, width=30)
        self.entry.grid(row=4, column=3, columnspan=4, padx=10, sticky='NW')
        self.entry.grid_columnconfigure(3, weight=1)
        self.submitButton = tk.Button(self.frame, text="Submit Edit", command=lambda: self.submit(self.nameList.get(tk.ACTIVE), self.NewPayGroupName))
        self.submitButton.grid(column=6, columnspan=2, row=7, padx=5, pady=5, sticky='SE') 
        self.cancelButton = tk.Button(self.frame, text="Cancel", command=self.close_windows)
        self.cancelButton.grid(column=4, columnspan=2, row=7, padx=(5,0), pady=5, sticky='SE')
        # endregion Label/Entry creation

    def submit(self, OldPayGroupName, NewPayGroupName):
        # region name validation
        self.flag = False
        for name in MainApplication.payGroups:
            if str(NewPayGroupName.get()) in name:
                self.mBox = tk.messagebox.showinfo("Error!","Name Already in Use")
                self.flag = True 
        # endregion name validation

        # region SQL submittion 
        if self.flag == False:
            # use sql statement to UPDATE values to new values
            OldPayGroupName = str(OldPayGroupName).replace("'", "%")
            self.SQLCommand = ("UPDATE [POSLabor].[dbo].[NBO_PayGroup] " \
                          "SET [PayrollGroupName]='" + str(NewPayGroupName.get()) + "' " + \
                          "WHERE [PayrollGroupName] like '" + OldPayGroupName + "';")          
            MainApplication.cursor.execute(self.SQLCommand)
            MainApplication.connection.commit()

            # region resetting menus
            OldPayGroupName = OldPayGroupName.strip('%')
            for index, item in enumerate(MainApplication.payGroups):
                if (OldPayGroupName in item):                    
                    MainApplication.payGroups[index] = str(NewPayGroupName.get())               
                    MainApplication.gEntry.set(MainApplication.payGroups[index])
            MainApplication.gEntry['values'] = MainApplication.payGroups   
            MainApplication.PayGroupVariable.set(NewPayGroupName.get())            
            # endregion resetting menus

            self.master.destroy()

        # endregion SQL submittion
        
    def close_windows(self):
        self.master.destroy()


def main():   
    #region login
    # opens config file
    Config = configparser.ConfigParser()
    Config.read("config.ini")
    
    # reading base64 login from config.ini
    driver =    Config.get("Login","Driver")
    server =    Config.get("Login","Server")
    database =  Config.get("Login","Database")
    uid =       Config.get("Login","uid")
    pwd =       Config.get("Login","pwd")

    # decoding credentials
    driver =    str(base64.b64decode(driver)).strip('b\'')
    server =    str(base64.b64decode(server)).strip('b\'')
    ind =       server.index("\\\\")
    server =    server[:ind] + server[ind+1:]
    database =  str(base64.b64decode(database)).strip('b\'')
    uid =       str(base64.b64decode(uid)).strip('b\'')
    pwd =       str(base64.b64decode(pwd)).strip('b\'')

    login =     ("Driver=%s;Server=%s;Database=%s;uid=%s;pwd=%s" % (driver, server, database, uid, pwd))

    connection = pypyodbc.connect(login)
    # endregion login

    # region open tkinter window
    root = tk.Tk()
    app = MainApplication(root, connection, server)
    root.mainloop()
    connection.close()
    # endregion open tkinter window


if __name__ == '__main__':
    main()

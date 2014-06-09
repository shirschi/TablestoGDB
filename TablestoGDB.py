#!/usr/bin/python
# -*- coding: iso-8859-1 -*-

import Tkinter
import tkMessageBox
from Tkinter import *
import arcpy, os
from arcpy import env
import tkFileDialog
import xlrd

class simpleapp_tk(Tkinter.Tk):
# User interface portion.
    # Setup frames.
    def __init__(self,parent):
        # Setup the main window.
        Tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()



    def initialize(self):

        #Setup the grid layout style
        self.grid()

        #Labels
        self.Dirlabel = Tkinter.Label(self, text="Select the parent directory.", font=('Times', 9, 'bold'))
        self.Dirlabel.grid(column=0, row=1, padx=8, pady=24, sticky='w')
        #  "gdb" will be a variable to hold the name of the geodatabase.
        self.gdb = Tkinter.StringVar()
        self.gdb = "ImportedTables.gdb"
        self.GDBlabel = Tkinter.Label(self, text="Type the name for your geodatabase?", font=('Times', 9, 'bold'))
        self.GDBlabel.grid(column=0, row=4, padx=8, pady=12, sticky='w')
        # Text entry box allowing the user to change the name of the geodatabase from the default of importedtables.gdb.
        # When the user clicks the import button it will record whatever text is in the box for the geodatabase
        # name and add.gdb if ext is not already present.
        self.DirEntry = Tkinter.Entry(self)
        self.DirEntry.grid(column=1, row=4, padx=8, pady=0, sticky='we')
        self.DirEntry.insert(0, self.gdb)
        self.Inst = Tkinter.Label(self, text="(Default is ImportedTables.gdb)", font=('Times', 8))
        self.Inst.grid(column=0, row=5, sticky='n')
        # Buttons
        self.FileBrowserB = Tkinter.Button(self,text=u"Click here to select a directory.",
                                command=self.FindDirectory)
        self.FileBrowserB.bind('<Button-1>', lambda event:self.OnButtonClick())
        self.FileBrowserB.grid(column=1,row=1,padx=8, sticky='w')
        self.ImportTablesB = Tkinter.Button(self, text=u"Click here to start importing tables", state=DISABLED, command= lambda: self.ImportTables(self.udt, self.gdb))
        self.ImportTablesB.grid(column=2, row=5, padx=8, pady=24, sticky='w')
        # Exit the program when the user clicks the red x to close the window.
        self.protocol("WM_DELETE_WINDOW", self.Exit)
        # Exit when pressing escape.
        self.bind('<Escape>', lambda e: self.quit())
        # Allows the window to be able to resize.
        self.resizable(True, True)
# Logical portion
    # Asks the user for the directory and assigns the full path to the variable "udt".
    def FindDirectory(self):
        directory = tkFileDialog.askdirectory(parent=root, title='Please select a directory.')
        # Create the \Database folder if it is not present.
        if not arcpy.Exists(directory + "\Database"):
            os.makedirs(directory + "\Database")
        self.udt = directory
    # Import button is disabled by default. This activates the button when the user clicks on the button to select a directory.
    # It would have been better to activate after a directory was confirmed but this way is was the best method I could find in the time available.
    def OnButtonClick(self):
        self.ImportTablesB.config(state=ACTIVE)

    # Used "quit()" to terminate the program. Loop continued when destroy() was used.
    def Exit(self):
        self.quit()

    # Function for the main logical portion of the script.
    def ImportTables(self, udt, gdb):
        # add the .gdb extention only if not already present.
        gdb = self.DirEntry.get()
        if gdb.split(".")[-1] != "gdb":
            gdb = gdb + ".gdb"

        # Create the requested database if it does not already exist.
        if not arcpy.Exists(udt + "\Database\\" + gdb):
            arcpy.CreateFileGDB_management(udt + "\Database", gdb, "CURRENT")

        # Setup rootdir to the path requested by user.
        rootdir = udt

        # set the workspace for the arcpy.ValidateTableName to work with the database created.
        env.workspace = udt + "\Database\\" + gdb

        # Uniquename is a function to ensure the imported table does not already exist and add a "_" & a "#" to the end if it does.
        def uniquename(un):
            un = un.split("\\")[-1]
            un = un.split(".")[0]
            eon = 1
            while arcpy.Exists(udt + "\Database\\" + gdb + "\\" + un):
                if un[-2] == "_" and un[-1].isdigit():
                    un = un[:-2]
                elif un[-3] == "_" and un[-2].isdigit():
                    un = un[:-3]
                un = un + "_" + str(eon)
                eon = eon + 1
            return un

        # Lessthan31 ensures the table will be 30 characters or less in order to be compatible with some oracle database front ends.
        def lessthan31(tablename):
            fn = tablename.split("\\")[-1]
            if len(fn) > 28:
                # Retain trailing numbers in case the tables are in sequence and over 30 characters.
                if (fn[-2] == "_" and fn[-1].isdigit()):
                    nn = fn[:28] + fn[-2] + fn[-1]
                elif fn[-1].isdigit():
                    nn = fn[:27] + fn[-1]
                else:
                    nn = fn[:28]
                tablename = tablename.replace(fn, nn)
            return tablename

        # Function to import .xls and .xlxs files.
        def importallsheets(in_excel, out_gdb):
            workbook = xlrd.open_workbook(in_excel)
            sheets = [sheet.name for sheet in workbook.sheets()]

            print('{} sheets found: {}'.format(len(sheets), ','.join(sheets)))
            for sheet in sheets:
                # The out_table is based on the input excel file name
                # a underscore (_) separator followed by the sheet name
                out_table = os.path.join(
                    out_gdb,
                    arcpy.ValidateTableName(
                        "{0}_{1}".format(os.path.basename(in_excel), sheet),
                        out_gdb))
                out_table = lessthan31(out_table)
                out_table = uniquename(out_table)

                print('Converting {} to {}'.format(sheet, out_table))
                # Perform the conversion and save the table in the geodatabase.
                arcpy.ExcelToTable_conversion(in_excel, out_table, sheet)

        # Iterate through all folders.
        for root,dirs,files in os.walk(rootdir, topdown = True):
                # skip the folder that contains the database.
                dirs[:] = [d for d in dirs if 'Database' not in d]
                # While iterating look for *.csv, .xls, xlxs or *.dbf files.
                for filename in files:
                    tablename = filename

                    # Extract *.csv or *.dbf.
                    if (filename.split('.')[1] == "csv") or (filename.split('.')[1] == "dbf"):
                        # Assign the entire path and filename to cf
                        cf = os.path.join(root, filename)
                        tablename = lessthan31(tablename)
                        tablename = udt + "\Database\\" + gdb + "\\" + tablename
                        tablename = uniquename(tablename)
                        # remove invalid characters for ArcGIS Tables.
                        tablename = arcpy.ValidateTableName(tablename)
                        arcpy.TableToTable_conversion(cf, udt + "\Database\\" + gdb + "\\", str(tablename))

                    elif (filename.split('.')[1] == "xls") or (filename.split('.')[1] == "xlsx"):
                        # Assign the entire path and filename to cf
                        cf = os.path.join(root, filename)
                        importallsheets(cf, gdb)
        # Inform the user when the import completes.
        tkMessageBox.showinfo("", "Import Complete")



# Tkinter app fundtion creates two windows. These lines hide an irrelevant window.
root = Tkinter.Tk()
root.withdraw()
# Create the loop and app.
if __name__ == "__main__":
    app = simpleapp_tk(None)
    app.title('Import tables to geodatabase tool')
    app.mainloop()
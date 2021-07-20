import tkinter as tk
import pandas as pd
import csv, os, xlrd
from tkintertable import *
from pandas import ExcelWriter
from pandas import ExcelFile
from pandastable import Table, TableModel
from tkinter import *
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror

class App:
    def __init__(self):
        self.root = Tk()
        self.root.title("Clicky Names")
        
        self.frame = Frame(self.root)
        self.frame.pack(fill=BOTH, expand=1)

        self.table = None
        self.table_model = None
        self.pt = None
        self.dataFrame = None
        self.full_name = None        
        self.first_name = None
        self.last_name = None
        self.middle_name = None
        self.row = None
        self.popup = None
        self.cell = None
        
        self.menu()
        self.root.mainloop()


    def format(self):

        #strip the comma from the Address column
        self.dataFrame['Address'] = self.dataFrame['Address'].str.replace(',', '')

        # strip the comma from the Vendor Name column
        self.dataFrame['Vendor Name'] = self.dataFrame['Vendor Name'].str.replace(',', '')

        # strip the dash from tax id/ssn column
        self.dataFrame['Tax ID Number'] = self.dataFrame['Tax ID Number'].str.replace('-', '')
        
        # format date as MM/DD/YYYY with no time stamp
        self.dataFrame['Last Check Date'] = self.dataFrame['Last Check Date'].dt.strftime('%m/%d/%Y')

        # drop the Address 1 and Address 2 columns
        self.dataFrame = self.dataFrame.drop(columns=['Address 1', 'Address 2'])
        
    # create the menu bar
    # user can open and save a file
    def menu(self):
        menu = Menu(self.root)
        self.root.config(menu=menu)
        file = Menu(menu)
        file.add_command(label='Open', command=self.open_file)
        file.add_command(label='Save', command=self.save_file)
        menu.add_cascade(label='File', menu=file)

    # opens dialog to select file
    def open_file(self):
        filename = askopenfilename()
        self.create_table(filename)

    # opens dialog to save a file
    def save_file(self):
        self.format()
        filename = asksaveasfilename(parent=self.root, filetypes=(("Excel file", "*.xlsx"), ("All files", "*.*")), title='Save file', defaultextension='.xlsx')
        self.dataFrame.to_excel(filename)
        
    # creates a table based on the selected excel file     
    def create_table(self, file):
        self.dataFrame = pd.read_excel(file)
        self.dataFrame.insert(2, 'LastName', '')
        self.dataFrame.insert(3, 'FirstName', '')
        self.dataFrame.insert(4, 'MiddleName', '')
        self.dataFrame.insert(5, 'Address', '')

        #concatenate Address 1 and Address 2
        cols = ['Address 1', 'Address 2']
        self.dataFrame["Address"] = self.dataFrame[cols].apply(lambda x: ' '.join(x.dropna()), axis=1)
        
        self.table = self.pt = Table(self.frame, dataframe=self.dataFrame, showtoolbar=False, showstatusbar=True)
        self.table_model = TableModel(self.dataFrame)
        self.pt.bind('<Button-1>', self.handle_left_click)
        self.pt.show()
        

    # left click event for table
    # left clicking on a cell will open a popup window to split the names of that cell
    def handle_left_click(self, event):
        self.row = self.pt.get_row_clicked(event)
        col = self.pt.get_col_clicked(event)
        self.cell = self.table_model.getValueAt(self.row, col)
        editable_cols_index = [2, 3, 4]
        if col == 1:
            self.popup = Toplevel()
            self.popup.title("Split Names")
            header = ['Full Name', 'Last Name', 'First Name', 'Middle Name']
            for i in header:
                if (i == 'Full Name'):
                    entry = Entry(self.popup, width=40)
                    entry.insert(0, i)
                    entry.grid(row=0, column=header.index(i))
                else:
                    entry = Entry(self.popup)
                    entry.insert(0, i)
                    entry.grid(row=0, column=header.index(i))
                
            self.full_name = Entry(self.popup, width=40)
            self.full_name.insert(0, self.cell)
            self.full_name.grid(row=1, column=0)
            
            self.last_name = Entry(self.popup)
            self.last_name.insert(0, self.table_model.getValueAt(self.row, col+1))
            self.last_name.grid(row=1, column=1)
            self.last_name.bind('<Button-1>', self.reset)
            
            self.first_name = Entry(self.popup)
            self.first_name.insert(0, self.table_model.getValueAt(self.row, col+2))
            self.first_name.grid(row=1, column=2)
            self.first_name.bind('<Button-1>', self.reset)

            self.middle_name = Entry(self.popup)
            self.middle_name.insert(0, self.table_model.getValueAt(self.row, col+3))
            self.middle_name.grid(row=1, column=3)
            self.middle_name.bind('<Button-1>', self.reset)

            self.full_name.bind('<ButtonRelease-1>', self.insert_name)
            self.popup.bind('<Button-3>', self.save_split_names)

    # right click event in the popup window
    # right clicking will insert the split names into the table
    def save_split_names(self, event):
        # self.table_model.setValueAt(self.full_name.get(), self.row, 1, df=self.dataFrame)
        self.table_model.setValueAt(self.last_name.get(), self.row, 2, df=self.dataFrame)
        self.table_model.setValueAt(self.first_name.get(), self.row, 3, df=self.dataFrame)
        self.table_model.setValueAt(self.middle_name.get(), self.row, 4, df=self.dataFrame)
        self.pt.redraw()
        self.pt.autoResizeColumns()
        self.popup.destroy()

    # selecting text in the 'Full Name' column will insert the name into the proper column   
    def insert_name(self, event):
        if(self.full_name.selection_present() == True):
            selection = self.full_name.selection_get()
            if(',' in selection): # if the selection has a comma
                sel_without_comma = selection.replace(',', '')
                if(self.last_name.get() == ''):
                    self.insert_into_widget(sel_without_comma, self.last_name)
                    self.delete_selection(self.full_name, selection)
                elif(self.first_name.get() == ''):
                    self.insert_into_widget(sel_without_comma, self.first_name)
                    self.delete_selection(self.full_name, selection)
                else:
                    self.insert_middle_name(sel_without_comma, self.middle_name)
                    self.delete_selection(self.full_name, selection)
            elif(selection != " "): # else as long as the selection isn't as empty space
                if(self.last_name.get() == ''):
                    self.insert_into_widget(selection, self.last_name)
                    self.delete_selection(self.full_name, selection)
                elif(self.first_name.get() == ''):
                    self.insert_into_widget(selection, self.first_name)
                    self.delete_selection(self.full_name, selection)
                else:
                    self.insert_into_widget(selection, self.middle_name)
                    self.delete_selection(self.full_name, selection)
                
                self.full_name.delete((len(self.full_name.get())) - 1)
                
    # deletes the selection from the vendor field
    def delete_selection(self, widget, selection):
        widget_names = widget.get().split(" ")
        widget.delete(0, tk.END)
        for word in widget_names:
            if word != selection:
                widget.insert(tk.END, word + " ")
                
    # in case the user makes a mistake, right click will reset the row in the popup window
    def reset(self, event):
        self.last_name.delete(0, tk.END)
        self.first_name.delete(0, tk.END)
        self.middle_name.delete(0, tk.END)
        self.full_name.delete(0, tk.END)
        self.full_name.insert(0, self.cell)
        
    # inserts text into the specified widget
    def insert_into_widget(self, selection, widget):
        widget.insert(0, selection)
        
app = App()

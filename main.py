# Calls the function that creates the GUI, sets up all of the button and has the mainloop
import tkinter as tk
import os
import openpyxl as xl

import addEntry
import newMonth
import info
import createGUI

# Note on the notation of retrieving values from excel
# sheet[row][col] BUT COLUMN is zero indexed while ROW is NOOOOTTTTT

# Simply open the Expenses excel file
def open_excel():
    # TestBudget.xlsx or whatever file needs to be in the same directory to work
    os.startfile(info.excelFile)
    os.startfile(info.excelFile)
    info.window.destroy()

# Set up and display GUI
info.window.title("Budget GUI - " + info.month)
createGUI.create_GUI(info.monthSheetData)

# Make the buttons
# Add Entry Button
tk.Button(info.window, text="New Entry", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen", activebackground = "darkolivegreen", command=addEntry.add_entry).grid(row=12, column = 0, sticky=tk.W, padx=5, pady=5)

# New Month button
tk.Button(info.window, text="New Month", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen", activebackground = "darkolivegreen", command=newMonth.new_month).grid(row=12, column=1, sticky=tk.W, padx=5, pady=5)

# Open Excel Button
tk.Button(info.window, text="Open Excel", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen", activebackground = "darkolivegreen", command=open_excel).grid(row=12, column=2, sticky=tk.W, padx=5, pady=5)

# Refresh Button
tk.Button(info.window, text="Refresh", font = "Calibri 12 bold", relief = 'groove', bg = "mediumseagreen", activebackground = "darkolivegreen", command=createGUI.updateGUI).grid(row=12, column=3, sticky=tk.W, padx=5, pady=5)

info.window.mainloop()


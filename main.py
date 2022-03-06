# Calls the function that creates the GUI, sets up all of the button and has the mainloop
# Also contains the open_excel() function

try:

    import os
    import tkinter as tk

    import addEntry
    import newMonth
    import info
    import createGUI
    import updateCategories

except Exception as e:
    print('Can not import files:' + str(e))
    input("Press Enter to exit!")
    os.exit(0)

# Note on the notation of retrieving values from excel
# sheet[row][col] BUT COLUMN is zero indexed while ROW is NOT
# sheet.cell(row = row, column = column) column is not zero indexed when using this method

# Simply open the Expenses excel file
def open_excel():
    # TestBudget.xlsx or whatever file needs to be in the same directory to work
    os.startfile(info.excelFile)
    info.window.destroy()


# Set up and display GUI
info.window.title("Budget GUI - " + info.month + " - " + info.excelFile)
numRows = createGUI.create_GUI(info.monthSheetData)
numRows += 1

# Make the buttons

# Add Entry Button
addEntry = tk.Button(info.window, text="New Entry", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen",
          activebackground = "darkolivegreen", command=addEntry.createEntryWindow).grid(row=numRows, column = 0, sticky=tk.W, padx=5, pady=5)

# New Month button
tk.Button(info.window, text="New Month", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen",
          activebackground = "darkolivegreen", command=newMonth.new_month).grid(row=numRows, column=1, sticky=tk.W, padx=5, pady=5)

# Open Excel Button
tk.Button(info.window, text="Open Excel", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen",
          activebackground = "darkolivegreen", command=open_excel).grid(row=numRows, column=2, sticky=tk.W, padx=5, pady=5)

# Refresh Button
tk.Button(info.window, text="Add Category", font = "Calibri 12 bold", relief = 'groove', bg = "mediumseagreen",
          activebackground = "darkolivegreen", command=updateCategories.createUpdateWindow).grid(row=numRows, column=3, sticky=tk.W, padx=5, pady=5)

info.window.mainloop()

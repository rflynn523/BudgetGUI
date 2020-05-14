# Responsible for getting the user inputs and correctly adding it to the entry table
import tkinter as tk
from tkinter import ttk
import openpyxl as xl
import info
import createGUI
import datetime

# This function makes the labels and button and shows the window to the user
def createEntryWindow():
    clear()

    # Labels for the new entry window.
    inputLabels = ["Date:", "Item:", "Vendor:", "Amount:", "Category:"]

    for label in inputLabels:
        tk.Label(addWindow, text=label).grid(row=inputLabels.index(label), column=0, padx=5, pady=5)

    # Submit Button
    tk.Button(addWindow, text="Submit", font="Calibri 12 bold", relief='groove', bg="mediumseagreen",
              activebackground="darkolivegreen", command=submit).grid(row=10, column=0, columnspan=2,
                                                                                   sticky=tk.S + tk.W + tk.E, padx=5,
                                                                                   pady=5)
    # Show the window
    addWindow.update()
    addWindow.deiconify()
    addWindow.geometry('+600+100')
    addWindow.title("New Entry")

# Write the data to the Excel file after the submit button is pressed
def submit():
    # Reload the Monthly Data sheet to allow for multiple enrtries
    newSheet = xl.load_workbook(info.excelFile, data_only=True)["Monthly"]

    # Get the next open row in the entry table
    openRow = info.getRowNum(newSheet, 24, 1)

    # Loop through the entries list and handle each column accordingly
    for col in range(len(entries)):
        input = entries[col].get()

        # Date format
        if (col == 0):
            if(input == "Today"):
                info.monthSheetEq[openRow][col].value = datetime.datetime.now().date()
            else:
                month,day, year = input.split("/")

                input = datetime.datetime((int)(year), (int)(month), (int)(day))
                info.monthSheetEq[openRow][col].value = input.date()

            info.monthSheetEq[openRow][col].number_format = info.dateFormat

        # Accounting format
        elif(col == 3):
            try:
                info.monthSheetEq[openRow][col].value = (float)(input)
                info.monthSheetEq[openRow][col].number_format = info.accountingFormat
            except:
                createGUI.displayMessage(input + " is not a numeric value! Change the amount entered in Excel")

        # Handle the abbreviated categories
        elif(col == 4):
            if(input == "Entertainment"):
                input = "ENT"
            elif(input == "Eating Out"):
                input = "EO"

            info.monthSheetEq[openRow][col].value = input

        # Remaining
        else:
            info.monthSheetEq[openRow][col].value = input

    # Save the file
    try:
        info.wbEq.save(info.excelFile)
        # Inform the user
        createGUI.displayMessage(itemEntry.get() + " was added to the Entry Table on Monthly. \n\n"
                                                   "Due to limitations of Openpyxl you need to go \n in and save the excel file"
                                                   " in order for \n the main window to properly display the new amounts.")

        # Close the window
        addWindow.withdraw()

    except:
        createGUI.displayMessage("The Excel workbook is already open, the entry was not saved")

    # Close the file
    info.wbEq.close()

    # Update the main window
    createGUI.updateGUI()

# Resets the entry fields
def clear():
    dateEntry.set("Today")
    for a in range(1, len(entries) - 1):
        entries[a].delete(0, "end")

    categoryEntry.set("Select")


# Window
addWindow = tk.Toplevel(info.window)
addWindow.withdraw()

# Entries
dateEntry = ttk.Combobox(addWindow, values=["Today", "Mon/Day/2020"], width=15, justify='center', font="Calibri 11")
dateEntry.grid(row=0, column=1, padx=5, pady=5)

itemEntry = tk.Entry(addWindow, justify='center', font="Calibri 11", width=18)
itemEntry.grid(row=1, column=1, padx=5, pady=5)

vendorEntry = tk.Entry(addWindow, justify='center', font="Calibri 11", width=18)
vendorEntry.grid(row=2, column=1, padx=5, pady=5)

amountEntry = tk.Entry(addWindow, justify='center', font="Calibri 11", width=18)
amountEntry.grid(row=3, column=1, padx=5, pady=5)

categoryEntry = ttk.Combobox(addWindow, values=info.categoryList, width=15, justify='center', state='readonly', font="Calibri 11")
categoryEntry.set('Select')
categoryEntry.grid(row=4, column=1, padx=5, pady=5)

entries = [dateEntry, itemEntry, vendorEntry, amountEntry, categoryEntry]


# Main RN creates the GUI and buttons and has the Open Excel function
import tkinter as tk
import os

import addEntry
import newMonth
import info

# Note on the notation of retrieving values from excel
# sheet[row][col] BUT COLUMN is zero indexed while ROW is NOOOOTTTTT

def create_GUI():
    sheet = info.monthSheetData

    budgetCell = sheet[16][0]
    spendingCell = sheet[16][2]
    perDayCell = sheet[19][3]

    # List of strings that holds the text
    textList = ["Budget Set At:", "Spending Money:", "Remaining Per Day:"]
    amountList = [budgetCell, spendingCell, perDayCell]

    colors = ["yellow", 'lightgreen', 'seagreen']

    # Labels for the first three (fixed) rows
    for x in range(len(textList)):
        # Actual label
        tk.Label(window, text=textList[x], font="Calibri 12 bold").grid(row=x, column=0,columnspan=2, sticky=tk.W, padx=5, pady=5)

        text = clean_values(amountList[x])
        if(text[2] == "-"):
            color = "orangered"
        else:
            color = colors[x]
        tk.Label(window, text=text, font="Calibri 12", relief='solid', bg = color, width=20).grid(row=x, column=2, columnspan=2, sticky=tk.E, padx=5, pady=5)

    create_category_table(info.monthSheetData)

# Function that creates and displays all the categories and there current amounts.
def create_category_table(sheet):

    # Initial row and column values
    # cell = sheet[row][col], 0 indexed
    row = 3
    col = 1

    # Loops through the category names
    # Starts at the rent cell
    # while sheet[row][col] is not None: ADD LATER
    for x in range(8):
        tk.Label(window, text=sheet[row][col].value, font='Calibri 12').grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        row += 1

    # Row and col are now moved to the amounts column
    row = 3
    col = 2

    # Loops through the category values
    for x in range(8):
        tk.Label(window, text = clean_values(sheet[row][col]), font='Calibri 12', relief='groove', bg='cyan', width=20).grid(row=row, column=2, columnspan=2, sticky=tk.E, padx=5, pady=5)
        row += 1

# This function is used to add dollar signs to money amounts
# along with rounding to two decimal places
def clean_values(cell):
    # Cast the full value to a string
    cleanString = str(cell.value)

    # Checks if value has a decimal place and adds one if it does not
    if("." not in cleanString):
        cleanString = cleanString + ".00"

    # Shaves off long decimals and rounds to two decimal places
    else:
        number = cleanString.split(".")
        whole = number[0]
        decimal = number[1]

        if(len(decimal) > 2):
            third = int(decimal[2])
            if(third >= 5):
                roundedUp = int(decimal[0] + decimal[1])
                roundedUp += 1
                cleanString = whole + "." + str(roundedUp)
            else:
                roundedDown = decimal[:2]
                cleanString = whole + "." + str(roundedDown)

    # Check for the dollar sign
    if(cleanString[0] != "$"):
        cleanString = "$ " + cleanString

    return cleanString

# Simply open the Expenses excel file
def open_excel():
    # TestBudget.xlsx or whatever file needs to be in the same directory to work
    os.startfile(info.excelFile)
    os.startfile(info.excelFile)
    window.destroy()

# Updates the main GUI values
def refresh():
    pass
    # Call the updateValues function



# Set up and display GUI
window = tk.Tk()
window.title("Budget GUI")
create_GUI()

# Buttons
# Add Item Button
tk.Button(window, text="New Entry", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen", activebackground = "darkolivegreen", command=addEntry.add_entry).grid(row=12, column = 0, sticky=tk.W, padx=5, pady=5)

# New Month button
tk.Button(window, text="New Month", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen", activebackground = "darkolivegreen", command=newMonth.new_month).grid(row=12, column=1, sticky=tk.W, padx=5, pady=5)

# Open Excel Button
tk.Button(window, text="Open Excel", font = "Calibri 12 bold", relief = 'groove', bg="mediumseagreen", activebackground = "darkolivegreen", command=open_excel).grid(row=12, column=2, sticky=tk.W, padx=5, pady=5)

# Refresh Button
tk.Button(window, text="Refresh", font = "Calibri 12 bold", relief = 'groove', bg = "mediumseagreen", activebackground = "darkolivegreen", command=refresh).grid(row=12, column=3, sticky=tk.W, padx=5, pady=5)

window.mainloop()


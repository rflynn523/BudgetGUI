# Creates the GUI and holds the updateGUI function
import tkinter as tk
import info
import openpyxl as xl

# Create the GUI by reading from the monthSheetData that is passed from main.py
def create_GUI(sheet):
    budgetCell = sheet[16][0]
    spendingCell = sheet[info.getRowNum(sheet, 2, 1, "Spending Money")][2]
    perDayCell = sheet[19][3]

    # List of strings that holds the text
    textList = ["Budget Set At:", "Spending Money:", "Remaining Per Day:"]
    amountList = [budgetCell, spendingCell, perDayCell]

    colors = ["yellow", 'lightgreen', 'seagreen']

    # Labels for the first three (fixed) rows
    for x in range(len(textList)):
        # Actual label
        tk.Label(info.window, text=textList[x], font="Calibri 12 bold").grid(row=x, column=0,columnspan=2, sticky=tk.W, padx=5, pady=5)

        text = clean_values(amountList[x])
        if(text[2] == "-"):
            color = "orangered"
        else:
            color = colors[x]
        tk.Label(info.window, text=text, font="Calibri 12", relief='solid', bg = color, width=20).grid(row=x, column=2, columnspan=2, sticky=tk.E, padx=5, pady=5)

    create_category_table(sheet)

# Function that creates and displays all the categories and there current amounts.
def create_category_table(sheet):
    # Find the last row in the monthly category table
    row = 3
    col = 2
    lastRow = info.getRowNum(sheet, row, 1)
    # monthCell = sheet[row][2]

    # Loops through the category names
    # Starts at the rent cell
    for x in range(lastRow - row):
        tk.Label(info.window, text=sheet[row][1].value, font='Calibri 12').grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        tk.Label(info.window, text = clean_values(sheet[row][col]), font='Calibri 12', relief='groove', bg='cyan', width=20).grid(row=row, column=2, columnspan=2, sticky=tk.E, padx=5, pady=5)
        row += 1

# This function is used to add dollar signs to money amounts
# along with rounding to two decimal places
def clean_values(cell):
    # Cast the full value to a string
    if(cell.value == None):
        cleanString = "0"

    else:
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
            # Get the third deciaml place
            third = int(decimal[2])

            if(third >= 5):
                roundedUp = int(decimal[0] + decimal[1])
                roundedUp += 1
                cleanString = whole + "." + str(roundedUp)
            else:
                roundedDown = decimal[:2]
                cleanString = whole + "." + str(roundedDown)

        elif (len(decimal) < 2):
            cleanString = whole + "." + decimal + "0"

    # Check for the dollar sign
    if(cleanString[0] != "$"):
        cleanString = "$ " + cleanString

    return cleanString

# Updates the main GUI values
def updateGUI():
    # Re-load the workbook
    newWbData = xl.load_workbook(info.excelFile, data_only=True)

    # Re-create the GUI
    create_GUI(newWbData['Monthly'])
# Creates the GUI and holds the updateGUI function
import tkinter as tk
import info
import openpyxl as xl

# Create the GUI by reading from the monthSheetData that is passed from main.py
def create_GUI(sheet, categoryList=None):
    budgetCell = sheet[18][0]
    spendingCell = sheet[info.getRowNum(sheet, 2, 1, "Spending Money")][2]
    perDayCell = sheet[19][3]

    # List of strings that holds the text
    textList = ["Budget Set At:", "Spending Money:", "Remaining Per Day:"]
    amountList = [budgetCell, spendingCell, perDayCell]

    colors = ["yellow", 'lightgreen', 'seagreen']

    # Labels for the first three (fixed) rows
    for x in range(1, len(textList)+1):
        # Actual label
        tk.Label(info.window, text=textList[x-1], font="Calibri 12 bold").grid(row=x, column=0,columnspan=2,
                                                                             sticky=tk.W, padx=5, pady=5)

        text = clean_values(amountList[x-1])
        if(text[2] == "-"):
            color = "orangered"
        else:
            color = colors[x-1]

        # Add the amount
        tk.Label(info.window, text=text, font="Calibri 12", relief='solid', bg = color,
                 width=20).grid(row=x, column=2, columnspan=2, sticky=tk.E, padx=5, pady=5)

    # Make the rest of the table for the non-fixed categories
    numRows = create_category_table(sheet, categoryList)

    return numRows

# Function that creates and displays all the categories and there current amounts.
def create_category_table(sheet, categoryList=None):
    row = 4

    if(categoryList == None):
        categoryList = info.categoryList

    allLabels = []

    # Loops through the category names
    for cat in categoryList:
        # Create Label for the name
        nameLabel = tk.Label(info.window, text=cat, font='Calibri 12')
        nameLabel.grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

        # Add the amount associated with that category
        amountLabel = tk.Label(info.window, text = clean_values(sheet[row-1][2]), font='Calibri 12', relief='groove', bg='cyan',
                 width=20)
        amountLabel.grid(row=row, column=2, columnspan=2, sticky=tk.E, padx=5, pady=5)

        row += 1
        allLabels.append(nameLabel)
        allLabels.append(amountLabel)

    # Return number of rows to format buttons correctly
    return row

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
    newWbData = xl.load_workbook(info.excelFilePath, data_only=True)

    # Update the Display Month
    info.window.title("Budget GUI - " + info.month + " " + info.excelFileName)

    # Loops through the category names starting with the rent cell
    categoryList = []
    # Load in the different categories
    row = 3
    col = 2
    lastRow = info.getRowNum(newWbData["Monthly"], row, 1)

    for x in range(lastRow - row):
        categoryList.append(newWbData["Monthly"][row][1].value)
        row += 1

    # Re-create the GUI
    create_GUI(newWbData['Monthly'], categoryList)

# Simply used to display error messages to the user
def displayMessage(message):
    errorWindow = tk.Toplevel(info.window)
    errorWindow.geometry('+925+300')
    errorWindow.title("Error Message")

    tk.Label(errorWindow, text=message).grid(row=0, column=0, padx=5, pady=5)

    tk.Button(errorWindow, text="Ok", font="Calibri 12 bold", relief='groove', bg="mediumseagreen",
              activebackground="darkolivegreen", command=errorWindow.destroy).grid(row=10, column=0, columnspan=2,
                                                                      sticky=tk.S + tk.W + tk.E, padx=5,
                                                                      pady=5)


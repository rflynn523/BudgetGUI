# Contains information and some helper functions that almost every file needs.

import openpyxl as xl
import tkinter as tk
import createGUI

# Helper Function
# If key is not passed then the function returns the open row number
# Otherwise it returns the row that contains the key
def getRowNum(sheet, startRow, startCol, key=None, month=None):
    # Simply return the next open row number with the given params
    if(key == None and month==None):
        while(sheet[startRow][startCol].value != None):
            startRow += 1
            if(startRow == 100):
                break

    # Return the row where the cell is equal to the key
    elif(month == None):
        # Makes sure program doesnt crash if string is handled wrong
        if(type(key) == str):
            while (str(sheet[startRow][startCol].value) != key):
                startRow += 1
                if (startRow == 100):
                    break
        else:
            while (sheet[startRow][startCol].value != key):
                startRow += 1
                if (startRow == 100):
                    break

    # Used for the Data Set Check
    # Returns the row when the date is equal to the month (passed in as a number)
    else:
        try:
            while ((sheet[startRow+1][startCol].value).month != month):
                startRow += 1
                if (startRow == 100):
                    break
        except:
            createGUI.displayMessage("The data from the Add entry button was not saved correctly")

    return startRow

# Window variable
window = tk.Tk()

# Get info from the config file
config = open(r"BudgetGuiConfig.txt", "r")
month = str(config.readline()).strip('\n')
excelFile = config.readline()
config.close()

# Formatting and other info
accountingFormat = r'_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)'
dateFormat = 'dd-mmm'

# Create the dictionary to map months to cells in the form of:
#  {"Month" : [row, col]}
months = ["Janurary", "Feburary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# Cells correspond to the above months and contain the first cell under 'Amount'
cells = [[5,3], [5,6], [5,9], [17,3], [17,6], [17,9], [29,3], [29,6], [29,9], [41,3], [41,6], [41,9]]

yearly_month_cells = {k:v for k,v in zip(months, cells)}

# Load the workbooks
wbData = xl.load_workbook(excelFile, data_only=True)
wbEq = xl.load_workbook(excelFile, data_only=False)

# Get the needed sheets
monthSheetData = wbData['Monthly']
yearSheetData = wbData['Yearly']
dataSetSheetData = wbData["Data Set"]

monthSheetEq = wbEq["Monthly"]
yearSheetEq = wbEq["Yearly"]
dataSetSheetEq = wbEq["Data Set"]

# Load in the different categories
row = 3
col = 2
lastRow = getRowNum(monthSheetEq, row, 1)

categoryList = []

# Loops through the category names starting with the rent cell
for x in range(lastRow - row):
    categoryList.append(monthSheetData[row][1].value)
    row += 1

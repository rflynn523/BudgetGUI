# Contains information and some helper functions that almost every file needs.

import openpyxl as xl
import tkinter as tk
import createGUI
from openpyxl.styles import Font, Fill, PatternFill, Border, Side

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
            createGUI.displayMessage("StartRow error? info.py line 41")

    return startRow

# Function to save a Back up file before the script makes changes to the main file
def createBackUpFile(newFileName):
    wbEq.save(newFileName)

def getCurrentMonth():
    pass
    # Read the current month from the excel file cell

# Window variable
window = tk.Tk()

# Get info from the config file

# Won't need the Config file when I implement the switch file function
config = open(r"BudgetGuiConfig.txt", "r")
excelFile = str(config.readline()).strip('\n')
config.close()

# Formatting and other info
accountingFormat = r'_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)'
dateFormat = 'dd-mmm'
noFill = PatternFill(fill_type=None)

side = Side(border_style=None)
noBorder = Border(
    left=side,
    right=side,
    top=side,
    bottom=side,
)

months = ["January", "Feburary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# Cells correspond to the above months and contain the first cell under 'Amount' on the Yearly sheet
# Updated the cells 5/2
ApartmentCells = [[5,3], [5,6], [5,9], [17,3], [17,6], [17,9], [29,3], [29,6], [29,9], [41,3], [41,6], [41,9]]
ExpenseCells = [[4,3], [4,6], [4,9], [21,3], [21,6], [21,9], [36,3], [36,6], [36,9], [52,3], [52,6], [52,9]]

# The two main excel sheets are Apartment and Expenses
# This handles the different yearly starting cells
# Apartment Worksheet
if(excelFile.strip('.xlsx') == "Apartment"):
    cells = ApartmentCells

# Expenses Worksheet
else:
    cells = ExpenseCells

# Create the dictionary to map months to cells in the form of:
#  {"Month" : [row, col]}
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

month = monthSheetData[23][0].value

categoryList = []

# Loops through the category names starting with the rent cell
# Load in the different categories
row = 3
col = 2
lastRow = getRowNum(monthSheetEq, row, 1)

for x in range(lastRow - row):
    categoryList.append(monthSheetData[row][1].value)
    row += 1

# File that runs the checks after New Month is pressed and shows the results to the user
import tkinter as tk
import openpyxl as xl
import info, newMonth

# Adds all the labels to the GUI and then calls all the check functions
def makeAllChecks():
    # Current Checks that are going to run
    checkList = ["Monthly Amounts Table Reset", "Current Month Check on Monthly", "Amounts Entered into Yearly",
                 "Totals entered into Yearly", "Category Totals reset on Monthly", "Groc and Gas Tables Updated",
                 "Entry Table Emptied", "Entries entered into Data Set"]

    # Add those labels to the new window
    for x in range(len(checkList)):
        tk.Label(checkWindow, text=checkList[x] + ":", font="Calibri 12").grid(row=x, column=0,columnspan=1, sticky=tk.W, padx=5, pady=5)

    # Relaod the workbooks
    wbData = xl.load_workbook(info.excelFile, data_only=True)
    wbEq = xl.load_workbook(info.excelFile, data_only=False)

    # Get the needed sheets
    monthSheetData = wbData['Monthly']
    yearSheetData = wbData['Yearly']
    dataSetSheetData = wbData["Data Set"]

    monthSheetEq = wbEq["Monthly"]
    yearSheetEq = wbEq["Yearly"]
    dataSetSheetEq = wbEq["Data Set"]

    # Make all the checks
    # Check that the amounts table on monthly is 0 for non fixed categories
    monthlyTableReset(monthSheetData)

    # Check that the correct month is displayed on Monthly
    currentMonthCheck(monthSheetData)

    # Check that the amounts were entered into the correct month on Yearly
    amountsIntoYearly(monthSheetData, yearSheetData)

    # Check that the 'old' months Total Spent, Total(Besides R/P), and NET values are
    # inserted to the table on Yearly.
    totalsIntoYearly(yearSheetData)

    # Check that the 'new' month cells are the actual equations
    monthlyTotalsReset(monthSheetEq)

    # Check that the grocery and gas tables were updated correctly????
    grocGasCheck(monthSheetEq)

    # Make sure the Entry table is empty
    emptyEntryTable(monthSheetData)

    # Make sure the data was inserted into the Data Set sheet with the MONTH() formula
    dataSetEntered(dataSetSheetData, yearSheetData)

    # Finally show the new window to the user
    showGUI()

# Check that the amounts table on monthly is 0 for non fixed categories
def monthlyTableReset(monthSheetData):
    row = 5
    col = 2
    # Find the last row in the monthly category table
    lastRow = info.getRowNum(monthSheetData, row, 1)
    monthlyAmounts = []

    # Loops through the amounts names
    # Starts after the fixed values
    for x in range(lastRow - row):
        monthlyAmounts.append(monthSheetData[row][col].value == None)
        row += 1

    # All values should be None for each category which
    if(False in monthlyAmounts):
        addResult(False, 0)
    else:
        addResult(True, 0)

# Check that the correct month is displayed on Monthly
def currentMonthCheck(monthSheetData):
    monthCheck = (info.month == monthSheetData.cell(row = 23, column = 1).value)
    addResult(monthCheck, 1)

# Check that the amounts were entered into the correct month on Yearly
def amountsIntoYearly(monthSheetData, yearSheetData):
    # Get the now previous month index
    prevMonth = info.months.index(info.month) - 1
    monthCellList = newMonth.getMonthStartCell(info.months[prevMonth])

    # Split into row and column
    monthRow, monthCol = monthCellList[0], monthCellList[1]

    # Get the number of categories
    row = 3
    lastRow = info.getRowNum(monthSheetData, row, 1)

    matchesMonthlyAmounts = []

    # Loop through the data on yearly
    # None corresponds to no data in the cell
    for x in range(lastRow - row):
        matchesMonthlyAmounts.append(yearSheetData[monthRow][monthCol].value == None)
        row += 1

        # If a false is found return right away
        if (False in matchesMonthlyAmounts):
            addResult(False, 2)
            return

    addResult(True, 2)

# Check that the 'old' months Total Spent, Total(Besides R/P), and NET values are
# inserted to the table on Yearly.
def totalsIntoYearly(yearSheetData):
    # With the zero indexed list, info.month is the month number of the previous month
    row = 22 + info.months.index(info.month)

    totalSpent = (yearSheetData.cell(row=row, column=13).value != None)
    totalBesidesPR = (yearSheetData.cell(row=row, column=14).value != None)
    net = (yearSheetData.cell(row=row, column=15).value != None)

    addResult((totalSpent and totalBesidesPR and net), 3)

    # Maybe check that they are reset on monthly too?

# Check that the 'new' month cells are the actual equations
def monthlyTotalsReset(monthSheetEq):
    # Find the last row in the monthly category table
    row = 5
    col = 2
    lastRow = info.getRowNum(monthSheetEq, row, 1)

    monthlyAmounts = []

    # Loops through the amounts names
    # Starts after the fixed values
    for x in range(lastRow - row):
        monthlyAmounts.append("=" in monthSheetEq[row][col].value)
        row += 1

        # Breaks out right away if true
        if (False in monthlyAmounts):
            addResult(False, 4)
            return

    addResult(True, 4)

# Chcek that the grocery and gas tables were updated correctly
def grocGasCheck(monthSheetEq):
    # Saves the next open row in the Grocery and Gas tables
    current = info.getRowNum(monthSheetEq, 31, 10) - 1

    # Get the Equations
    grocTotalEq = monthSheetEq[current][9].value
    grocAvgEq = monthSheetEq[current][10].value
    gasTotalEq = monthSheetEq[current][13].value
    gasAvgEq = monthSheetEq[current][14].value

    # Get the Data
    grocTotalData = monthSheetEq[current-1][9].value
    grocAvgData = monthSheetEq[current-1][10].value
    gasTotalData = monthSheetEq[current-1][13].value
    gasAvgData = monthSheetEq[current-1][14].value

    # The old cell with the equation should be overwritten by the value which is a float type
    grocDataCheck = (type(grocTotalData) == float or type(grocAvgData) == float)
    gasDataCheck = (type(gasTotalData) == float or type(gasAvgData) == float)

    # The new cell should now contain the equation
    grocEquationCheck = ("=" in grocTotalEq or "=" in grocAvgEq)
    gasEquationCheck = ("=" in gasTotalEq or "=" in gasAvgEq)

    addResult((grocDataCheck or gasDataCheck or grocEquationCheck or gasEquationCheck), 5)

# Make sure the Entry table is empty
def emptyEntryTable(monthSheetData):
    addResult((info.getRowNum(monthSheetData, 25, 1) == 25), 6)

 # Make sure the data was inserted into the Data Set sheet
def dataSetEntered(dataSetSheetData, yearSheetData):
    # Get the first row number that has the index of the month
    if (info.month == "January"):
        firstRow = info.getRowNum(dataSetSheetData, 3, 3, month=12) + 1
    else:
        firstRow = info.getRowNum(dataSetSheetData, 3, 4, month=info.months.index(info.month)) + 1

    # Get the last row number
    lastRow = info.getRowNum(dataSetSheetData, firstRow, 4) - 1

    # Add up the total from each entry
    sum = 0
    num = lastRow - firstRow
    for x in range(num+1):
        sum += dataSetSheetData.cell(row=firstRow, column = 8).value
        firstRow += 1

    # Find the previous Total (Besides R/P)
    # prevMonth = info.months[info.months.index(info.month) - 1]

    row = 23 + (info.months.index(info.month) - 1)
    totalBesidesPR = yearSheetData.cell(row=row, column=14).value

    addResult((round(sum,2) == totalBesidesPR), 7)

# Helper funciton that adds the results to the pop-up window
def addResult(result, num):
    if result:
        text = "Success"
    else:
        text = "Fail"

    # Add the result label to the windowS
    tk.Label(checkWindow, text=text, font="Calibri 12 bold").grid(row=num, column=1, columnspan=1, sticky=tk.W,
                                                                      padx=5, pady=5)

# After all the checks have been complete, show the window to the user
def showGUI():
    # Add Close button that just closes the window
    tk.Button(checkWindow, text="Close", font="Calibri 12 bold", relief='groove', bg="mediumseagreen",
              activebackground="darkolivegreen", command=checkWindow.destroy).grid(row=10, column=0, columnspan=2, sticky=tk.S+tk.W+tk.E, padx=5, pady=5)

    checkWindow.update()
    checkWindow.deiconify()
    checkWindow.geometry('+600+100')
    checkWindow.title("New Month Check")

# Window
checkWindow = tk.Toplevel(info.window)
checkWindow.withdraw()

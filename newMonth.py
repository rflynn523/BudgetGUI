# Handles all of the necessary steps and functions need to reset the spreadsheet for a new month
import info
import createGUI
import newMonthCheck

def new_month():
    # Copy the VALUES from the monthly Category Table to the corresponding table in Yearly
    copySummaryTable(info.monthSheetData, info.yearSheetEq, info.categoryList)

    # Also copy the Total Spent, Total Besides P/R, and Net values from Monthly to Yearly
    copyTotalValues(info.monthSheetData, info.yearSheetData, info.yearSheetEq)

    # Get the data from the Entries table from Monthly, write that data to the 'Data Set'
    # sheet and then clear the Entries table
    updateEntryTable(info.monthSheetData, info.monthSheetEq, info.dataSetSheetEq)

    # # Update the month in the BudgetGuiConfig file
    try:
        updateMonth(info.months[info.months.index(info.month) + 1])
    except:
        createGUI.displayMessage("Month is December, check the current months")

    # Save only the EQUATIONS workbook file
    try:
        info.wbEq.save(info.excelFile)
    except:
        createGUI.displayMessage("Close the excel file, check if it is saved correctly")

    # Perform all of the checks
    newMonthCheck.makeAllChecks()

    # Call the updateValues function to update the GUI once its complete
    createGUI.updateGUI()

# Copy only the values from the summary table to the "Yearly" Sheet
def copySummaryTable(monthSheetData, yearSheetEq, categoryList):
    # Initialize some needed variables
    row = 3
    amountsList = []
    monthCell = monthSheetData[row][2]

    # Loop down the amounts column, saving all of the numbers
    for l in range(len(categoryList)):
        # Clean values are not needed because excel needs to see the values as numbers not strings
        amountsList.append(monthCell.value)
        row += 1
        monthCell = monthSheetData[row][2]

    # Write to the yearly sheet of the EQUATIONS file for the current month according to the config file
    monthCellList = getMonthStartCell()
    row, col = monthCellList[0], monthCellList[1]

    # Write in the data
    for amt in amountsList:
        yearSheetEq.cell(row = row, column = col).value = amt
        row += 1

# Copy Total Spent, Total(Besides R/P), and NET cells to the yearly sheet
def copyTotalValues(monthSheetData, yearSheetData, yearSheetEq):
    # Get the data from the DATA workbook
    totalSpent = monthSheetData[info.getRowNum(monthSheetData, 2, 1, "Total")][2].value
    totBesidesRP = monthSheetData[info.getRowNum(monthSheetData, 2, 1, "Total (Besides R/P)")][2].value
    net = monthSheetData[info.getRowNum(monthSheetData, 2, 1, "Spending Money")][2].value

    # Find the next open row in the yearly table
    nextOpen = info.getRowNum(yearSheetData, 22, 12)

    # Writing data to EQUATION file
    yearSheetEq.cell(row=nextOpen, column=13).value = totalSpent
    yearSheetEq.cell(row=nextOpen, column=14).value = totBesidesRP
    yearSheetEq.cell(row=nextOpen, column=15).value = net

# Cut the data from the Monthly Entries table and paste it into the Data Set sheet
def updateEntryTable(monthSheetData, monthSheetEq, dataSetSheet):
    # Get the last row number of the data
    lastRow = info.getRowNum(monthSheetData, 24, 1) - 1
    dataList = []

    # Get each data row into a list and then add that list to dataList
    for i in range(25, lastRow + 1):
        dataRow = []
        for j in range(1, 6):
            dataRow.append(monthSheetData.cell(row = i, column = j).value)

        dataList.append(dataRow)

    nextOpen = info.getRowNum(dataSetSheet, 3, 4)

    # Copy the dataList into the Data Set sheet and insert the MONTH() formula, row by row
    for d in range(nextOpen, nextOpen + len(dataList)):
        temp = dataList[d - nextOpen]
        dataSetSheet.cell(row=d, column=4).value = "=MONTH(E" + str(d) + ")"
        for c in range(5, 10):
            dataSetSheet.cell(row = d, column = c).value = temp[c - 5]

            # Date format
            if (c == 5):
                dataSetSheet.cell(row=d, column=c).number_format = info.dateFormat

            # Money format
            if (c == 8):
                dataSetSheet.cell(row=d, column=c).number_format = info.accountingFormat

    # Clear the entries table from the "Monthly" sheet
    for i in range(25, lastRow + 1):
        for j in range(1, 6):
            monthSheetEq.cell(row = i, column = j).value = None

# Writes the newMonth to the config file and also updates the month from info.py
def updateMonth(newMonth):
    # Read everything from the text file
    with open(r"BudgetGuiConfig.txt", 'r') as oldFile:
        data = oldFile.readlines()

    # Change the first line containing the month
    data[0] = newMonth + '\n'

    # Save the second line (the excel file name)
    data[1] = info.excelFile

    # Write back all of the data to config file
    with open(r'BudgetGuiConfig.txt', 'w') as newFile:
        newFile.writelines(data)

    # Update program month
    info.month = newMonth

    # Write the new month to the monthly cell
    info.monthSheetEq.cell(row = 23, column = 1).value = newMonth

# Helper function that determines the yearly month starting cell by using month
# and yearly_month_cells from info.py
def getMonthStartCell(month=None):
    if (month == None):
        return info.yearly_month_cells.get(info.month)
    else:
        return info.yearly_month_cells.get(month)

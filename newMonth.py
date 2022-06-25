# Handles all of the necessary steps and functions need to reset the spreadsheet for a new month
import info
import createGUI
import newMonthCheck

def new_month():
    # Save a backup
    info.createBackUpFile("NewMonth_Backup_" + info.excelFile)

    # Copy the VALUES from the monthly Category Table to the corresponding table in Yearly
    copySummaryTable(info.monthSheetData, info.yearSheetEq, info.categoryList)

    # Also copy the Total Spent, and Net values from Monthly to Yearly
    copyTotalValues(info.monthSheetData, info.yearSheetData, info.yearSheetEq)

    # # Get the data from the Entries table from Monthly, write that data to the 'Data Set'
    # # sheet and then clear the Entries table
    # updateEntryTable(info.monthSheetData, info.monthSheetEq, info.dataSetSheetEq)

    # # Update the month in the BudgetGuiConfig file
    # if(info.month == "December"):
    #     updateConfigFile(info.months[0])

    # else: 
    #     updateConfigFile(info.months[info.months.index(info.month) + 1])

    # Save only the EQUATIONS workbook file
    try:
        info.wbEq.save(info.excelFile)
    except:
        createGUI.displayMessage("Close the excel file, check if it is saved correctly")

    # # Perform all of the checks
    # newMonthCheck.makeAllChecks()

    # # Call the updateValues function to update the GUI once its complete
    # createGUI.updateGUI()

# Copy the Summary Table from Monthly to Yearly by grabbing both the category and amounts
def copySummaryTable(monthSheetData, yearSheetEq, categoryList):
    # Initialize some needed variables
    row = 2
    namesList = []
    amountsList = []

    nameMonthCell = monthSheetData[row][1]
    amountMonthCell = monthSheetData[row][2]

    # Loop down the category table on the monthly sheet 
    for l in range(len(categoryList) + 1):
        # Clean values are not needed because excel needs to see the values as numbers not strings

        namesList.append(nameMonthCell.value)
        amountsList.append(amountMonthCell.value)

        row += 1

        nameMonthCell = monthSheetData[row][1]
        amountMonthCell = monthSheetData[row][2]

    # Get the next open row in the column of months
    # Always start the search at row 23 and column C
    nextOpenRow = info.getRowNum(yearSheetEq, 23, 2)

    monthTitleCell = nextOpenRow
    categoryColumn = 2
    amountColumn = 3
    
    # Write in the Month name with formatting
    monthCell = yearSheetEq[monthTitleCell][categoryColumn]

    monthCell.value = info.month
    monthCell.fill = info.greenFill
    monthCell.font = info.yearlyMonth

    row = monthTitleCell + 1 

    # Write in the category names and the amounts
    for name, amount in zip(namesList, amountsList):
        currentCategoryCell = yearSheetEq[row][categoryColumn]
        currentAmountCell = yearSheetEq[row][amountColumn]

        # Specific formats for the column titles
        if name == "Category" or amount == "Amount":
            currentCategoryCell.font = info.boldFont
            currentAmountCell.font = info.boldFont

        # Write and format the category name
        currentCategoryCell.value = name
        currentCategoryCell.fill = info.lightGreenFill
        currentCategoryCell.border = info.allBorders

        # Write and format the category's amounr
        currentAmountCell.value = amount
        currentAmountCell.fill = info.lightGreenFill
        currentAmountCell.border = info.allBorders

        row += 1


# Copy Total Spent, Total(Besides R/P), and NET cells to the yearly sheet
def copyTotalValues(monthSheetData, yearSheetData, yearSheetEq):
    # Get the data from the DATA workbook
    totalSpent = monthSheetData[info.getRowNum(monthSheetData, 2, 1, "Total")][2].value
    net = monthSheetData[info.getRowNum(monthSheetData, 2, 1, "Spending Money")][2].value

    # Find the next open row in the yearly table
    nextOpenRow = info.getRowNum(yearSheetData, 4, 3)

    # Writing data to EQUATION file
    yearSheetEq.cell(row=nextOpenRow, column=4).value = totalSpent
    yearSheetEq.cell(row=nextOpenRow, column=5).value = net

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
        for j in range(1, 7):
            monthSheetEq.cell(row = i, column = j).value = None

    # Save the number of entries for checking purposes
    info.numEntries = (lastRow+1) - 25

# Writes the newMonth to the config file and also updates the month from info.py
def updateConfigFile(newMonth):
    # Read everything from the text file
    with open(r"BudgetGuiConfig.txt", 'r') as oldFile:
        data = oldFile.readlines()

    # Write back all of the data to config files
    with open(r'BudgetGuiConfig.txt', 'w') as newFile:
        newFile.writelines(data)

    # Update program month
    info.month = newMonth

    # Write the new month to the monthly cell
    info.monthSheetEq.cell(row = 23, column = 1).value = newMonth

    info.window.title("Budget GUI - " + info.month + " - " + info.excelFile)

# Helper function that determines the yearly month starting cell by using month
# and yearly_month_cells from info.py
def getMonthStartCell(month=None):
    if (month == None):
        return info.yearly_month_cells.get(info.month)
    else:
        return info.yearly_month_cells.get(month)

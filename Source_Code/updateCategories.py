# Holds the functions to create a new category.
# This file will insert the text into all the correct places needed to make a new category
# Need to add:
    # - Formating of the month tables on Yearly insertion
    # - Equations of total/average tables on Yearly insertion/ month shifted rows
    # - Automate the month start cells updating with the shift, on info.py
    # - Optimize the three shifting functions

import info
import tkinter as tk
import createGUI
import newMonth

# Display to the user the current categories
def createUpdateWindow():
    # Simple label
    tk.Label(updateWindow, text="Select the category to insert before.").grid(row=1, column=1, columnspan=2,
                                                                             sticky=tk.S + tk.W + tk.E, padx=5,
                                                                             pady=5)
    # Display all the categories as buttons for selection
    inputLabels = info.categoryList
    for label in inputLabels:
        row = inputLabels.index(label)

        tk.Button(updateWindow, text=label, font="Calibri 12 bold", relief='groove', bg="mediumseagreen",
                  activebackground="darkolivegreen", command=lambda label=label: insert(label)).grid(row = row+2, column=1, columnspan=2,
                                                                          sticky=tk.S + tk.W + tk.E, padx=5,
                                                                          pady=5)

    # Show the user the now fully created window
    updateWindow.update()
    updateWindow.deiconify()
    updateWindow.geometry('+600+100')
    updateWindow.title("New Entry")

# Insert the new category into the correct places
def insert(insertBefore):
    # Save a backup
    info.createBackUpFile("NewMonth_BackUp_" + info.excelFile);

    # Grab the user data
    categoryName = newNameEntry.get()

    # Find out how many rows are under the insertBefore row
    catBelow = len(info.categoryList) - (info.categoryList.index(insertBefore) + 1)

    # First handle the monthly categories down shift and insert on 'Monthly'
    shiftMonthlyCategories(catBelow, categoryName)

    # Insert the new category into all months
    # Shift all the other month rows below down to make room for the insert
    shiftAllMonths(catBelow, categoryName)

    # Do the same thing as shiftMonthlyCategories for the Yearly Totals and Averages table
    shiftYearlyTotalsAverages(catBelow, categoryName)

    # Save the file
    try:
        info.wbEq.save(info.excelFile)
        # Inform the user
        createGUI.displayMessage( str(categoryName) +" was inserted to the Monthly and Yearly sheets\n\n"
                                 "Due to limitations of Openpyxl you need to go \n in and save the excel file"
                                 " in order for \n the main window to properly display the new amounts.")

        # Close the window
        updateWindow.withdraw()

    except:
        createGUI.displayMessage("The Excel workbook is already open, the entry was not saved")

# Handle the monthly categories down shift and insert
def shiftMonthlyCategories(catBelow, categoryName):
    # Get the last row number from the excel file
    lastRow = info.getRowNum(info.monthSheetEq, 3, 1)

    # Figure out the row number in the excel file that needs to be used for insertions
    # The nth row from the 'Categories' cell
    catBelowIndex = lastRow - catBelow - 1

    # Read the category names and possible equations from catBelow to lastRow and store in a list
    row = catBelowIndex
    cutNames = []
    cutAmts = []
    while (row < lastRow + 6):
        cutNames.append(info.monthSheetEq[row][1].value)
        cutAmts.append(info.monthSheetEq[row][2].value)
        row += 1

    # Insert the new category
    info.monthSheetEq[catBelowIndex][1].value = categoryName
    info.monthSheetEq[catBelowIndex][2].value = None

    # Write the categories and the amounts to the cell one row down
    row = catBelowIndex + 1
    ind = 0
    while (row < lastRow + 7):
        info.monthSheetEq[row][1].value = cutNames[ind]
        info.monthSheetEq[row][2].value = cutAmts[ind]

        ind += 1
        row += 1

# Makes the insertion on 'Yearly' for the totals and averages table
def shiftYearlyTotalsAverages(catBelow, categoryName):
    # Get the last row number from the excel file
    lastRow = info.getRowNum(info.yearSheetEq, 6, 11)

    # Figure out the row number in the excel file that needs to be used for insertions
    # The nth row down from the 'Categories' cell
    catBelowIndex = lastRow - catBelow - 1

    # Read the category names and possible equations from catBelow to lastRow and store in a list
    row = catBelowIndex
    cutTotalNames = []
    cutTotalAmts = []
    cutAvgNames = []
    cutAvgAmts = []
    while (row < lastRow + 3):
        cutTotalNames.append(info.yearSheetEq[row][11].value)
        cutTotalAmts.append(info.yearSheetEq[row][12].value)

        cutAvgNames.append(info.yearSheetEq[row][14].value)
        cutAvgAmts.append(info.yearSheetEq[row][15].value)

        row += 1

    # Insert the new category
    # Total
    info.yearSheetEq[catBelowIndex][11].value = categoryName
    info.yearSheetEq[catBelowIndex][12].value = None

    # Average
    info.yearSheetEq[catBelowIndex][14].value = categoryName
    info.yearSheetEq[catBelowIndex][15].value = None

    # Write the categories and the amounts to the cell one row down
    row = catBelowIndex + 1
    ind = 0
    while (row < lastRow + 3):
        info.yearSheetEq[row][11].value = cutTotalNames[ind]
        info.yearSheetEq[row][14].value = cutAvgNames[ind]

        info.yearSheetEq[row][12].value = cutTotalAmts[ind]
        info.yearSheetEq[row][15].value = cutAvgAmts[ind]

        ind += 1
        row += 1

# Update all of the months on the yearly sheet
def shiftAllMonths(catBelow, categoryName):
    # Reverse the months list from info
    reverseMonths = info.months[::-1]

    x = 1
    for r in reverseMonths:
        # Get the actual title cell for each month
        # Returns the cell for the first amount cell, ex (Rent:_____)
        revRow, revCol = findMonthTitleCell(r)
        revCol -= 1

        # Get the last row number from the excel file
        lastRow = info.getRowNum(info.yearSheetEq, revRow, 1)

        # Need a different catBelowIndex because last row is different for some months
        if(x != 1 or x != 4 or x != 7 or x != 10):

            # Figure out the row number in the excel file that needs to be used for insertions
            # The nth row from the 'Categories' cell
            catBelowIndex = lastRow - catBelow - 1

        else:
            catBelowIndex = lastRow - catBelow

        # Read the category names and possible equations from catBelow to lastRow and store in a list
        currentRow = catBelowIndex
        cutNames = []
        cutAmts = []

        while (currentRow <= lastRow):
            # Prevents 0-indexing the January row
            if(currentRow == 0):
                catBelowIndex = 1
                lastRow += 1

            cutNames.append(info.yearSheetEq[currentRow][revCol].value)
            cutAmts.append(info.yearSheetEq[currentRow][revCol+1].value)
            currentRow += 1

        # Insert the new category
        info.yearSheetEq[catBelowIndex][revCol].value = categoryName
        info.yearSheetEq[catBelowIndex][revCol+1].value = None

        # Write the categories and the amounts to the cell one row down
        # row and lastRow need to be shifted down one
        row = catBelowIndex + 1
        lastRow += 1
        ind = 0

        # Insert the saved data
        while (row <= lastRow):
            info.yearSheetEq[row][revCol].value = cutNames[ind]
            info.yearSheetEq[row][revCol+1].value = cutAmts[ind]

            ind += 1
            row += 1

        # Shift the entire row down only when we reach the left month in the row
        if (x % 3 == 0):
            shiftRowOfMonthsDown(r)

        x += 1

# Helper function that shifts the entire row of months down by the appropriate amount
def shiftRowOfMonthsDown(leftMonth):
    print(leftMonth)
    # Given rightMonth
    ind = info.months.index(leftMonth)

    # Get the title cells
    rightMonthCell = findMonthTitleCell(leftMonth)
    midMonthCell = findMonthTitleCell(info.months[ind+1])
    leftMonthCell = findMonthTitleCell(info.months[ind+2])

    # Assign the correct shift amount
    # As the bottom rows need to shift more than the upper rows
    if (leftMonth == "October"):
        shiftAmount = 3
    elif (leftMonth == "July"):
        shiftAmount = 2
    elif (leftMonth == "April"):
        shiftAmount = 1
    else: # (leftMonth == "Janurary"):
        return

    # Update the info start cells
    # info.cells[ind][0] += shiftAmount
    # info.cells[ind+1][0] += shiftAmount
    # info.cells[ind+2][0] += shiftAmount

    # Get the last row num which should be the same for all three months
    lastRow = info.getRowNum(info.yearSheetEq, leftMonthCell[0], 1)

    rightMonthData = []
    midMonthData = []
    leftMonthData = []
    currentRow = leftMonthCell[0]

    # Save the data into lists and clear the cells
    while(currentRow < lastRow):
        # Right
        rightMonthData.append([info.yearSheetEq.cell(row=currentRow, column=rightMonthCell[1]).value,
                               info.yearSheetEq.cell(row=currentRow, column=rightMonthCell[1]+1).value])

        # Mid
        midMonthData.append([info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1]).value,
                            info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1]+1).value])

        # Left
        leftMonthData.append([info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1]).value,
                              info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1]+1).value])

        currentRow += 1

    # Shift down the appropriate amount
    currentRow = leftMonthCell[0] + shiftAmount
    lastRow += shiftAmount

    ind = 0
    # Re-insert all the data from the three lists
    while (currentRow < lastRow):
        # Values
        # Right
        info.yearSheetEq.cell(row = currentRow, column = rightMonthCell[1]).value = rightMonthData[ind][0]
        info.yearSheetEq.cell(row = currentRow, column = rightMonthCell[1]+1).value = rightMonthData[ind][1]

        # Mid
        info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1]).value = midMonthData[ind][0]
        info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1] + 1).value = midMonthData[ind][1]

        # Left
        info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1]).value = leftMonthData[ind][0]
        info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1] + 1).value = leftMonthData[ind][1]

        ind += 1
        currentRow += 1

    # Clear the cells that were just copied and moved down
    currentRow = leftMonthCell[0]
    lastRow = leftMonthCell[0] + shiftAmount
    while(currentRow < lastRow):
        # Values
        # Right
        info.yearSheetEq.cell(row=currentRow, column=rightMonthCell[1]).value = None
        info.yearSheetEq.cell(row=currentRow, column=rightMonthCell[1] + 1).value = None

        # Mid
        info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1]).value = None
        info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1] + 1).value = None

        # Left
        info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1]).value = None
        info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1] + 1).value = None

        # Fill
        # Right
        info.yearSheetEq.cell(row=currentRow, column=rightMonthCell[1]).fill = info.noFill
        info.yearSheetEq.cell(row=currentRow, column=rightMonthCell[1] + 1).fill = info.noFill

        # Mid
        info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1]).fill = info.noFill
        info.yearSheetEq.cell(row=currentRow, column=midMonthCell[1] + 1).fill = info.noFill

        # Left
        info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1]).fill = info.noFill
        info.yearSheetEq.cell(row=currentRow, column=leftMonthCell[1] + 1).fill = info.noFill

        currentRow += 1

# Helper that gets the actual title cell for each month
# Moved to the left one and up two cells
def findMonthTitleCell(month):
    row, col = newMonth.getMonthStartCell(month)
    row -= 2
    col -= 1
    return row, col

# Window to be shown later
updateWindow = tk.Toplevel(info.window)
updateWindow.withdraw()

# Label for the user text input
tk.Label(updateWindow, text="Enter the new category name").grid(row=0, column=1, padx=5, pady=5)
newNameEntry = tk.Entry(updateWindow, justify='center', font="Calibri 11", width=18)
newNameEntry.grid(row=0, column=2, padx=5, pady=5)
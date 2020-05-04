# Contains information and some helper functions that almost every file needs.
import openpyxl as xl

# Helper Function
def getOpenRow(sheet, startRow, startCol):
    while(sheet[startRow][startCol].value != None):
        startRow += 1

    return startRow

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

monthSheetEq = wbEq["Monthly"]
yearSheetEq = wbEq["Yearly"]
dataSetSheetEq = wbEq["Data Set"]
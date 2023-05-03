import openpyxl
from openpyxl import Workbook
import os

print("Hi! Welcome to the Bill Reader, an SMS Fisher tool by Amani Medcroft.")
print("Please note that this tool requires the catalog to be in a very specific format.")
print("Please check your bill spreadsheet to make sure that the format is correct.")
print("If you encounter any bugs, issues, or have any complaints/suggestions, just let me know!")
print()
file = input("Please enter the filepath for the catalog spreadsheet (.xlsx file) you'd like to input: ")
file = file.replace('"', '')
exchRate = float(input("What is the exchange rate today? (1.00 EUR = X.XX USD): "))

splitFile = file.split('\\')
splitFile = splitFile[-1].split('.')
fileName = splitFile[0] + '_Fishbowl'
folder = os.path.expanduser('~\Downloads')
bill = openpyxl.load_workbook(file, data_only=True)
billSheet = bill.active
newWorkbook = Workbook()
newSheet = newWorkbook.active
maxRow = billSheet.max_row
billRow = 1
newRow = 2

# Initializing spreadsheet by adding column names
# newSheet['A1'] = 'PartNumber'

# Reading address and stuff.
receiverName = billSheet.cell(row=14, column=98).value
receiverAddress = billSheet.cell(row=18, column=98).value
receiverSuite = billSheet.cell(row=19, column=98).value
receiverZipCityState = billSheet.cell(row=20, column=98).value

for num in range(0, maxRow - 1):

    # Checks for description value. If none, skip row.
    if billSheet.cell(row=billRow, column=33).value is None:
        billRow += 1
        continue

    # Collects all necessary information from bill
    # If an item has a code, reads the size and quantity information, saving it in a dict (size:quantity of that size).
    # Also reads color (color) and first two letters of color (shortColor).
    itemCode = billSheet.cell(row=billRow, column=4).value
    if len(itemCode) > 3:
        billSizeRow = billRow + 5
        billQuanRow = billSizeRow + 1

        # If row with sizes isn't in place, searches for the size row below row with item code.
        if billSheet.cell(row=billSizeRow, column=28).value is None:
            # tempBillRow = billRow + 1
            # while len(str(billSheet.cell(row=tempBillRow, column=28).value)) < 1:
            #     tempBillRow += 1
            #     billSizeRow = tempBillRow
            #     billQuanRow = tempBillRow + 1
            billRow += 1
            continue

        sizeQuanDict = {}

        color = str(billSheet.cell(row=billSizeRow, column=6).value)
        shortColor = color[0:2]

        for tempNum in [28, 38, 42, 53, 61, 72, 83, 90, 99, 109]:
            if billSheet.cell(row=billQuanRow, column=tempNum).value is not None:
                size = billSheet.cell(row=billSizeRow, column=tempNum).value
                quantity = billSheet.cell(row=billQuanRow, column=tempNum).value
                sizeQuanDict.update({size:quantity})

    sizes = sizeQuanDict.keys()

    itemFullNames = []
    itemQuantities = []
    for sizeKey in sizes:
        itemFullNames.append(f"{itemCode} - {shortColor} - {sizeKey}")
        itemQuantities.append(sizeQuanDict[sizeKey])

    # Outputs information into new formatted worksheet
    for item in range(0, len(itemFullNames)):
        newSheet.cell(row=newRow, column=1).value = itemFullNames[item]
        newSheet.cell(row=newRow, column=2).value = itemQuantities[item]
        newRow +=1


    billRow += 1

newWorkbook.save(filename=f'{folder}\\{fileName}.xlsx')



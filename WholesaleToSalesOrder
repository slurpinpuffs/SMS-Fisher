import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
import os

version = '1.0'

print(f"Hi! Welcome to the Wholesale to Sales Order tool ver. {version}, an SMS Fisher tool by Amani Medcroft.")
print("Please note that this tool requires a wholesale order made with the Catalog to Wholesale tool.")
print("If you encounter any bugs, issues, or have any complaints/suggestions, just let me know!")
print()
file = input("Please enter the filepath for the catalog spreadsheet (.xlsx file) you'd like to input: ")
file = file.replace('"', '')
print("Great! Now we need some info about the customer.")
customerNameIn = input("Customer's team/company name in Fishbowl (new customer created if none found): ")
billToNameIn = input("Customer name for billing: ")
billToAddressIn = input("Customer address: ")
billToCityIn = input("Customer city: ")
billToStateIn = input("Customer state: ")
billToZipIn = input("Customer ZIP code: ")
billToCountryIn = input("Customer country: ")
carrierNameIn = input("Carrier name: ")
taxRateIn = input("Tax rate name: ")

isNew = 'n'
splitFile = file.split('\\')
splitFile = splitFile[-1].split('.')
fileName = splitFile[0] + '_SO'
folder = os.path.expanduser('~\Downloads')
saveName = f'{folder}\\{fileName}.xlsx'
wholesaleWb = openpyxl.load_workbook(file, data_only=True)
whSheet = wholesaleWb.active
newWorkbook = Workbook()
newSheet = newWorkbook.active
maxRow = whSheet.max_row
whRow = 3
newRow = 4
offset = 0

# Initializing spreadsheet by adding column names
newSheet['A1'] = 'Flag'
newSheet['B1'] = 'SONum'
newSheet['C1'] = 'Status'
newSheet['D1'] = 'CustomerName'
newSheet['E1'] = 'CustomerContact'
newSheet['F1'] = 'BillToName'
newSheet['G1'] = 'BillToAddress'
newSheet['H1'] = 'BillToCity'
newSheet['I1'] = 'BillToState'
newSheet['J1'] = 'BillToZip'
newSheet['K1'] = 'BillToCountry'
newSheet['L1'] = 'ShipToName'
newSheet['M1'] = 'ShipToAddress'
newSheet['N1'] = 'ShipToCity'
newSheet['O1'] = 'ShipToState'
newSheet['P1'] = 'ShipToZip'
newSheet['Q1'] = 'ShipToCountry'
newSheet['R1'] = 'CarrierName'
newSheet['S1'] = 'TaxRateName'
newSheet['T1'] = 'PriorityId'
newSheet['A2'] = 'Flag'
newSheet['B2'] = 'SOItemTypeID'
newSheet['C2'] = 'ProductNumber'
newSheet['D2'] = 'ProductDescription'
newSheet['E2'] = 'ProductQuantity'
newSheet['F2'] = 'UOM'
newSheet['G2'] = 'ProductPrice'
newSheet['H2'] = 'Taxable'
newSheet['I2'] = 'TaxCode'
newSheet['J2'] = 'Note'
newSheet['K2'] = 'ShowItem'
newSheet['L2'] = 'KitItem'
newSheet['M2'] = 'RevisionLevel'
newSheet['N2'] = 'CustomerPartNumber'
newSheet['O2'] = 'CFI-'

# Adding customer info
newSheet['A3'] = 'SO'
newSheet['C3'] = '20'
newSheet['D3'] = customerNameIn
newSheet['E3'] = billToNameIn
newSheet['F3'] = billToNameIn
newSheet['G3'] = billToAddressIn
newSheet['H3'] = billToCityIn
newSheet['I3'] = billToStateIn
newSheet['J3'] = billToZipIn
newSheet['K3'] = billToCountryIn
newSheet['L3'] = billToNameIn
newSheet['M3'] = billToAddressIn
newSheet['N3'] = billToCityIn
newSheet['O3'] = billToStateIn
newSheet['P3'] = billToZipIn
newSheet['Q3'] = billToCountryIn
newSheet['R3'] = carrierNameIn
newSheet['S3'] = taxRateIn

# Checks where Quantity column is on Wholesale sheet to account for offset
if whSheet['B2'].value == 'Quantity':
    offset = 0
else:
    if whSheet['C2'].value == 'Quantity':
        offset = 1

for num in range(0, maxRow - 1):
    quantity = whSheet.cell(row=whRow, column=2+offset).value

    # If item has no quantity, skip it!
    if isinstance(quantity,int):
        if quantity == 0:
            whRow += 1
            continue
    else:
        whRow += 1
        continue
    
    sku = whSheet.cell(row=whRow, column=5+offset).value

    # If item has no SKU, skip it!
    if sku == 'None':
        whRow += 1
        continue
    colorCode = whSheet.cell(row=whRow,column=6+offset).value
    size = whSheet.cell(row=whRow,column=8+offset).value
    price = whSheet.cell(row=whRow, column=9+offset).value

    # Combines size, color code, and SKU into partNum
    if(size != 'None') and (colorCode != 'None') and (colorCode != "--"):
        partNumParts = [str(sku), str(colorCode), str(size)]
        partNum = " - ".join(partNumParts)
    elif(size == 'None') and (colorCode != 'None') and (colorCode != "--"):
        partNumParts = [str(sku), str(colorCode), 'UNI']
        partNum = " - ".join(partNumParts)
    elif(size != 'None') and ((colorCode == 'None') or (colorCode == "--")):
        partNumParts = [str(sku), str(size)]
        partNum = " - ".join(partNumParts)
    else:
        partNumParts = [str(sku), 'UNI']
        partNum = " - ".join(partNumParts)

    # Outputs information into new formatted worksheet
    newSheet.cell(row=newRow, column=1).value = 'Item'
    newSheet.cell(row=newRow, column=2).value = '10'
    newSheet.cell(row=newRow, column=3).value = partNum
    newSheet.cell(row=newRow, column=5).value = quantity
    newSheet.cell(row=newRow, column=6).value = 'ea'
    newSheet.cell(row=newRow, column=7).value = price

    whRow += 1
    newRow += 1

newWorkbook.save(filename=saveName)

print()
print(f'Success! The wholesale order has been converted. The output file is at "{saveName}"')

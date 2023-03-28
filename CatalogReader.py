
import openpyxl
from openpyxl import Workbook
import os

file = input("Please enter the filepath for the catalog spreadsheet you'd like to input: ")
file = file.replace('"', '')
isNew = input("Would you like to only import items marked 'NEW' in the catalog? (y/n): ")
exchRate = float(input("What is the exchange rate today? (1.00 EUR = X.XX USD): "))

splitFile = file.split('/')
splitFile = splitFile[-1].split('.')
splitFile = splitFile[0]
fileName = splitFile + '_Fishbowl'
fileName = 'SantiniSpringSummer2023Collection_Fishbowl'
folder = os.path.expanduser('~/Downloads')
catalog = openpyxl.load_workbook(file, data_only=True)
catSheet = catalog.active
newWorkbook = Workbook()
newSheet = newWorkbook.active
maxRow = catSheet.max_row
catRow = 1
newRow = 2

# Initializing spreadhseet by adding column names
newSheet['A1'] = 'PartNumber'
newSheet['B1'] = 'PartDescription'
newSheet['C1'] = 'PartDetails'
newSheet['D1'] = 'UOM'
newSheet['E1'] = 'UPC'
newSheet['F1'] = 'PartTypeID'
newSheet['G1'] = 'Active'
newSheet['H1'] = 'StdCost'
newSheet['I1'] = 'Tracks-Lot Number'
newSheet['J1'] = 'Tracks-Revision Level'
newSheet['K1'] = 'Tracks-Expiration Date'
newSheet['L1'] = 'Tracks-Serial Number'
newSheet['M1'] = 'AssetAccount'
newSheet['N1'] = 'COGSAccount'
newSheet['O1'] = 'AdjustmentAccount'
newSheet['P1'] = 'ScrapAccount'
newSheet['Q1'] = 'VarianceAccount'
newSheet['R1'] = 'ABCCode'
newSheet['S1'] = 'Weight'
newSheet['T1'] = 'WeightUOM'
newSheet['U1'] = 'Width'
newSheet['V1'] = 'Height'
newSheet['W1'] = 'Len'
newSheet['X1'] = 'SizeUOM'
newSheet['Y1'] = 'ConsumptionRate'
newSheet['Z1'] = 'PartURL'
newSheet['AA1'] = 'PartRevision'
newSheet['AB1'] = 'ProductNumber'
newSheet['AC1'] = 'ProductDescription'
newSheet['AD1'] = 'ProductDetails'
newSheet['AE1'] = 'Price'
newSheet['AF1'] = 'ProductSKU'
newSheet['AG1'] = 'ProductUPC'
newSheet['AH1'] = 'ProductActive'
newSheet['AI1'] = 'ProductTaxable'
newSheet['AJ1'] = 'ProductSOItemTypeID'
newSheet['AK1'] = 'IncomeAccount'
newSheet['AL1'] = 'ProductWeight'
newSheet['AM1'] = 'ProductWeightUOM'
newSheet['AN1'] = 'ProductWidth'
newSheet['AO1'] = 'ProductHeight'
newSheet['AP1'] = 'ProductLen'
newSheet['AQ1'] = 'ProductSizeUOM'
newSheet['AR1'] = 'Vendor'
newSheet['AS1'] = 'DefaultVendor'
newSheet['AT1'] = 'VendorPartNumber'
newSheet['AU1'] = 'Cost'
newSheet['AV1'] = 'VendorUOM'


for num in range(0, maxRow - 1):
    # If row is marked as NEW, then:
    # Collects all necessary information from catalog
    new = catSheet.cell(row=catRow, column=2).value

    if (isNew == 'y') or (isNew == 'Y'):
        if (new != 'NEW'):
            catRow += 1
            continue
        
    sku = catSheet.cell(row=catRow, column=8).value
    uom = catSheet.cell(row=catRow, column=5).value
    if (catSheet.cell(row=catRow, column=9).value != ''):
        color = catSheet.cell(row=catRow, column=9).value
    else:
        color = 'None'
    if (catSheet.cell(row=catRow, column=11).value != ''):
        size = catSheet.cell(row=catRow, column=11).value
    else:
        size = 'UNI'
    if(size != 'None'):
        partNumParts = [str(sku), str(color), str(size)]
        partNum = " - ".join(partNumParts)
    else:
        partNumParts = [str(sku), str(color), 'UNI']
        partNum = " - ".join(partNumParts)
    shortDesc = catSheet.cell(row=catRow, column=6).value
    longDesc = catSheet.cell(row=catRow, column=7).value
    cost = (catSheet.cell(row=catRow, column=27).value) * exchRate
    salesPrice = catSheet.cell(row=catRow, column=28).value

    # Outputs information into new formatted worksheet
    newSheet.cell(row=newRow, column=1).value = partNum
    newSheet.cell(row=newRow, column=2).value = shortDesc
    newSheet.cell(row=newRow, column=3).value = longDesc
    if(uom):
        newSheet.cell(row=newRow, column=4).value = int(uom)
    else:
        newSheet.cell(row=newRow, column=4).value = uom
    newSheet.cell(row=newRow, column=7).value = 'TRUE'
    newSheet.cell(row=newRow, column=8).value = cost
    newSheet.cell(row=newRow, column=28).value = partNum
    newSheet.cell(row=newRow, column=29).value = shortDesc
    newSheet.cell(row=newRow, column=30).value = longDesc
    newSheet.cell(row=newRow, column=31).value = salesPrice
    newSheet.cell(row=newRow, column=34).value = 'TRUE'
    newSheet.cell(row=newRow, column=44).value = 'Santini Maglificio Sportivo'
    newSheet.cell(row=newRow, column=46).value = partNum
    newSheet.cell(row=newRow, column=47).value = cost

    catRow += 1
    newRow += 1

newWorkbook.save(filename=f'{folder}\\{fileName}.xlsx')



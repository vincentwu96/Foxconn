import openpyxl
import os
import xlrd
from datetime import datetime

print('Loading please wait...')

path = r'.'
sheets = os.listdir(path)

new_wb = openpyxl.Workbook()    # New WB
new_sheet = new_wb.active
new_sheet.title = 'New Sheet'
new_sheet['D1'] = 'Part #'      # Set up Header
new_sheet['E1'] = 'ATO #'
new_sheet['F1'] = 'Part Description'
new_sheet['G1'] = 'Part SN'
new_sheet['J1'] = 'Unit SN'
new_sheet['L1'] = 'WO'
new_sheet['M1'] = 'Qty'
new_sheet['N1'] = 'Failure Symptom'
new_sheet['P1'] = 'MRB #'
new_sheet['Q1'] = 'MRB Date'
new_sheet['R1'] = 'Mfg Name'
row = 2

for sheet in sheets:
    if sheet.lower().endswith(".xlsx"):
        wb = openpyxl.load_workbook(os.path.join(path, sheet))
        sheet1 = wb.get_sheet_by_name('Sheet1')

        if(":" in str(sheet1['A20'].value)):
            new_sheet['D' + str(row)] = str(sheet1['A20'].value).split(':', 1)[1].lstrip()     # Part No
        else:
            new_sheet['D' + str(row)] = str(sheet1['A20'].value)
        if(":" in str(sheet1['E20'].value)):
            new_sheet['E' + str(row)] = str(sheet1['E20'].value).split(':', 1)[1].lstrip()     # ATO Number
        else:
           new_sheet['E' + str(row)] = str(sheet1['A20'].value)
        if(":" in str(sheet1['H20'].value)):
            new_sheet['F' + str(row)] = str(sheet1['H20'].value).split(':', 1)[1].lstrip()     # Part Description
        else:
           new_sheet['F' + str(row)] = str(sheet1['H20'].value)
        new_sheet['G' + str(row)] = sheet1['A24'].value                                         # Part SN
        new_sheet['J' + str(row)] = sheet1['C24'].value                                         # System/Unit SN
        new_sheet['L' + str(row)] = sheet1['G24'].value                                         # WO
        new_sheet['M' + str(row)] = sheet1['R24'].value                                         # Qty
        new_sheet['N' + str(row)] = sheet1['K24'].value                                         # Failure Symptom sheet1['L24'].value
        new_sheet['P' + str(row)] = sheet1['L2'].value                                          # MRB No
        new_sheet['Q' + str(row)] = sheet1['L4'].value.strftime("%m/%d/%Y")                     # MRB Date
        new_sheet['R' + str(row)] = sheet1['I24'].value                                         # Mfg Name
        new_wb.save(r'./Output/MRB Report.xlsx')
        if(not sheet1['A25'].value == None or sheet1['C25'].value == None):
            print('Check', sheet1['L2'].value)
        int(row)
        row = row + 1
        
    if sheet.lower().endswith(".xls"):
        wb = xlrd.open_workbook(os.path.join(path, sheet), formatting_info = True)
        sheet2 = wb.sheet_by_index(0)

        if(":" in str(sheet2.cell_value(19,0))):
            new_sheet['D' + str(row)] = str(sheet2.cell_value(19,0)).split(':', 1)[1].lstrip()  # Part No
        else:
            new_sheet['D' + str(row)] = str(sheet2.cell_value(19,0))
            
        if(":" in str(sheet2.cell_value(19,4))):
            new_sheet['E' + str(row)] = str(sheet2.cell_value(19,4)).split(':', 1)[1].lstrip()  # ATO Number
        else:
            new_sheet['E' + str(row)] = str(sheet2.cell_value(19,4))
            
        if(":" in str(sheet2.cell_value(19,7))):
            new_sheet['F' + str(row)] = str(sheet2.cell_value(19,7)).split(':', 1)[1].lstrip()  # Part Description
        else:
            new_sheet['F' + str(row)] = str(sheet2.cell_value(19,7))
        new_sheet['G' + str(row)] = str(sheet2.cell_value(23,0))                                # Part SN
        new_sheet['J' + str(row)] = str(sheet2.cell_value(23,2))                                # System/Unit SN
        if("." in str(sheet2.cell_value(23,6))):
            new_sheet['L' + str(row)] = str(sheet2.cell_value(23,6)).split('.', 1)[0]           # WO
        else:
            new_sheet['L' + str(row)] = str(sheet2.cell_value(23,6))
        if("." in str(sheet2.cell_value(23,17))):
            new_sheet['M' + str(row)] = str(sheet2.cell_value(23,17)).split('.', 1)[0]          # Qty
        else:
            new_sheet['M' + str(row)] = str(sheet2.cell_value(23,17))
        new_sheet['N' + str(row)] = str(sheet2.cell_value(23,10))                               # Failure Symptom    str(sheet2.cell_value(23,10)) 
        new_sheet['P' + str(row)] = str(sheet2.cell_value(1,11))                                # MRB No
        new_sheet['Q' + str(row)] = str(datetime(*xlrd.xldate_as_tuple(sheet2.cell_value(3,11), wb.datemode)).strftime("%m/%d/%Y"))   # MRB Date
        new_sheet['R' + str(row)] = str(sheet2.cell_value(23,8))                                # Mfg Name
        new_wb.save(r'./Output/MRB Report.xlsx')

        if(not str(sheet2.cell_value(24,0)).strip() == '' or str(sheet2.cell_value(23,2)).strip() == ''):
            print('Check', str(sheet2.cell_value(1,11)))
        int(row)
        row = row + 1
        
print('Completed!')
os.system("pause")


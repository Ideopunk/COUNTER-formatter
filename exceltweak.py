from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import string
import re
import sys

# if len(sys.argv) != 3:
#    print("Please include the name of the file being tweaked and a file name where the output should go")

# wb = load_workbook(sys.argv[1])

wb = load_workbook('TR_J1_Input.xlsx')
ws = wb.active

# METADATA SECTION

metadataKeeps = ['Report_Name', 'Report_ID', 'Reporting_Period', 'Created']
metadataGoodbyes = []

# Find the rows we don't want
for count, row in enumerate(ws.rows, 1):
    # Reach the end of the metadata section
    if row[0].value == None:
        break
    if row[0].value not in metadataKeeps:
        metadataGoodbyes.append(count)

# Purge the rows we don't want
for row in reversed(metadataGoodbyes):
    ws.delete_rows(row)

# Reformat the dates in the remaining rows.
rp = ws['B3'].value
rp = re.sub(r'[\w_]+?=', '', rp)
rp = re.sub(';', ' to', rp)
ws['B3'].value = rp

cr = ws['B4'].value
cr = re.sub('T.*', '', cr)
ws['B4'].value = cr

# Style metadata area (do when reinserted, actually)

# for row in ws.iter_rows(min_row = 1, max_col = 2, max_row = 4): 
#     for cell in row:
#         cell.font = Font(name = 'Calibri', size = 12, bold = True)
        # background color...



#DATA SECTION

rowCount = ws.max_row
print(rowCount)

#keep the metadata that's going to be wiped when columns are removed, reinsert later...
keepMetadatas = []
for Bcell in ws.iter_rows(min_col=2, min_row=1, max_col=2, max_row=4, values_only=True):
    Bcell = ''.join(Bcell)
    keepMetadatas.append(Bcell)

dataGoodbyes = ['Publisher', 'Publisher_ID', 'DOI', 'Proprietary_ID', 'URI']
dataGoodbyeColumns = []

# find the table columns we don't want
for count, column in enumerate(ws.iter_cols(min_row=6, max_col=24, max_row=6, values_only=True), 1):
    # Reach the end of the metadata section
    try: 
        column = ''.join(column)
    except:
        break
    if column in dataGoodbyes:
        dataGoodbyeColumns.append(count)

# Purge the columns we don't want
for column in reversed(dataGoodbyeColumns):
    ws.delete_cols(column)

# find the 'total' column and move it to the end. 
origcolumn = ''
movecolumn = ''
breaker = 0
for count, column in enumerate(ws.iter_cols(min_col=1, min_row=6, max_col=24, max_row=6, values_only=True), 1):
    for cell in column:
        print(cell)
        if cell == 'Reporting_Period_Total':
            print('total')
            print(count)
            origcolumn = count
            columnletter = string.ascii_uppercase[origcolumn - 1]
            print(columnletter)
        if cell == None:
            movement = count - origcolumn
            print(movement)
            print(f'{columnletter}:{columnletter}')
            ws.move_range(f'{columnletter}1:{columnletter}{rowCount}', cols = movement)
            # Delete original
            ws.delete_cols(origcolumn)
            breaker = 1
            break
    if breaker == 1:
        break

# reinsert metadata
print(keepMetadatas)

for row in ws.iter_rows(min_col=2, min_row=1, max_col=2, max_row=4):
    for cell in row: 
        cell.value = keepMetadatas.pop(0)
        cell.font = Font(name = 'Calibri', size = 12, bold = True)

# END
#wb.save(sys.argv[2])
wb.save('Output.xlsx')
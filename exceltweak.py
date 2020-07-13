from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import string
import re
import sys

def findMetric(ws):
    # Find the column with metric types
    for row in ws.iter_rows(min_col = 1, min_row = tablerow, max_col= columnCount, max_row = tablerow):
        for cell in row:
            if cell.value == 'Metric_Type':
                metricColumn = cell.column
                return metricColumn


def removeMetrics(check, ws):
    doomlist = []

    # Find the column with metric types
    metricColumn = findMetric(ws)

    # Go through each row to check its metric type
    for row in ws.iter_rows(min_col = metricColumn, min_row=tablerow, max_col = metricColumn, max_row = rowCount):
        for cell in row:
            if cell.value == check:
                doomlist.append(cell.row)
    
    # Destroy the rows we don't want. 
    for row in reversed(doomlist):
        ws.delete_rows(row)

def sheetsplit(ws):
    print('sheetsplit!')
    wb.copy_worksheet(ws)
    if wbtype == 'TR_B1':
        ws.title = 'TR B1 Unique COUNTER 5'
    else:        
        ws.title = 'TR J1 Unique COUNTER 5'
    removeMetrics('Total_Item_Requests', ws)

    # switch to other sheet
    ws = wb['Sheet1 Copy']
    if wbtype == 'TR_B1':
        ws.title = 'TR B1 Total COUNTER 5'
    else:
        ws.title = 'TR J1 Total COUNTER 5'
    removeMetrics('Unique_Item_Requests', ws)


def tablesplit(ws):
    print('tablesplit!')
    metricColumn = findMetric(ws)
    for row in ws.iter_rows(min_col = 1, min_row = tablerow, max_col = columnCount, max_row = rowCount):
        print(row)
        if row[metricColumn].value == 'Unique_Item_Requests':
            continue
        else:
            for cell in row: 
                col = cell.column
                col = string.ascii_uppercase[col]
                newrow = cell.row - tablerow + rowCount
                ws[f'{col}{newrow}'] = cell.value
            if row[metricColumn].value == 'Total_Item_Requests':
                ws.delete_rows(cell.row)



# if len(sys.argv) != 3:
#    print("Please include the name of the file being tweaked and a file name where the output should go")

# wb = load_workbook(sys.argv[1])

wb = load_workbook('TR_B1_Input.xlsx')
ws = wb.active
wbtype = ws['B2'].value

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

#keep the metadata that's going to be wiped when columns are removed, reinsert later...
keepMetadatas = []
for Bcell in ws.iter_rows(min_col=2, min_row=1, max_col=2, max_row=4, values_only=True):
    Bcell = ''.join(Bcell)
    keepMetadatas.append(Bcell)



# Columns to be removed. 
dataGoodbyes = ['Publisher', 'Publisher_ID', 'DOI', 'Proprietary_ID', 'URI']

if wbtype == 'TR_B1':
    dataGoodbyes.append('Print_ISSN')
    dataGoodbyes.append('Online_ISSN')

dataGoodbyeColumns = []


# find header row of table to be used through rest of program as variable

tablerow = ''
for count, row in enumerate(ws.iter_rows(min_col=1, min_row=1, max_col=1, max_row=25, values_only=True), 1):
    for cell in row:
        if cell == 'Title':
            tablerow = count



# find the table columns we don't want
columnCount = ws.max_column
for count, column in enumerate(ws.iter_cols(min_row=tablerow, max_col=columnCount, max_row=tablerow, values_only=True), 1):
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

# reinsert metadata
for row in ws.iter_rows(min_col=2, min_row=1, max_col=2, max_row=4):
    for cell in row: 
        cell.value = keepMetadatas.pop(0)
        cell.font = Font(name = 'Calibri', size = 12, bold = True)

# find the 'total' column and move it to the end. Along the way, change the titles of columns. 
origcolumn = ''
movecolumn = ''
breaker = 0
monthcheck = 0

# + 1 to give space for the reinsertion
columnCount = ws.max_column + 1
for count, column in enumerate(ws.iter_cols(min_col=1, min_row=tablerow, max_col=columnCount, max_row=tablerow, values_only=True), 1):
    for cell in column:
        if cell == 'Reporting_Period_Total':
            # cell = 'YTD Total'
            origcolumn = count
            columnletter = string.ascii_uppercase[origcolumn - 1]
        if cell == None:
            movement = count - origcolumn
            ws.move_range(f'{columnletter}1:{columnletter}{rowCount}', cols = movement)
            
            # Delete the original column 
            ws.delete_cols(origcolumn)
            breaker = 1
            break
    if breaker == 1:
        break


# henceforth, this is how wide this table is! 
columnCount = ws.max_column

# switch headers to what we like
for column in ws.iter_cols(min_col = 1, min_row = tablerow, max_col = columnCount, max_row = tablerow):
    for cell in column:
        cell.value = cell.value.replace('-2020', '')
        cell.value = cell.value.replace('Reporting_Period_Total', 'YTD Total')


# add row of sums! And bold them! 
ws.insert_rows(tablerow + 1)
rowCount = ws.max_row # henceforth, this is how tall the table is!

for count, column in enumerate(ws.iter_cols(min_col = origcolumn, min_row = tablerow + 1, max_col = columnCount, max_row = tablerow), 5):
    columnletter = string.ascii_uppercase[count]
    cellcode = f'{columnletter}{tablerow + 1}'
    ws[cellcode].value = f"=SUM({columnletter}{tablerow + 2}:{columnletter}{rowCount})"
    ws[cellcode].font = Font(name = 'Calibri', size = 12, bold = True)

# add 'title' to first cell in that row
ws[f'A{tablerow + 1}'].value = 'Total'
ws[f'A{tablerow + 1}'].font = Font(name = 'Calibri', size = 12, bold = True)

# separate totals and uniques

if (rowCount - tablerow) > 20:
    sheetsplit(ws)
else: 
    tablesplit(ws)


# END
#wb.save(sys.argv[2])
wb.save('Output.xlsx')
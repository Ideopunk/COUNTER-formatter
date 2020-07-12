from openpyxl import Workbook
from openpyxl import load_workbook
import re
import sys

# if len(sys.argv) != 3:
#    print("Please include the name of the file being tweaked and a file name where the output should go")

wb = load_workbook('TR_J1_Input.xlsx')

ws = wb.active

print(ws.title)
print(ws['B1'].value)

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
    print(row[0].value)

print(metadataGoodbyes)

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


# End
wb.save('Output.xlsx')
from openpyxl import Workbook
from openpyxl import load_workbook
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

for count, row in enumerate(ws.rows, 1):
    # Reach the end of the metadata section
    if row[0].value == None:
        break
    if row[0].value not in metadataKeeps:
        metadataGoodbyes.append(count)
    print(row[0].value)

print(metadataGoodbyes)

for row in reversed(metadataGoodbyes):
    ws.delete_rows(row)




# End
wb.save('Output.xlsx')
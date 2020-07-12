from openpyxl import Workbook
from openpyxl import load_workbook
import sys

# if len(sys.argv) != 3:
#    print("Please include the name of the file being tweaked and a file name where the output should go")

wb = load_workbook('TR_J1_Input.xlsx')

ws = wb.active

print(ws.title)
print(ws['B1'].value)


wb.save('Output.xlsx')
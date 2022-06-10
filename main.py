import xlrd
import argparse
from copy import copy
import random, string
import pandas as pd

# will compare the new generated barcode to the codes existing in the list
def Compare(new_brc, list_brc):
    for itr in list_brc:
        if itr == new_brc:
            return False
    return True

# ONLY SUPPORTS XLS
# file should be in the same project folder
loc = "inventory_6_09_22.xls"

# open the workbook and grab the column values (of C in this case)
# C column is 2
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

# empty list
list_of_barcodes = []

# here we can itr through the entire sheet and grab the Barcode values
# barcodes start on the 7th row, column stays at 2
for i in range(7, sheet.nrows):
    if sheet.cell_value(i, 2) and sheet.cell_value(i,2) != "Barcode":
        #print(sheet.cell_value(i, 2))
        list_of_barcodes.append(sheet.cell_value(i,2))
    
# pass in information from the CLI
parser = argparse.ArgumentParser()
parser.add_argument("Amount", type = int)
parser.add_argument("Length", type = int)
parser.add_argument("--str", dest = "String") # make this optional

args = parser.parse_args()
amnt = args.Amount
length = args.Length
if args.String:
    str = args.String
else:
    str = ''


# generate new barcodes and compare to make sure barcodes do not repeat
asc = string.ascii_uppercase
num = string.digits
complete_possible = asc + num
temp_brc = str
new_barcodes = []

for i in range(amnt):
    for j in range(length - len(str)):
        temp_brc += random.choice(complete_possible)
    
    if Compare(temp_brc, list_of_barcodes):
        new_barcodes.append(copy(temp_brc))
    else:
        i -= 1

    temp_brc = str

#print(new_barcodes)

# output the complete list of barcodes
df = pd.DataFrame(new_barcodes)
df.to_csv('new_barcodes.csv')
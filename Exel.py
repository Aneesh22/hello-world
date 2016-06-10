import openpyxl
import csv
with open('C:/aneesh/MsuPatches.csv', 'rb') as f:
    reader = csv.reader(f)
    for row in reader:
        for x in row[0:2]:
            '''#print x
wb = openpyxl.Workbook('C:/aneesh/MsuPatches.csv')
sheet = wb.get_sheet_by_name('MsuPatches')'''

list=[
    'h','hyg',7,4,'abc','zzz','AS',20.0]
list.sort(reverse=True)
print list
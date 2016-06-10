import openpyxl

month = raw_input('Enter the Secutity Patches Month: EG: January,February:')
month1=month
FileName='C:/Users/aneesh goel/Desktop/PythonProjects/Security Patches/SecurityPatches_'+month+'2016.xlsx'
Filename1=open('C:/Users/aneesh goel/Desktop/PythonProjects/Security Patches/sheetname.txt','w')
Filename1.write(FileName)
wb = openpyxl.Workbook()
wb.save(FileName)
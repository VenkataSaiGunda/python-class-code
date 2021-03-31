import openpyxl as op
wb=op.load_workbook('/content/realestatedata.xlsx')
mysheets=wb.sheetnames
print("print all the sheets in the excel")
for i in mysheets:
  print(i)
sheet=wb['venkata sai']
print("type of the sheet")
print(type(sheet))
print("sheet name")
print(sheet.title)

#output:
print all the sheets in the excel
Sheet2
venkata sai
type of the sheet
<class 'openpyxl.worksheet.worksheet.Worksheet'>
sheet name
venkata sai



import openpyxl
wb = openpyxl.load_workbook('/content/realestatedata.xlsx')
mysheets = wb.sheetnames

print("print all sheets in Excel Workbook")
for x in mysheets:
  print(x)

realestatedata = wb['Sheet2']

print("print type sheet")
print(type(realestatedata))
print("sheet name") 
print(realestatedata.title)

print("Value of A1")
data1 = realestatedata['A1']
print(data1.value)

for i in range(1, 190):
	print(realestatedata.cell(row=i, column=10).value)

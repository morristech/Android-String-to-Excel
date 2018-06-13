from xml.dom import minidom
import xlsxwriter

# parse  xml files by name
mydoc = minidom.parse('/values-tr/strings.xml')
mydocEn = minidom.parse('/values/strings.xml')

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('HPTranslate.xlsx') # increase here every build
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

#get items
items = mydoc.getElementsByTagName('string')
itemsEn = mydocEn.getElementsByTagName('string')


#write all items data 
for elem in items:
    worksheet.write(row, col,     elem.attributes['name'].value)
    if hasattr(elem.firstChild, 'data'):
    	worksheet.write(row, col + 1, elem.firstChild.data)
    row += 1
for elemEn in itemsEn:
	tempList = [i for i,elem in enumerate(items) if elem.attributes['name'].value == elemEn.attributes['name'].value]
	actualRow = 0
	if len(tempList) == 0:
		actualRow = row
		worksheet.write(actualRow, col,     elemEn.attributes['name'].value)
		row += 1
	elif len(tempList) == 1:
		actualRow = tempList[0]
	else :
		break 
	if hasattr(elemEn.firstChild, 'data'):
		worksheet.write(actualRow, col + 2,elemEn.firstChild.data)
		
workbook.close()

from xml.dom import minidom
import xlsxwriter
import os
import fnmatch

def dir_list_folder(head_dir, dir_name):
    outputList = []
    for root, dirs, files in os.walk(head_dir):
    	for filename in fnmatch.filter(files, 'strings.xml'):
        	outputList.append(os.path.join(root, filename))
        # for d in dirs:
        #     if dir_name in d:
        #         outputList.append(os.path.join(root, d))
    return outputList

def askForProjectPath():
	return raw_input("Write the project path (Eg: /Users/mobilion/DemoProject) : \n")

def askForExcelName():
	return raw_input("Write the excell name (Eg: DemoProject) : \n")

# Set first values
row = 0
col = 0
tempItems = []
allNameAttributes = []
paths = dir_list_folder(askForProjectPath() + '/app/src/main/res/', 'values')
excelName = askForExcelName()

if len(paths) != 0:
	# Get all string tags
	for path in paths:
		tempItems.append(minidom.parse(path).getElementsByTagName('string'))
	# Create a workbook and add a worksheet.
	workbook = xlsxwriter.Workbook(excelName + '.xlsx')
	worksheet = workbook.add_worksheet()

	indexOfTempItems = 1
	for items in tempItems:
		for elem in items:
			tempList = [i for i,tempElem in enumerate(allNameAttributes) if tempElem == elem.attributes['name'].value]
		  	actualRow = 0
			if len(tempList) == 0:
				actualRow = row
				worksheet.write(actualRow, col,     elem.attributes['name'].value)
				allNameAttributes.append(elem.attributes['name'].value)
				row += 1
			elif len(tempList) == 1:
				actualRow = tempList[0]
			else :
				break 
			if hasattr(elem.firstChild, 'data'):
				worksheet.write(actualRow, col + indexOfTempItems, elem.firstChild.data)
		indexOfTempItems += 1
	workbook.close()
else :
	print("No path found")

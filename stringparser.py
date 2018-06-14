# from xml.dom import minidom
import xlsxwriter
import os
import fnmatch
import xlrd
import xml.etree.ElementTree as ET  

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
	return raw_input("Write the excel name (Eg: DemoProject) : \n")

def askForDirection():
	answer = raw_input("Write E to xml -> excel\nWrite X to excel -> xml : \n")
	if answer == "E":
		xmlToExcel()
	elif answer == "X":
		excelToXml()
	else :
		print("Run again. Please write only E or X.")

def excelToXml():
	dataExcel = xlrd.open_workbook(askForExcelName() + '.xls')
	table = dataExcel.sheets()[0]
	for j in range(table.ncols):
		if j + 1 < table.ncols:
			# create the file structure
			data = ET.Element('resources')  
			for i in range(table.nrows):
				if i > 0: 
					items = ET.SubElement(data, 'string')
					items.set('name',table.cell_value(i, 0))
					items.text = table.cell_value(i, j + 1)
			# create a new XML file with the results
			mydata = ET.tostring(data)
			folderName = "values-" + table.cell_value(0, j + 1)
			if not os.path.exists(folderName):
				os.makedirs(folderName) 
			myfile = open(folderName + "/strings.xml", "w")  
			myfile.write(mydata)  

def xmlToExcel():
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
			tempItems.append(ET.parse(path).getroot().findall('string'))
		# Create a workbook and add a worksheet.
		workbook = xlsxwriter.Workbook(excelName + '.xls')
		worksheet = workbook.add_worksheet()

		indexOfTempItems = 1
		for items in tempItems:
			for elem in items:
				tempList = [i for i,tempElem in enumerate(allNameAttributes) if tempElem == elem.get('name')]
			  	actualRow = 0
				if len(tempList) == 0:
					actualRow = row
					worksheet.write(actualRow, col,     elem.get('name'))
					allNameAttributes.append(elem.get('name'))
					row += 1
				elif len(tempList) == 1:
					actualRow = tempList[0]
				else :
					break 
				worksheet.write(actualRow, col + indexOfTempItems, elem.text)
			indexOfTempItems += 1
		workbook.close()
	else :
		print("No path found")

askForDirection()

# from xml.dom import minidom
import xlsxwriter
import os
import fnmatch
import xlrd
import xml.etree.ElementTree as ET  
import io

myKeysAndValues = []

def dir_list_folder(head_dir, file_name):
    outputList = []
    for root, dirs, files in os.walk(head_dir):
    	for filename in fnmatch.filter(files, file_name):
        	outputList.append(os.path.join(root, filename))
    return outputList

def askForProjectPath():
	return raw_input("Write the project path (Eg: /Users/mobilion/DemoProject) : \n")

def askForProjectPathiOS():
	return raw_input("Write the project's Assets path (Eg: /Users/mobilion/DemoProjectIos/DemoProject/Assets) : \n")

def askForExcelName():
	return raw_input("Write the excel name (Eg: DemoProject) : \n")

def askForDirection():
	answer = raw_input("Write X to xml -> excel\nWrite E to excel -> xml & strings \nWrite S to strings -> excel : \n")
	if answer == "X":
		xmlToExcel()
	elif answer == "E":
		excelToBoth()
	elif answer == "S":
		stringsToExcel()
	else :
		print("Run again")

def excelToXml(dataExcel):
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
			myfile.close()

def excelToStrings(dataExcel):
	table = dataExcel.sheets()[0]
	for j in range(table.ncols):
		if j + 1 < table.ncols:
			folderName = table.cell_value(0, j + 1) + ".lproj"
			if not os.path.exists(folderName):
				os.makedirs(folderName) 
			myfile = io.open(folderName + "/Localizable.strings", encoding='utf_8', mode='w')   
			for i in range(table.nrows):
				if i > 0:
					mydata = "\"" + table.cell_value(i, 0) + "\" = \"" + table.cell_value(i, j + 1) + "\";\n"
					myfile.write(mydata)
			myfile.close()

def xmlToExcel():
	# Set first values
	row = 0
	col = 0
	tempItems = []
	allNameAttributes = []
	paths = dir_list_folder(askForProjectPath() + '/app/src/main/res/', 'strings.xml')
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

def appendString(line):
	myKeysAndValues.append( [n.replace("\"", "") for n in line.strip().split('=')])

def stringsToExcel():
	# Set first values
	row = 0
	col = 0
	allNameAttributes = []
	paths = dir_list_folder(askForProjectPathiOS(), 'Localizable.strings')
	

	if len(paths) != 0:
		excelName = askForExcelName()
		# Create a workbook and add a worksheet.
		workbook = xlsxwriter.Workbook(excelName + '.xls')
		worksheet = workbook.add_worksheet()

		indexOfTempItems = 1
		# Get all string tags
		for path in paths:
			lines = open(path).read().split('\n')
			for line in lines:
				appendString(line)
				for elem in myKeysAndValues:
					if len(elem) == 2 :
						tempList = [i for i,tempElem in enumerate(allNameAttributes) if tempElem == elem[0]]
					  	actualRow = 0
						if len(tempList) == 0:
							actualRow = row
							worksheet.write(actualRow, col,     elem[0])
							allNameAttributes.append(elem[0])
							row += 1
						elif len(tempList) == 1:
							actualRow = tempList[0]
						else :
							break 
						worksheet.write(actualRow, col + indexOfTempItems, elem[1].decode('utf-8'))
			indexOfTempItems += 1
		workbook.close()
	else :
		print("No path found")
	
def excelToBoth():
	dataExcel = xlrd.open_workbook(askForExcelName() + '.xls')
	excelToXml(dataExcel)
	excelToStrings(dataExcel)

askForDirection()

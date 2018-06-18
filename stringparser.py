# from xml.dom import minidom
import xlsxwriter
import os
import fnmatch
import xlrd
import xml.etree.ElementTree as ET  
import io



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
	row = 1
	keyCol = 0
	tempItems = []
	allNameAttributes = []
	paths = dir_list_folder(askForProjectPath() + '/app/src/main/res/', 'strings.xml')
	excelName = askForExcelName()

	if len(paths) != 0:
		# Create a workbook and add a worksheet.
		workbook = xlsxwriter.Workbook(excelName + '.xls')
		worksheet = workbook.add_worksheet()

		# Get all string tags
		for path in paths:
			tempItems.append(ET.parse(path).getroot().findall('string'))
			isoLangPath = path.split('/')
			isoLang = ''
			if len(isoLangPath) >= 2:
				isoLang = isoLangPath[ len(isoLangPath) - 2].replace("values-", "")
			worksheet.write(0, keyCol + indexOfLangCol, isoLang)

		indexOfLangCol = 1
		for items in tempItems:
			for elem in items:
				tempList = [i for i,tempElem in enumerate(allNameAttributes) if tempElem == elem.get('name')]
			  	actualRow = 1
				if len(tempList) == 0:
					actualRow = row
					worksheet.write(actualRow, keyCol,     elem.get('name'))
					allNameAttributes.append(elem.get('name'))
					row += 1
				elif len(tempList) == 1:
					actualRow = tempList[0] + 1 #since first row is second one
				else :
					break 
				worksheet.write(actualRow, keyCol + indexOfLangCol, elem.text)
			indexOfLangCol += 1
		workbook.close()
	else :
		print("No path found")

def stringsToExcel():
	# Set first values
	row = 1
	keyCol = 0
	allKeys = []
	paths = dir_list_folder(askForProjectPathiOS(), 'Localizable.strings')
	

	if len(paths) != 0:
		excelName = askForExcelName()
		# Create a workbook and add a worksheet.
		workbook = xlsxwriter.Workbook(excelName + '.xls')
		worksheet = workbook.add_worksheet()

		indexOfLangCol = 1
		for path in paths:
			myKeysAndValues = []
			lines = open(path).read().split('\n')
			isoLangPath = path.split('/')
			isoLang = ''
			if len(isoLangPath) >= 2:
				isoLang = isoLangPath[ len(isoLangPath) - 2].replace(".lproj", "")

			worksheet.write(0, keyCol + indexOfLangCol, isoLang)
			for line in lines:
				myKeysAndValues.append( [n.replace("\"", "") for n in line.strip().split('=')])
				for elem in myKeysAndValues:
					if len(elem) == 2 :
						tempList = [i for i,tempKey in enumerate(allKeys) if tempKey == elem[0]]
					  	actualRow = 1
						if len(tempList) == 0:
							actualRow = row
							worksheet.write(actualRow, keyCol,     elem[0])
							allKeys.append(elem[0])
							row += 1
						elif len(tempList) == 1:
							actualRow = tempList[0] + 1 #since first row is second one
						else :
							break 
						worksheet.write(actualRow, keyCol + indexOfLangCol, elem[1].decode('utf-8'))
			indexOfLangCol += 1
		workbook.close()
	else :
		print("No path found")
	
def excelToBoth():
	dataExcel = xlrd.open_workbook(askForExcelName() + '.xls')
	excelToXml(dataExcel)
	excelToStrings(dataExcel)

askForDirection()

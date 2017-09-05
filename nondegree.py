"""
File that is supposed to check a students full class academic record.  If they complete all the required classes, they are
then tagged as a completer for the program.
"""

import openpyxl, os, logging, collections

os.chdir('M:\Rentention\\Nondegree Python Project\\')

wb = openpyxl.load_workbook("Copy.xlsx")

sheet = wb.get_active_sheet()

studentclasses = []

for row in range(3, sheet.max_row + 1):
	id = sheet['B' + str(row)].value
	previousid = sheet['B'+str(row-1)].value
	classnumber = sheet['AK'+str(row)].value
	if(id==previousid):
		studentclasses.append(classnumber)
	else:
		studentclasses = []

print(studentclasses)
	
	



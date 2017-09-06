"""
File that is supposed to check a students full class academic record.  If they complete all the required classes, they are
then tagged as a completer for the program.
"""

import openpyxl, os, logging, collections
logging.basicConfig(level=logging.WARNING, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir('M:\Rentention\\Nondegree Python Project\\')

wb = openpyxl.load_workbook("Copy.xlsx")

sheet = wb.get_active_sheet()
studentclasses = {}

programs = collections.defaultdict(lambda: collections.defaultdict(int))

"""

Maybe a dictionary like {'IP': {"201310":1}}

"""

for row in range(2, sheet.max_row+1):
	
	id = sheet['B' + str(row)].value
	nextid = sheet['B'+str(row+1)].value
	classnumber = str(sheet['AK'+str(row)].value)
	grade = sheet['AS' + str(row)].value
	program = sheet['AB' + str(row)].value
	year = sheet['AE' + str(row)].value
	logging.warning("Row {}: Year: {} Class: {}".format(row,year, classnumber))
	if(id==nextid):
		studentclasses[classnumber] = grade
		logging.debug('Adding {}'.format(classnumber))
	else:
		studentclasses[classnumber] = grade
		#student teaching
		if(("5417" or "4403" in studentclasses.keys()) and 
		((studentclasses.get("5417") or studentclasses.get("4403")) != "F")):
			logging.warning("{}".format(studentclasses))
			logging.warning("{} found. Adding {}".format(program,year))
			programs[program][year] += 1
		studentclasses = {}


	



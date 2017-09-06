"""
File that is supposed to check a students full class academic record.  If they complete all the required classes, they are
then tagged as a completer for the program.
"""

import openpyxl, os, logging, collections,pprint
logging.basicConfig(level=logging.WARNING, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir('M:\Rentention\\Nondegree Python Project\\')

wb = openpyxl.load_workbook("AllData.xlsx")

sheet = wb.get_active_sheet()
studentclasses = {}
yearandprogram = {}

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
	logging.debug("Row {}: Year: {} Class: {}".format(row,year, classnumber))
	if(id==nextid):
		yearandprogram[classnumber] = [program, year]
		studentclasses[classnumber] = grade
		logging.debug('Adding {}\n'.format(classnumber))
	else:
		#The script has found a new student, so the script will check to see if
		#the student is a completer or not 
		studentclasses[classnumber] = grade
		#Saves program and year for each class
		yearandprogram[classnumber] = [program, year]
		logging.warning("YEAR AND PROGRAM {}\n\n".format(yearandprogram))
		#Checks to see if student has passed student teaching
		if(("5417" or "4403" in studentclasses.keys()) and 
		((studentclasses.get("5417") or studentclasses.get("4403")) != "F")):
			logging.debug("{}".format(studentclasses))
			#Returns the list with the correct program and class number
			info = [x for x in [yearandprogram.get(key) for key in ["5417","4403"]] if x is not None]
			logging.warning("{}".format(info))
			try:
				newyear = info[0][1]
				newprogram = info[0][0]
				logging.warning("{} found. Adding {}".format(newprogram,newyear))
				programs[newprogram][newyear] += 1
			#If the student has not taken ST will cause an index out of bound error
			except:
				pass
		studentclasses = {}
		yearandprogram = {}


pprint.pprint(programs)



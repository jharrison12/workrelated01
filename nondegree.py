"""
File that is supposed to check a students full class academic record.  If they complete all the required classes, they are
then tagged as a completer for the program.
"""

import openpyxl, os, logging, collections,pprint
logging.basicConfig(level=logging.WARNING, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir('M:\Rentention\\Nondegree Python Project\\')

wb = openpyxl.load_workbook("Copy.xlsx")

sheet = wb.get_active_sheet()
studentclasses = {}
yearandprogram = {}

programs = collections.defaultdict(lambda: collections.defaultdict(int))

ellmed = ["5013","5033","5043","5053"]
elleds = ["6013","6033","6043","6053"]

"""
TODO:Maybe check if the subect area if correct. 

"""

#Method that removes failing and incomplete classes from student's dict
def checkforf(edclasses):
	logging.debug("EDCLASSES {}".format(edclasses))
	newclasses = {k:v for k,v in edclasses.items() if (edclasses[k] != "F") or (edclasses[k] != "I")}
	logging.debug("NEWCLASSES {}".format(newclasses))
	return newclasses
	
def classcompare(edclasses, programclasses):
	#Frozenset can comare whether one set is a subset 
	set1 = set([element for element in edclasses.keys()])
	set2 = set([element for element in programclasses])
	logging.warning("SET1 {} \n\nSET 2 {} is {}\n\n".format(set1, set2, set2 <= set1))
	return (set2 <= set1)
	
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
		logging.debug("YEAR AND PROGRAM {}\n\n".format(yearandprogram))
		studentclasses = checkforf(studentclasses)
		logging.debug(studentclasses)
		#Checks to see if student has passed student teaching
		if (classcompare(studentclasses, ["4403"])
		 or classcompare(studentclasses, ["5417"])):
			#Returns the list with the correct program and class number by ignoring None
			info = [x for x in [yearandprogram.get(key) for key in ["5417","4403"]] if x is not None]
			logging.debug("{}".format(info))
			try:
				newyear = info[0][1]
				newprogram = info[0][0]
				logging.debug("{} found. Adding {}".format(newprogram,newyear))
				programs[newprogram][newyear] += 1
			#If the student has not taken ST  try block will cause an index out of bound error
			except:
				pass
		elif (classcompare(studentclasses,ellmed)):
			info = [x for x in [yearandprogram.get(key) for key in ellmed] if x is not None]
			finalyear = max([element[1] for element in info])
			logging.warning("ELL info block {} \nand year block {}".format(info,finalyear))
			try:
				logging.debug("{} found. Adding {}".format(newprogram,newyear))
				programs["ELL Endorsement"][finalyear] += 1
			#If the student has not taken ST  try block will cause an index out of bound error
			except:
				pass
		#TODO DO ell eds
		studentclasses = {}
		yearandprogram = {}
		


pprint.pprint(programs)


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

ellmed = ["5013","5033","5043","5053"]
elleds = ["6013","6033","6043","6053"]
icc = ["5033","5273","5293","5283"]
icceds = ["6033","6273","6293","6283"]
reading = ["5743", "5753","5763","5773","5783"]
readingeds = ["6743", "6753","6763","6773","6783"]
admin = ["5233","5333","5253","5483","5551","5562","5583","5663"] #add logic for ICM 5003
admineds = ["6233","6333","6253","6483","6551","6562","6583","6903", "6913"]

"""
TODO:Maybe check if the subect area if correct. 
TODO: Are we sure it checks the last semester of the course in the program courses?
TODO: Fix that its only showing non-degree and not degree students who have takent the courses

"""

#Method that removes failing and incomplete classes from student's dict
def checkforf(edclasses):
	logging.debug("EDCLASSES {}".format(edclasses))
	newclasses = {k:v for k,v in edclasses.items() if (edclasses[k] != "F") or (edclasses[k] != "I") or(edclasses[k] != None)}
	logging.debug("NEWCLASSES {}".format(newclasses))
	return newclasses
	
def classcompare(edclasses, programclasses):
	#This method will compare whether all the classes in a particular program are 
	#found in a subset of all the classes that a student has taken
	#This will return true if this is the case. 
	set1 = set([element for element in edclasses.keys()])
	set2 = set([element for element in programclasses])
	logging.debug("SET1 {} \n\nSET 2 {} is {}\n\n".format(set1, set2, set2 <= set1))
	return (set2 <= set1)
	
def programcheck(programclasses, yearandprogram):
	#This method will find the last semester that they were enrolled in a program if classcompare
	#returns true.  It will then add the program name and the semester to the programs dictionary
	#If a student is in a degree-seeking program and completes the requisite courses
	#they WILL be captured by this method. 
	info = {v:x for x,v in [yearandprogram.get(key) for key in programclasses]if x is not None}
	logging.warning("INFO looks likes {}".format(info))
	finalyear = max([element for element in info.keys()])
	programname = info[finalyear] 
	programs[programname][finalyear] += 1
	
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
	else: #New student found.  Process previous student data
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
			programcheck(ellmed,yearandprogram)
		elif (classcompare(studentclasses, elleds)):
			programcheck(elleds,yearandprogram)
		elif (classcompare(studentclasses,icc)):
			programcheck(icc, yearandprogram)
		elif (classcompare(studentclasses,icceds)):
			programcheck(icceds,yearandprogram)		
		elif (classcompare(studentclasses,reading)):
			programcheck(reading,yearandprogram)		
		elif (classcompare(studentclasses,readingeds)):
			programcheck(readingeds,yearandprogram)		
		elif (classcompare(studentclasses,admin)):
			programcheck(admin,yearandprogram)		
		elif (classcompare(studentclasses,admineds)):
			programcheck(admineds,yearandprogram)
		studentclasses = {}
		yearandprogram = {}
		


pprint.pprint(programs)

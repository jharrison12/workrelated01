"""
File that is supposed to check a students full class academic record.  If they complete all the required classes, they are
then tagged as a completer for the program.
"""

import openpyxl, os, logging, collections,pprint
logging.basicConfig(level=logging.WARNING, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir('M:\Rentention\\Nondegree Python Project\\')

wb = openpyxl.load_workbook("AlLData.xlsx")

sheet = wb.get_active_sheet()
studentclasses = {}
yearandprogram = {}

programs = collections.defaultdict(lambda: collections.defaultdict(int))

ellmed = ["EGEL5013","EGEL5033","EGEL5043","EGEL5053"]
elleds = ["EGEL6013","EGEL6033","EGEL6043","EGEL6053"]
icc = ["EG5033","EG5273","EG5293","EG5283"]
icceds = ["EG6033","EG6273","EG6293","EG6283"]
reading = ["EG5743", "EG5753","EG5763","EG5773","EG5783"]
readingeds = ["EG6743", "EG6753","EG6763","EG6773","EG6783"]
admin = ["EG5233","EG5333","EG5253","EG5483","EG5551","EG5562","EG5583","EG5663"] #ICM 5003 no longer offered
admineds = ["EG6233","EG6333","EG6253","EG6483","EG6551","EG6562","EG6583","EG6903", "EG6913"]
abacert = ["EGSE5053", "EGSE5063", "EGSE5073", "EGSE5083", "EGSE5102", "EGSE5112"]
profabacert = ["EGSE5053", "EGSE5063", "EGSE5073", "EGSE5083", "EGSE5102", "EGSE5112", "EGSE5133", "EGSE5143", "EGSE5122"]
spedendorse = ["EGSE5023", "EGSE 5033", "EGSE5043", "EGSE5053", "EGSE5213", "EGSE5223"]


"""
TODO: Maybe check if the subect area if correct. 
TODO: Are we sure it checks the last semester of the course in the program courses?

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
	#If a student is in a degree-seeking program and completes the requisite non-degree courses
	#they WILL be captured by this method. 
	info = {v:x for x,v in [yearandprogram.get(key) for key in programclasses]if x is not None}
	logging.debug("INFO looks likes {}".format(info))
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
	subjectcode = sheet['AJ' + str(row)].value
	classnumber = subjectcode + classnumber
	logging.warning("Row {}: Year: {} Class: {}".format(row,year, classnumber))
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
		elif (classcompare(studentclasses,abacert)):
			programcheck(abacert,yearandprogram)		
		elif (classcompare(studentclasses,profabacert)):
			programcheck(profabacert,yearandprogram)		
		elif (classcompare(studentclasses,spedendorse)):
			programcheck(spedendorse,yearandprogram)
		studentclasses = {}
		yearandprogram = {}


#The dict comprehension will only pull the programs if they match the program names below. So if a program name ELL Endorsement instead of ELL Endorsement Program.
#it will not be pulled
programs = {k:v for k,v in programs.items() if k in ["Admin License", "ELL Endorsement Program", "Instructional Coaching Certificate", 
													"Reading Endorsement", "Teacher Licensure Program", "Gen Std:Appl Behavr Analy,Cert", 
													"Instructional Coaching, Certif"]}

pprint.pprint(programs)

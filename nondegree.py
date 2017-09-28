"""
File that is supposed to check a students full class academic record.  If they complete all the required classes, they are
then tagged as a completer for the program.
"""

import openpyxl, os, logging, collections,pprint, re
logging.basicConfig(level=logging.WARNING, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir('M:\Rentention\\Nondegree Python Project\\')

wb = openpyxl.load_workbook("AllData.xlsx")

sheet = wb.get_active_sheet()
studentclasses = {}
yearandprogram = {}

programs = collections.defaultdict(lambda: collections.defaultdict(int))


#Must use regex as some students were allowed to switch from 5000 level to 6000 level mid-program
ellmed = [re.compile("EGEL[5|6]013"),re.compile("EGEL[5|6]033"),re.compile("EGEL[5|6]043"),re.compile("EGEL[5|6]053")]
icc = [re.compile("EG[5|6]033"),re.compile("EG[5|6]273"),re.compile("EG[5|6]293"),re.compile("EG[5|6]283")]
reading = [re.compile("EG[5|6]743"),re.compile("EG[5|6]753"),re.compile("EG[5|6]763"),re.compile("EG[5|6]773"),re.compile("EG[5|6]783")]
admin = [re.compile("EG[5|6]233"),re.compile("EG[5|6]333"),re.compile("EG[5|6]253"),
		 re.compile("EG[5|6]483"),re.compile("EG[5|6]551"),re.compile("EG[5|6]562"),
		 re.compile("EG[5|6]583"),re.compile("EG[5|6]663")] #ICM 5003 no longer offered
abacert = ["EGSE5053", "EGSE5063", "EGSE5073", "EGSE5083", "EGSE5102", "EGSE5112"]
profabacert = ["EGSE5053", "EGSE5063", "EGSE5073", "EGSE5083", "EGSE5102", "EGSE5112", "EGSE5133", "EGSE5143", "EGSE5122"]
spedendorse = ["EGSE5023", "EGSE 5033", "EGSE5043", "EGSE5053", "EGSE5213", "EGSE5223"]


"""
TODO: Check to make sure if regex works. 
TODO: Are we sure it checks the last semester of the course in the program courses?

You can't check whether a student switched programs because students have taken courses that apply for certificate progrmas under degree
seeking programs. 

"""

#Method that removes failing and incomplete classes from student's dict
def checkforf(edclasses):
	logging.debug("\nEDCLASSES {} \n Length: {}\n".format(edclasses.items(), len(edclasses)))
	logging.debug(edclasses.items())
	# The logic  in the dict comprehension does not work for if 
	#((v != 'F') or (v!= "I") or(v != None) but does work for the 
	# The line below the commented line
	#newclasses = {k:v for k,v in edclasses.items()} if ((v != 'F') or (v!= "I") or(v != None))}
	newclasses = {k:v for (k,v) in edclasses.items() if v not in ["F", "I", None]}
	logging.warning("\nNEWCLASSES {}\n Lenght: {}\n".format(newclasses,len(newclasses)))
	return newclasses
	
def classcompare(edclasses, programclasses):
	#This method will compare whether all the classes in a particular program are 
	#found in a subset of all the classes that a student has taken
	#This will return true if this is the case. 
	try:
		for edclass in programclasses:
			logging.debug("{} in programclasses".format(edclass))
			if any(edclass.match(regex) for regex in edclasses):
				logging.debug("ONE FOUND")
				continue
			else:
				return False	
		return True
	except:
		set1 = set([element for element in edclasses.keys()])
		set2 = set([element for element in programclasses])
		logging.debug("SET1 {} \n\nSET 2 {} is {}\n\n".format(set1, set2, set2 <= set1))
		return set2 <= set1
			
		
	
def programcheck(programclasses, yearandprogram):
	#This method will find the last semester that they were enrolled in a program if classcompare
	#returns true.  It will then add the program name and the semester to the programs dictionary
	#If a student is in a degree-seeking program and completes the requisite non-degree courses
	#they WILL be captured by this method.
	info = {}
	logging.debug("\nYear and program:{} \n\n Programclasses{}\n".format(yearandprogram, programclasses))
	try:
		for key in programclasses:
			doesthiswork = [key.match(regex).group(0) for regex in yearandprogram.keys() if key.match(regex)]
			logging.debug("Does this work: {}".format(doesthiswork))
			logging.debug("Yearandprogram items {}".format([regex for regex in yearandprogram.items()]))
			info.update({x:v for v,x in [yearandprogram.get(item) for item in doesthiswork]})
			logging.debug("Adding {} to info".format(info))
	except:
		info = {x:v for v,x in[yearandprogram.get(key) for key in programclasses]}

	logging.debug("INFO looks likes {}".format(info))
	finalyear = max([element for element in info.keys()])
	programname = info[finalyear] 
	logging.debug("Addind {} {}".format(programname, finalyear))
	programs[programname][finalyear] += 1

	
for row in range(2, sheet.max_row+1):
	id = sheet['B' + str(row)].value
	nextid = sheet['B'+str(row+1)].value
	classnumber = str(sheet['AK'+str(row)].value)
	grade = sheet['AS' + str(row)].value
	program = sheet['AB' + str(row)].value
	nextprogram = sheet['AB' + str(row+1)].value
	year = sheet['AE' + str(row)].value
	subjectcode = sheet['AJ' + str(row)].value
	classnumber = str(subjectcode) + classnumber
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
		logging.debug("\n \n {}".format(studentclasses))
		#Checks to see if student has passed student teaching
		if (classcompare(studentclasses, ["EG4403"])
		 or classcompare(studentclasses, ["EG5417"])):
			#Returns the list with the correct program and class number by ignoring None
			info = [x for x in [yearandprogram.get(key) for key in ["EG5417","ED4403"]] if x is not None]
			logging.debug("{}".format(info))
			try:
				newyear = info[0][1]
				newprogram = info[0][0]
				logging.debug("{} found. Adding {}".format(newprogram,newyear))
				programs[newprogram][newyear] += 1
			#If the student has not taken ST  try block will cause an index out of bound error
			except:
				pass
		elif (classcompare(studentclasses, ["EG513V"])):
			info = [x for x in [yearandprogram.get(key) for key in ["EG513V"]] if x is not None]
			newprogram = info[0][0]
			programs[newprogram]["TOOK TLM: TO CHECK"] += 1
		elif (classcompare(studentclasses,ellmed)):
			programcheck(ellmed,yearandprogram)
		elif (classcompare(studentclasses,icc)):
			programcheck(icc, yearandprogram)	
		elif (classcompare(studentclasses,reading)):
			programcheck(reading,yearandprogram)			
		elif (classcompare(studentclasses,admin)):
			programcheck(admin,yearandprogram)		
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

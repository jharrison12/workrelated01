"""
File that checks a student's id and semester versus the previous row id and semester.
This only works if the id column is sorted and the semester column is sorted based upon the student id
Student id must be in row F, Semester Code must be in row B and the semester codes are static in the dict below
(not ideal).

Also you must have two or three previous semesters included before the semester you want to start measuring.
Why?  Because you must capture the returning students but you cannot do that if you don't have the semester they
started.

"""

import openpyxl, os, logging, collections

#logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir('M:\Rentention\\')

wb = openpyxl.load_workbook("EnrollmentfileforBCR.xlsx")
logging.debug("Has opened excel")

sheet = wb.get_active_sheet()

semesterdict = {"201730": "201720",'201720': '201710','201710': '201630', "201630": "201620",
                "201620": "201610", "201610": "201530",
                "201530": "201520", "201520": "201510",
                "201510": "201430", "201430": "201420",
                "201420": "201410", "201410": "201330",
                "201330": "201320", "201320": "201310",
                "201310": "201230", "201230": "201220",
                "201220": "201210"}

# This will leave row 1 empty
for row in range(3, sheet.max_row + 1):
    id = sheet['F' + str(row)].value
    # logging.DEBUG('Found id {}'.format(id))
    # logging.DEBUG("Working on row {}".format(row))
    # If previous row id is the same and semester is previous semester
    # then the student is continuing
    if (sheet['F' + str(row - 1)].value == id) and (
                sheet['B' + str(row - 1)].value == semesterdict[sheet['B' + str(row)].value]) and (
        sheet['Z' + str(row - 1)].value == sheet['Z' + str(row)].value):
        sheet['G' + str(row)].value = 'C'
    # Otherwise they are returning
    elif (sheet['F' + str(row - 1)].value == id) and (
        sheet['Z' + str(row - 1)].value == sheet['Z' + str(row)].value):
        sheet['G' + str(row)].value = 'R'
    #If Program is different, but ID is the same
	#and semester is not consecutive. They are returning new
    elif (sheet['F' + str(row-1)].value==id) and (
          sheet['B' + str(row - 1)].value != semesterdict[sheet['B' + str(row)].value]) and (
          sheet['Z' + str(row - 1)].value != sheet['Z' + str(row)].value):
          sheet['G'+str(row)].value="RN"
	#Otherwise if they id is the same and the program is not the came then they
	#must have switched programs since semesters will be consecutive
    elif (sheet['F' + str(row-1)].value==id) and (
        sheet['Z' + str(row - 1)].value != sheet['Z' + str(row)].value):
        sheet['G'+str(row)].value= "S"
    # if the previous row is a different id student is a beginner
    else:
        sheet['G' + str(row)].value = "B"

wb.save("testfile1.xlsx")

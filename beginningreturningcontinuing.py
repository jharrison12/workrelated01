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

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir("C:\\Users\\harrisonjd\\Documents")

wb = openpyxl.load_workbook("Enrollment Fall2014-PResent(1).xlsx")
logging.debug("Has opened excel")

sheet = wb.get_active_sheet()

semesterdict = {'201710': '201630', "201630": "201620",
                "201620": "201610", "201610": "201530",
                "201530": "201520", "201520": "201510",
                "201510": "201430", "201430": "201420",
                "201420": "201410", "201410": "201330"}

#This will leave row 1 empty
for row in range(3, sheet.max_row + 1):
    id = sheet['F' + str(row)].value
    logging.INFO("Found id {}".format(id))
    logging.INFO("Working on row {}".format(row))
    #If previou row id is the same and semester is previous semester
    #then the student is continuing
    if (sheet['F' + str(row - 1)].value == id) and (
        sheet['B' + str(row - 1)].value == semesterdict[sheet['B' + str(row)].value]):
        sheet['G' + str(row)].value = 'C'
    #Otherwise they are returning
    elif (sheet['F' + str(row-1)].value == id):
        sheet['G' + str(row)].value = 'R'
    #if the previous row is a different id student is a beginner
    else:
        sheet['G' + str(row)].value = "B"

wb.save("testfile.xlsx")

"""
File that checks a how many students matriculated from ICC into other programs
This only works if the id column is sorted and the semester column is sorted based upon the student id


Also you must have two or three previous semesters included before the semester you want to start measuring.
Why?  Because you must capture the returning students but you cannot do that if you don't have the semester they
started.

"""

import openpyxl, os, logging, collections

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir("M:\Working Folder")

wb = openpyxl.load_workbook("ICMATRICULATION.xlsx")
logging.debug("Has opened excel")

sheet = wb.get_active_sheet()

semesterdict = {'201710': '201630', "201630": "201620",
                "201620": "201610", "201610": "201530",
                "201530": "201520", "201520": "201510",
                "201510": "201430", "201430": "201420",
                "201420": "201410", "201410": "201330",
                "201330": "201320", "201320": "201310",
                "201310": "201230", "201230": "201220",
                "201220": "201210"}

iccdict = ["Instructional Coaching, Certif", "Instructional Coaching, EdS","Instructional Coaching, MED"]
matriculator = 0

# This will leave row 1 empty
for row in range(3, sheet.max_row + 1):
    id = sheet['F' + str(row)].value
    program = sheet['Z' + str(row)].value
    # logging.DEBUG('Found id {}'.format(id))
    # logging.DEBUG("Working on row {}".format(row))
    # if student is instructional coaching
    if program in iccdict:
        if (sheet['F' + str(row + 1)].value == id) and (sheet['Z'+ str(row + 1)].value not in iccdict):
            matriculator += 1
            sheet['AA'+ str(row)].value = 'Y'
    else:
        pass
    # Otherwise they are returning

print(matriculator)


wb.save("testfile.xlsx")


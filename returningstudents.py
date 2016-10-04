### Jonathan Harrison
# file that opens an enrollment file and tries to determine
# how many students are returning

import openpyxl, os, logging, collections

logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir("M:\Working Folder")

wb = openpyxl.load_workbook("Enrollment Fall2014-PResent.xlsx")
logging.debug("Has opened excel")

sheet = wb.get_active_sheet()

semesterdict = collections.OrderedDict()
graduatedorstoppedout = {}

for row in range(2, sheet.max_row + 1):
    semester = sheet['B' + str(row)].value
    program = sheet['N' + str(row)].value
    lnumber = sheet['F' + str(row)].value
   # logging.info("{} {} {}".format(semester, program, lnumber))

    semesterdict.setdefault(semester, {})
    graduatedorstoppedout.setdefault(semester, {})
    graduatedorstoppedout[semester].setdefault(program)
    semesterdict[semester].setdefault(program, set())
    semesterdict[semester][program].add(lnumber)

logging.info(semesterdict.values())
logging.info(graduatedorstoppedout)

#for i, v in semesterdict.items():
  #  print(i,v)

def comparedict(d1,d2, items):
    keymatch = {}
    for i in d1.keys():
        if i[-2:] == '10':
            keymatch[i] = i[:-2] +'20'
        elif i[-2:] == "20":
           keymatch[i]  = i[:-2] + '30'
        else:
           keymatch[i]  = '20{}10'.format(str(int(i[2:4]) + 1))

    for i in d1.values():
        items.update()






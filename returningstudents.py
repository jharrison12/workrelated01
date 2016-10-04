### Jonathan Harrison
# file that opens an enrollment file and tries to determine
#how many students are returning

import openpyxl, os, logging
logging.basicConfig(level=logging.INFO, format=' %(asctime)s - %(levelname)s- %(message)s')

os.chdir("M:\Working Folder")

wb = openpyxl.load_workbook("Enrollment Fall2014-PResent.xlsx")
logging.debug("Has opened excel")
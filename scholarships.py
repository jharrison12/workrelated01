#! python3
# Pulling scholarship codes from banner
#

import pyautogui, openpyxl, pyperclip, os, time


os.chdir('M:\Assessment (New)')

iterations = int(input("how many iterations?"))

nwb = openpyxl.Workbook()
sheet1 = nwb.get_active_sheet()
time.sleep(5)
#pyautogui.hotkey('alt', 'tab')

for i in range(1, iterations):
	pyautogui.hotkey('ctrl', 'c')
	sheet1.cell(row=i, column=2).value = pyperclip.paste()
	pyautogui.press('down')
	#time.sleep(1)

	
nwb.save('studentteachingspring2011.xlsx')
	
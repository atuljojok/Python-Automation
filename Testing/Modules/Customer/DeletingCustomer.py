
import pyautogui
from win32api import GetSystemMetrics
import sys, os
pyautogui.FAILSAFE=False


path="E:\Study\Testing\Modules\Results"
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'Common'))

import ExcelUtil


# print("click on the botton left to launch the start menu")
# pyautogui.click(600,1050)
# print("Select the SalesMate + Application")
# pyautogui.typewrite("SalesMate +")
# print("Press Enter Key to Launch SalesMate + Application and Wait for 10 Sec")
# pyautogui.press("enter",1,10)
# #Sometimes SalesMate + might take 10 seconds to load it . So 10 sec delay
# print("Press Enter Key again to close the Tip Of the Day Dialog")
# pyautogui.press("enter")
pyautogui.click(20,190)
print('click customer deleting  option in Sales mate +')
pyautogui.press(['alt','c'])
pyautogui.press('d')
print('Entering customer Id')


pyautogui.typewrite("9")
pyautogui.click(1150,800)

pyautogui.click(1050,850)

pyautogui.click(1000,590)
print('customer id deleted')
pyautogui.click(1000,580)
pyautogui.press('enter',1)


print("Log the Results to the CSV File")
if os.path.exists(path+"/SalesMatePlus.mdb"):
    ExcelUtil.SaveTestResultToCSV("3","2","Pass",path+"/SalesMatePlus.mdb")
else:
    ExcelUtil.SaveTestResultToCSV("3","2","Fail","") 

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

pyautogui.click(1310,800)
print('open the stock adding option')
pyautogui.press(['alt','s'])
pyautogui.press('enter',1,2)

pyautogui.press('enter',1)
pyautogui.typewrite('2')

pyautogui.press('enter',1)
pyautogui.typewrite('Apple')

pyautogui.press('enter',1)
pyautogui.typewrite('Fruits')

pyautogui.press('enter',1)
pyautogui.typewrite('50kg')

pyautogui.press('enter',1)
pyautogui.typewrite('king fruits PVT  ')

pyautogui.press('enter',1)
pyautogui.typewrite('40')

pyautogui.press('enter',1)
pyautogui.typewrite('35')

pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.typewrite('Nothing')
pyautogui.press('enter',1)
pyautogui.typewrite('fruit section level-2')

pyautogui.click(599,580)
pyautogui.press('enter',1)
pyautogui.moveTo(750,620)
pyautogui.click(750,620)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.typewrite('50')
pyautogui.press('enter',1)
pyautogui.typewrite('60')
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)
pyautogui.press('enter',1)




print("Log the Results to the CSV File")
if os.path.exists(path+"/SalesMatePlus.mdb"):
    ExcelUtil.SaveTestResultToCSV("3","2","Pass",path+"/SalesMatePlus.mdb")
else:
    ExcelUtil.SaveTestResultToCSV("3","2","Fail","")




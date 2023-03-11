
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

print('click customer Adding option in Sales mate +')
pyautogui.press(['alt','c'])
pyautogui.press('enter')
print('Entering customer details')

pyautogui.click(900,305)
pyautogui.typewrite("atul ")
pyautogui.click(905,350)
pyautogui.typewrite("jojo ")
pyautogui.click(905,380)
pyautogui.typewrite("Mr")
pyautogui.click(905,450)
pyautogui.typewrite("kavalckal,champakulam,Alapppuzha")
pyautogui.click(905,480)
pyautogui.typewrite("9100054067")

pyautogui.click(905,510)
pyautogui.typewrite("atul@mail.com")

pyautogui.click(905,575)

pyautogui.click(1000,590)

pyautogui.click(905,585)
pyautogui.typewrite('good service')

pyautogui.click(745,685)
pyautogui.click(920,685)
pyautogui.typewrite("10")

pyautogui.click(1050,850)
print('saving customer details')
pyautogui.click(1050,580)



print("Log the Results to the CSV File")
if os.path.exists(path+"/SalesMatePlus.mdb"):
    ExcelUtil.SaveTestResultToCSV("3","2","Pass",path+"/SalesMatePlus.mdb")
else:
    ExcelUtil.SaveTestResultToCSV("3","2","Fail","")
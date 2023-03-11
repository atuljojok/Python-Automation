
import pyautogui
from win32api import GetSystemMetrics
import sys, os
pyautogui.FAILSAFE=False


sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'Common'))

import ExcelUtil

print("Executing Backup Data Test Case")

print("Create a Database Backup directory called SMP")

path="E:/SMP"
if not os.path.exists(path):
    os.makedirs(path)

#get the width and height of the monitor
width= GetSystemMetrics(0)
height=GetSystemMetrics(1)
#click on the botton left to launch the start menu
print("click on the botton left to launch the start menu")
pyautogui.click(580,1050)
print("Select the SalesMate + Application")
pyautogui.typewrite("SalesMate +")

print("Press Enter Key to Launch SalesMate + Application and Wait for 10 Sec")
pyautogui.press("enter",1,10)
#Sometimes SalesMate + might take 10 seconds to load it . So 10 sec delay
print("Press Enter Key again to close the Tip Of the Day Dialog")
pyautogui.press("enter")
print("Now alt and f shortcut key in  Salesmate to invoke the File Root menu")
pyautogui.press(['alt','f'])

print("Now press 'b' to invoke the Backup Folder Dialog")
pyautogui.press("b")

print('go to e drive')
pyautogui.typewrite("new volume")
pyautogui.press('enter',1)

print('enter arrow key to expand tree')
pyautogui.press('right',1)

print('smp data folder')
pyautogui.typewrite('smp')
pyautogui.press('enter',5)

print('press enter to backup data')
pyautogui.press('enter',5) 

print("Log the Results to the CSV File")
if os.path.exists(path+"/SalesMatePlus.mdb"):
    ExcelUtil.SaveTestResultToCSV("1","2","Pass",path+"/SalesMatePlus.mdb")
else:
    ExcelUtil.SaveTestResultToCSV("1","2","Fail","")


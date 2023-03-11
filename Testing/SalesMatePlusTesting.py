import pyautogui
from subprocess import call
import sys, os
import os.path

sys.path.append(os.path.join(os.path.dirname(__file__), 'Modules\Common'))

import ExcelUtil

print("Running SalesMate + Ver 5.0 Test Cases ..")

my_path = os.path.abspath(os.path.dirname(__file__))

call(["python", os.path.join(my_path, "Modules\FileMenu.py")])

call(["python", os.path.join(my_path, "Modules\ViewMenu.py")])

call(["python", os.path.join(my_path, "Modules\SetupMenu.py")])


call(["python", os.path.join(my_path, "Modules\CustomerMenu.py")])

call(["python", os.path.join(my_path, "Modules\TransactionMenu.py")])

call(["python", os.path.join(my_path, "Modules\SalesMenu.py")])

call(["python", os.path.join(my_path, "Modules\AccountsMenu.py")])

call(["python", os.path.join(my_path, "Modules\AccountsMenu.py")])

call(["python", os.path.join(my_path, "Modules\ReportsMenu.py")])

call(["python", os.path.join(my_path, "Modules\POSMenu.py")])

call(["python", os.path.join(my_path, "Modules\HelpMenu.py")])

print("Dumping all the Results form CSV file to Original test case Document... ")

ExcelUtil.SaveTestResultToExcel("salesmatep_plus_test.csv")
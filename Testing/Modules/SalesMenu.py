from subprocess import call
import os.path


#execute python code

my_path = os.path.abspath(os.path.dirname(__file__))

call(["python", os.path.join(my_path, "Stock\AddingStock.py")])
print("Running Sales Menu Test Cases...")
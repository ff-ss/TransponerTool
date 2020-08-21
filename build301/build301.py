#TRASPONER-TOOL
# FRANCISCO SALCEDO 2020
#BUILD 3.0.1

# TEST FOR CORRECT OPERATING SYSTEM
from transponer_error import testSystem

if not testSystem():
    exit()

# INITIALIZE IMPORTS, ONCE OS IS VALID
import transponer_functions

from colorama import init
from colorama import Fore, Back, Style

#RUN PROGRAM

init() # ANSI COLOR INITIALIZE
transponer_functions.outputcolor() # SET CONSOLE ANSI COLOR

transponer_functions.mainMenu() # SHOW MAIN MENU
transponer_functions.openFile() # SHOW FILE SELECT PROMPT
transponer_functions.openSheet() # SHOW SHEET SELECT PROMPT
transponer_functions.columnWarning() # SHOW COLUMN WARNING AND DISCLAIMER
transponer_functions.assignColumns() # SHOW ASSIGNMENT PROMPT
transponer_functions.setExport() # SHOW EXPORT PROMPT
transponer_functions.compute() # COMPUTE TRANSPOSE WITH SELECTED VARIABLES

while True: # CONTINUE WITH SAME FILE?
    if transponer_functions.newquery(): # SHOW CONTINUE PROMPT
        transponer_functions.openSheet()
        transponer_functions.assignColumns()
        transponer_functions.compute()
    else:
        break

transponer_functions.endProgram() # SHOW RESULT, AND SAVE FILE

import BGG_GameID_Collector
import PAX_Title_Corrector
import BGG_Metadata_Collector

from pathlib import Path #used for printing current working directory (cwd)
import sys #used to quit script
import csv
import openpyxl
from difflib import get_close_matches
from time import sleep

def print_menu():
    print('\n') 
    print (30 * '-')
    print ("   M A I N - M E N U")
    print (30 * '-')
    print ("1. BGG ID# Index Update")
    print ("2. Game Title Correction & BGG ID# Association")
    print ("3. Game Metadata Collection")
    print ("4. Print current working directory")
    print ("5. Exit")
    print (30 * '-')

loop = True

while loop:  ## Continues executing until loop is false
    print_menu()
    choice = int(input('Please enter your choice [1-5]: '))

    ### Take actions as per selected menu-option ###
    if choice == 1:
        print ('\n' + 'Attempting to locate existing BGG ID# index...')
        BGG_GameID_Collector.BGGextract()

    elif choice == 2:
        print ('\n' + 'Starting game title correction... [this may take some time, ~20 seconds]')
        PAX_Title_Corrector.main()

    elif choice == 3:
        print('\n' + 'Starting collection of metadata...')
        BGG_Metadata_Collector.data_collect()

    elif choice == 4: 
        print('\n' + 'Current working directory is: ' + str(Path.cwd()))
        sleep(1)

    elif choice == 5:
        print ('\n' + 'Exiting scipt. Have a nice day!')
        sleep(1)
        sys.exit()

    else:    
        print ("Invalid number. Please try again.")


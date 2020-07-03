import BGG_GameID_Collector
import PAX_Title_Corrector
import BGG_Metadata_Collector

from pathlib import Path #used for printing current working directory (cwd)
import sys #used to quit script
import csv
import openpyxl
from difflib import get_close_matches

def print_menu():
    print (30 * '-')
    print ("   M A I N - M E N U")
    print (30 * '-')
    print ("1. BGG ID# Index Update")
    print ("2. Game Title Matching")
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
            print ("Starting BGG ID# collection...")
            #Determine range of BGG ID #s to extract. If integers are not provided for inputs, set range from 1 to highest number available on BGG website
    
            #Set filename, and if necessary, set first item as 1
            filename = input("Enter the filename of past BGG titles export, or leave blank to create new. Please use / as backslash if typing in directory path: ")
            if filename == '':
                filename = 'BGG_ID_spreadsheet.xlsx'
                first_item = 1
            else:
                #TO DO - Check to ensure that file exists
                first_item = 'max_row'

            #Set upper bound of ID# range
            try:
                last_item = int(input('Enter highest BGG ID value to capture. Blank input will set script to max ID # available: ')) 
            except:
                last_item = BGG_GameID_Collector.BGGmaxitem() #call function to extract highest possible ID# from GeekFeed's RSS

            
            #Call the main function for extracting data from BGG, passing all of the above variables
            BGG_GameID_Collector.BGGextract(filename, first_item, last_item)


    elif choice == 2:
            print ("Starting matching of game titles. Initialization will take some time...")
            PAX_Title_Corrector.main()

    elif choice == 3:
            print ("Starting collection of metadata...")
            BGG_Metadata_Collector.data_collect()


    elif choice == 4: 
            print('Current working directory is: ' + str(Path.cwd()))


    elif choice == 5:
            print ("Exiting scipt. Have a nice day!")
            sys.exit()


    else:    ## default ##
            print ("Invalid number. Please try again.")


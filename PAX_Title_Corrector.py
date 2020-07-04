import csv
import openpyxl
from difflib import get_close_matches
import sys 
from time import sleep 
from datetime import datetime

#All code contained within main() function as to prevent automatic execution when this file is imported into main menu script

def main():
    ########################################################################
    # Function to write a row to the PAX Corrections csv
    ########################################################################

    def TitleWriting(game, newname = '.', status = 'pass'):
        
        # Test for correction conditions . 
        # 1) Newname == . means no value was passed, so write row as-is and add the BGG ID#
        # 2) A blank newname means correction was attempted but failed, so write row with a 0 in BGG ID# field. Log error.
        # 3) Status flag set to fail means game was manual renamed. Need to write newname, but not attempt to find BGG ID#. Log error. 
        # 4) If not blank or ., and status flag not changed, write the newname variable as game title and look up its BGG ID#
        if newname == '.':
            TitleWriter.writerow([game, int(PAXids[PAXnames.index(game)]), BGGnames[game]])
        elif newname == '':
            TitleWriter.writerow([game, int(PAXids[PAXnames.index(game)]), 0])
            ErrorWriter.write(game + '\n')
        elif status == 'fail':
            TitleWriter.writerow([newname, int(PAXids[PAXnames.index(game)]), 0])
            ErrorWriter.write(newname + '\n')
        else:
            TitleWriter.writerow([newname, int(PAXids[PAXnames.index(game)]), BGGnames[newname]])
        return()

    ########################################################################
    # Function to display list of matches and take user input
    ########################################################################

    def MatchSelector(match, title):
        
        print(title + ' - Has potential correction from BGG search:')
        for index, name in enumerate(match):
            print(index + 1, name) #index is incremented for user-friendliness, so list does not start at 0
        print(len(match) + 1, 'No match') #Add an extra entry, one above match list's range, for "N/A" option
        #TO DO - Work around instances where there are multiple identical options. Pull in BGG metadata?

        ### Take user input to select match            
        selected = 0 
        while True:
            #Ensure an integer is input
            try: 
                selected = int(input('Please select a match: '))
            except ValueError:
                print('Input must be numeric')
                continue

            #Then ensure that integer is within range (match index plus 2: 1 so that it doesn't start at 0, plus another 1 to account for "N/A" option)
            if (int(selected) not in range (1,len(match) + 2)):  
                print ('Sorry, input is out of range')
                continue
            else:
                break
        
        return selected

    ###################################################################################################################################
    # Function to drive attempting of a match. Calls MatchSelector and TitleWriting functions. Can sense if first or subsequent attempt.
    ###################################################################################################################################

    def AttemptMatch(game, BGGnames, manual_name = ''):
        if manual_name == '':
            match = get_close_matches(game, BGGnames.keys(), n=3, cutoff = 0.7)
        else:
            match = get_close_matches(manual_name, BGGnames.keys(), n=3, cutoff = 0.7)

        if len(match) != 0:
            selected = MatchSelector(match,game)

            ### Check if a match was selected, or if user chose N/A "No match" option, which will not be a valid index of match list    
            if int(selected) <= int(len(match)):
                print('Thank you for selecting input #' + str(selected) +": " + match[int(selected) - 1])
                TitleWriting(game, match[int(selected) - 1])
                return 'success'
            else:
                return 'fail'  #no match selected
        else:
            return 'fail' #no matches found

    ########################################################################
    # Prepare file I/O
    ########################################################################

    ###### LOAD BGG NAMES ######
    # Open BGG workbook, set active sheet. Initialize a dictionary of BGG names, then iterate reading game name as key and ID# as value
    
    try:
        wb = openpyxl.load_workbook('BGG_ID_spreadsheet_complete.xlsx')
    except: 
        print('Error: Could not locate index of BGG names/IDs. Please load BGG_ID_spreadsheet_complete.xlsx into current working directory and restart script')
        sys.exit()
    sheet = wb.active
    BGGnames = {}
    for x in range(1,sheet.max_row + 1):
        BGGnames[str(sheet.cell(x,2).value)] = sheet.cell(x,1).value


    ###### LOAD INPUT FILE TO COMPARE (Titles Report) ######

    # Initialize lists
    PAXnames = []
    PAXids = []

    # Present user with menu selection of frequently-used .csv files, or allow to input own filename
    print('\n') 
    print (30 * '-')
    print ("   D A T A - I N P U T")
    print (30 * '-')
    print ("1. Tabletop Library's Titles Report (TTLibrary_Titles.csv)")
    print ("2. Prior Export From Corrections Script (PAXcorrections.csv)")
    print ("3. Prior Export (Manual Filename Entry)")
    print (30 * '-')
    
    choice = int(input('Please enter your choice [1-3]: '))

    if choice == 1:
        try:
            PAXgames = open('TTLibrary_Titles.csv', mode='r', newline='')
            reader = csv.reader(PAXgames) # Read .csv into a variable
            for rows in reader:  # Iterate through input csv and append different elements of each row to the appropriate list
                PAXnames.append(rows[0])
                PAXids.append(rows[4])
            header = next(reader)
            header.append('BGG ID')
        except:
            print('Error: Could not locate Titles report. Please load TTLibrary_Titles.csv into current working directory and restart script')
            sleep(3)
            sys.exit()
    elif choice == 2:
        try:
            PAXgames = open('PAXcorrections.csv', 'r', newline='', encoding='utf-16')
            reader = csv.reader(PAXgames) # Read .csv into a variable
            for rows in reader:  # Iterate through input csv and append different elements of each row to the appropriate list
                PAXnames.append(rows[0])
                PAXids.append(rows[1])
            header = next(reader)
        except:
            print('Error: Could not locate past script output. Please load PAXcorrections.csv into current working directory and restart script')
            sleep(3)
            sys.exit()
    elif choice == 3:
        PAX_Titles_path = input("Enter the name of past script output .csv file: ")
        PAXgames = open(PAX_Titles_path, 'r', newline='')
        reader = csv.reader(PAXgames) # Read .csv into a variable
        for rows in reader:  # Iterate through input csv and append different elements of each row to the appropriate list
            PAXnames.append(rows[0])
            PAXids.append(rows[1])
            header = next(reader)
 


    ###### OPEN FILES FOR WRITING ######
    #Open output csv for writing corrected titles, set the writer object, and write the header
    PAXcorrections = open('PAXcorrections.csv', 'w', newline='', encoding='utf-16')
    TitleWriter = csv.writer(PAXcorrections, delimiter=',', escapechar='\\', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
    TitleWriter.writerow(header)

    #Open a text file for mismatch error logging. Write header message with timestamp.
    ErrorWriter = open('Title_Mismatch_Log.txt', 'a+') #append mode used to continue growing log
    ErrorWriter.write('\n \nTitle Corrector was run on ' + str(datetime.now()) + '\nThe following games could not be linked to BGG ID#s: \n')

    ###### COMPARE NAMES & WRITE ROWS ######
    for game in PAXnames:
        if game not in BGGnames.keys():
            print('Attempting match of ' + game)
            status = AttemptMatch(game, BGGnames)
            
            # If match was not found, try again with manually-input revision to title.
            if status == 'fail':
                print('No match was found on first attempt')
                manual_title = input('Please manually enter a corrected title. If unsure, leave blank to skip: ')
                if manual_title == '':
                    TitleWriting(game, manual_title)
                    status = 'done'  #revert the status code returned by AttemptMatch() so user is not prompted for manual input again in next code block (if status == 'fail')
                else:
                    status = AttemptMatch(game, BGGnames, manual_title)
                    
            if status == 'fail':
            # If match was still not found on second attempt, allow for full manual correction without BGG link
                print(game + ' - Has no BGG match found, and was unable to be corrected through manual adjustment')
                manual_title = input('Please manually enter a corrected title. This will be written to the PAX dB without attempt to find a BGG match. If unsure, leave blank to skip: ')
                TitleWriting(game, manual_title, 'fail')
        
        ### Clear match was found between PAX & BGG
        else:
            print(game +  ' - No correction needed. BGG ID# is ' + str(BGGnames[game]) + ' PAX ID# is ' + str(PAXids[PAXnames.index(game)]))
            TitleWriting(game)

        print('\n') 

    PAXcorrections.close()
    ErrorWriter.close()

if __name__ == "__main__":
    main()
from pathlib import Path #used in handling Path objects
import csv
import openpyxl
from difflib import get_close_matches

print('Current working directory is: ' + str(Path.cwd()))

########################################################################
# Function to write a row to the PAX Corrections csv
########################################################################

def TitleWriting(game, newname = ''):
    global PAXnames
    global PAXvaluable
    global PAXcount
    global PAXids
    if newname == '':
        TitleWriter.writerow([game, PAXpublisher[PAXnames.index(game)], bool(PAXvaluable[PAXnames.index(game)]), int(PAXcount[PAXnames.index(game)]), int(PAXids[PAXnames.index(game)])])
    else:
        TitleWriter.writerow([newname, PAXpublisher[PAXnames.index(game)], bool(PAXvaluable[PAXnames.index(game)]), int(PAXcount[PAXnames.index(game)]), int(PAXids[PAXnames.index(game)])])
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
# Function to drive attempting of a match. Calls MatchSelector and TitleWriter functions. Can sense if first or subsequent attempt.
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


###### LOAD BGG NAMES ######
# Open BGG workbook, set active sheet. Initialize a dictionary of BGG names, then iterate reading game name as key and ID# as value
wb = openpyxl.load_workbook('BGG_ID_spreadsheet_complete.xlsx')
sheet = wb.active
BGGnames = {}
for x in range(1,sheet.max_row + 1):
    BGGnames[str(sheet.cell(x,2).value)] = sheet.cell(x,1).value


###### LOAD PAX TITLES ######
# Read in elements of .csv to different lists, iterating over every row in the PAX Titles csv
PAXgames = open('TTLibrary_Titles_withID.csv', mode='r', newline='')
#PAXgames = open('PAXcorrections - June27.csv', mode='r', newline='')
reader = csv.reader(PAXgames)

#Initialize lists
PAXnames = []
PAXpublisher = []
PAXvaluable = []
PAXcount = []
PAXids = []

#Iterate through PAX Titles csv and append different elements of each row to the appropriate list
header = next(reader)
print(header)
for rows in reader:
    PAXnames.append(rows[0])
    PAXpublisher.append(rows[1])
    PAXvaluable.append(rows[2])
    PAXcount.append(rows[3])
    PAXids.append(rows[4])


###### OPEN NEW CSV FOR WRITING ######
#Open file for writing, set the writer object, and write the header
PAXcorrections = open('PAXcorrections.csv', 'w', newline='', encoding='utf-16')
TitleWriter = csv.writer(PAXcorrections, delimiter=',', escapechar='\\', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
TitleWriter.writerow(header)


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
                TitleWriting(game)
            else:
                status = AttemptMatch(game, BGGnames, manual_title)
                
        if status == 'fail':
        # If match was still not found on second attempt, allow for full manual correction without BGG link
            print(game + ' - Has no BGG match found, and was unable to be corrected through manual adjustment')
            manual_title = input('Please manually enter a corrected title. This will be written to the PAX dB without attempt to find a BGG match. If unsure, leave blank to skip: ')
            TitleWriting(game, manual_title)
    
    ### Clear match was found between PAX & BGG
    else:
        print(game +  ' - No correction needed. BGG ID# is ' + str(BGGnames[game]) + ' PAX ID# is ' + str(PAXids[PAXnames.index(game)]))
        TitleWriting(game)

    print('\n') 

PAXcorrections.close()

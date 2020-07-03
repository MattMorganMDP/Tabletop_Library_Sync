import csv
import requests   #support for pulling contents of webpages
from bs4 import BeautifulSoup  #function for reading the page's XML returned by requests library
from time import sleep  #sleep function allows pausing of script, to avoid getting rate-limited by BGG
from random import randint  #generate random integers, used in randomizing wait time
import itertools #uses zip function to iterate over two lists concurrently

###### LOAD PAX TITLES ######
# Read in elements of .csv to different lists, iterating over every row in the PAX Titles csv
PAX_Titles_path = input("Enter the filename of PAX Titles report - please use / as backslash if typing in directory path: ")
if PAX_Titles_path == '':
    PAXgames = open('PAXcorrections.csv', 'r', newline='', encoding='utf-16')
    #PAXgames = open('PAXcorrections - June27.csv', 'r', newline='', encoding='utf-16')
    #PAXgames = open('TTLibrary_Titles-testshort.csv', mode='r', newline='')
    #PAXgames = open('TTLibrary_Titles_withID.csv', mode='r', newline='')
    #PAXgames = open('PAXcorrections_Excel.csv', mode='r', newline='')
else:
    PAXgames = open(PAX_Titles_path, 'r', newline='')
reader = csv.reader(PAXgames)

#Initialize lists
PAXnames = []
PAXids  = []
BGGids = []
ID_range = []

#use next() function to clear the first row in CSV reader, but replace header value with new list of column names for export
header = next(reader)
header = ['Title', 'PAX ID', 'BGG ID', 'Min Player', 'Max Player', 'Year Published', 'Playtime', 'Minimum Age', 'Avg Rating', 'Weight']
print(header)
for rows in reader:
    PAXnames.append(rows[0])
    PAXids.append(rows[4])
    BGGids.append(rows[5])


###### OPEN NEW CSV FOR WRITING ######
#Open file for writing, set the writer object, and write the header
BGGmetadata = open('BGGmetadata.csv', 'w', newline='', encoding='utf-16')
DataWriter = csv.writer(BGGmetadata, delimiter=',', escapechar='\\', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
DataWriter.writerow(header)


###### PARSE METADATA FROM BGG API ######

base_url = 'https://www.boardgamegeek.com/xmlapi2/thing?id='

#Set ID_range to be 100-item chunks from master set of BGG ID#s
for count, IDs in enumerate(BGGids):
    ID_range.append(IDs)
    if (count+1)%100 == 0:
        URL_args = ','.join(list(map(str,ID_range))) 
        url = base_url + URL_args  + '&stats=1'
        print(url)
        
        #Use requests and BeautifulSoup to extract and read XML. Separaetly pull XML tags: (1) of <name> with type "primary", (2) of <item>
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'lxml')
        soup_min = soup.find_all('minplayers') 
        soup_max = soup.find_all('maxplayers')
        soup_year = soup.find_all('yearpublished')
        soup_time = soup.find_all('playingtime')
        soup_age = soup.find_all('minage')
        soup_rating = soup.find_all('average')
        soup_weight = soup.find_all('averageweight')
       
        #Regex processing of soup objects. Include BGG_id sequence from ID_range to use as index value of PAXnames/PAXids when writing csv
        for min_player, max_player, year, time, age, rating, weight, BGG_id in zip(soup_min, soup_max, soup_year, soup_time, soup_age, soup_rating, soup_weight, ID_range):
            game_min_player = min_player.attrs['value']
            game_max_player = max_player.attrs['value']
            year_published = year.attrs['value']
            play_time = time.attrs['value']
            min_age = age.attrs['value']
            avg_rating = rating.attrs['value']
            avg_weight = weight.attrs['value']

            #Write row to CSV only if game has a BGG ID#. Behavior dependent on PAX_Title_Corrector.py behavior that writes zeros to blank BGG ID# fields
            if BGG_id != 0:
                DataWriter.writerow([PAXnames[BGGids.index(BGG_id)], PAXids[BGGids.index(BGG_id)], BGG_id, game_min_player, game_max_player, year_published, play_time, min_age, avg_rating, avg_weight])
                print(PAXnames[BGGids.index(BGG_id)])

        sleep(randint(15,30))  #sleep to prevent rate-limit

        # Clear out ID_range to accept a fresh set of 100 IDs on next loop iteration
        ID_range = []

BGGmetadata.close()
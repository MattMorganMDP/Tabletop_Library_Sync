from pathlib import Path #used in handling Path objects
import csv
import requests   #support for pulling contents of webpages
from bs4 import BeautifulSoup  #function for reading the page's XML returned by requests library
import re    #support for regular expressions
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
header = next(reader)
#TO DO - Set header to be actual fields we want to collect
header = ['Title', 'PAX ID', 'BGG ID', 'Min Player', 'Max Player']
print(header)
header.append('BGG ID')
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

#Set ID_range to be 100-item lists of IDs
for count, IDs in enumerate(BGGids):
    ID_range.append(IDs)
    if (count+1)%100 == 0:
        URL_args = ','.join(list(map(str,ID_range))) 
        url = base_url + URL_args  
        print(url)
        
        #Use requests and BeautifulSoup to extract and read XML. Separaetly pull XML tags: (1) of <name> with type "primary", (2) of <item>
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'lxml')
        soup_min = soup.find_all('minplayers') 
        soup_max = soup.find_all('maxplayers') 
        #TO DO - Continue adding soup objects for desired metadata fiels

        #Regex processing of soup objects. Include BGG_id sequence from ID_range to use as index value of PAXnames/PAXids when writing csv
        for min_player, max_player, BGG_id in zip(soup_min, soup_max, ID_range):
            game_min_player = min_player.attrs['value']
            game_max_player = max_player.attrs['value']

            #Write row to CSV only if game has a BGG ID#
            if BGG_id != 0:
                DataWriter.writerow([PAXnames[BGGids.index(BGG_id)], PAXids[BGGids.index(BGG_id)], BGG_id, game_min_player, game_max_player ])
                #print(min_player, max_player, BGG_id)
                print(PAXnames[BGGids.index(BGG_id)])

        sleep(randint(15,30))  #sleep to prevent rate-limit or DOS

        # Clear out ID_range to accept a fresh set of 100 IDs
        ID_range = []

BGGmetadata.close()
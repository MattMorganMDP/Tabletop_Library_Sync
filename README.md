# Tabletop_Library_Sync

This project is intended to sycronize data between the PAX Tabletop Library and BoardGameGeek.com
Goals are to:
1. Increase accuracy of information contained in the PAX TTL db (game titles, publishers)
2. Increase variety of information contained in the PAX TTL db (game metadata)

This is accomplished by linking game ID numbers between PAX and BGG

Step 1 - Script to pull (or update existing copy of) BGG game title & ID# pairs

Step 2 - Script to consume PAX TTL's Titles report and generate a .csv file that can be used to correct PAX TTL db

Step 3 - Script to generate PAX & BDD ID# table, along with metadata for each title gathered from BGG's XMLAPI2

Additional efforts:
- Modification of PAX TTL's titles.rb model to tweak the Titles report, adding PAX ID# as a field, and eliminating initcaps previously applied to all titles

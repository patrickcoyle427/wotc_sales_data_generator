#!/usr/bin/env python3

'''
wotc_sales_data_generator.py

Takes a CSV from lightspeed retail POS, pulls out the sales information that Wizards
Of the Coast needs, then writes them to an xlsx file that can be sent for reporting.

-Written by Patrick Coyle

USAGE:

- On first run of this script, two directories will be created in the same directory as this
  file: to_parse and parsed_data

- Place Lightspeed Retail POS sales data into the to_parse folder, these files are .csv files
  that are created in sales by line in the reporting section of Lightspeed Retail POS

- WOTC wants 1 month of all Magic: The Gathering and Dungeons and Dragons Sales. To create
  the data that will be used for this script, In Lightspeed POS, go to Reports, then Sales
  Lines. Set the date range to be the month you want sales from. For Magic: The Gathering
  search the "MTG" tag, and exclude "accessory". For Dungeons and Dragons, search the "dnd"
  tag and exclude "mini".

'''

#TO DO
# Load in template xlsx file
# Check test xlsx load.py to see how to implement it into the main code
# Write data to a copy of the template
# Get UPC data from somehwere (file? Ask user?
#   If ask user, its very easy to save their answers for that session and
#   the future! This can save a dictionary of all answers given.

import csv, os, os.path

# csv - reads the data from Lightspeed Retail POS
# os - used for creating folders and checking folder contents
# os.path - Used for checking if the 'to_parse' and 'parsed_files'
#           folders exist

# will need to add from openpyxl import load_workbook

def start_parse():

    # Runs script through each step of the parsing process

    dir_exist = dir_check()

    if dir_exist == True:

        findings = find_csv_names()

        pulled_data = pull_data(findings)

        count = 0

        for i in pulled_data:

            print(i)
            count+=1

        print(count)
        

def dir_check():

    # Will create necessary directories on first launch of script,
    # then will pass each other time. Gives a message to let the user know
    # what to do with the created directories

    if os.path.isdir('to_parse') and os.path.isdir('parsed_files'):

        return True

    else:

        create_these = ('to_parse', 'parsed_files')

        for i in create_these:

            if not os.path.isdir(i):

                os.makedirs(i)

        print('Folders for parsing created. Please place all files that need to be parsed',
              'into the "to_parse" folder, then run this script again.')

        return False

def find_csv_names():

    # Finds the names of the files to be parsed.

    return [file for file in os.listdir('to_parse') if file.endswith('.csv')]

def pull_data(names):

    # Takes the data from each CSV and loads it into memory to merged and written
    # to a single .xlsx file.

    wotc_id = '5658'
    # Alternate Universes's wotc ID
    # WOTC ID number is the first column in the spreadsheet that will be written

    data = []

    for file in names:

        with open(f'to_parse/{file}', newline='') as csvfile:

            reader = csv.reader(csvfile, delimiter=',')

            next(reader, None)
            # Skips the header line

            for row in reader:

                data.append((wotc_id, row[1], row[0], row[2], row[3], row[4], row[5]))


    return data

        

if __name__ == '__main__':

    start_parse()

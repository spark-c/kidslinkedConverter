# A rewrite of the original script for better organization and clarity. 02/24/2021

#This program scrapes a formatted document and converts the found text into an organized spreadsheet.
#Contact info must be in blocks with at least one empty line between companies.
#First line of a block must be company name.

#Collin Sparks 11/9/2019
#Last updated 12/4/2019
#python 3.8.0

# Format will be Company(0), Name(1), Email(2), Phone(3), Address(4)

import re, pprint, logging, openpyxl, os, sys
from decouple import config
from Company import * # defines the "Company" object
from kidslinked_regex import * # regex definitions

SOURCE = config('SOURCE') # sets whether the contact info source is local (clipboard.txt) or remote (AJAX)
if SOURCE not in ['local', 'remote']:
    SOURCE = 'local'
    print('Environment SOURCE not set correctly! Defaulting to local. (local, remote)')


def infoScrape(sourceDoc):
    all_companies = []
    sourceDoc = sourceDoc.split('\n\n') # splits on empty lines
    print('sourcedoc: ' + str(sourceDoc))
    for block in sourceDoc: # each block is a company's info
        pprint.pprint('block: ' + str(block))
        rawData = block.splitlines() # assumes that each line is a different piece of data to be regex'd
        pprint.pprint('rawdata: ' + str(rawData))

        #COMPANY NAME
        if len(rawData) < 2: # Deletes empty lists
            del rawData
            continue

        while rawData[0] == '': # Deletes beginning empty lines
            del rawData[0]

        if dbaRegex.search(rawData[0]):         # 'dba' = 'doing business as', which needs removed.
            dbaMatch = dbaRegex.search(rawData[0])
            dbaResult = dbaMatch.group(2)
            rawData[0] = dbaResult.strip()

        all_companies.append(Company(rawData[0])) # assumes that the first line is the company name.
        newComp = all_companies[-1]

        del rawData[0]

        # NAMES
        foundNames = []
        for item in rawData:
             if nameRegex.search(item) != None:
                 foundNames.append(item.strip())

        for person in foundNames:
            newComp.add('contact', person)

        # EMAILS
        foundEmails = []
        for item in rawData:
             if emailRegex.search(item) != None:
                 emailMatch = emailRegex.search(item)
                 emailResult = emailMatch.group(1)
                 foundEmails.append(emailResult.strip())

        for email in foundEmails:
            newComp.add('email', email)

        # PHONES
        foundPhones = []
        for item in rawData:
             if phoneRegex.search(item) != None:
                 phoneMatch = phoneRegex.search(item)
                 phoneResult = phoneMatch.group()
                 foundPhones.append(phoneResult.strip())

        for phone in foundPhones:
            newComp.add('phone', phone)

        # ADDRESSES
        foundAddress = []
        foundAddress1 = ''
        foundAddress2 = ''
        for item in rawData:
             if address1Regex.search(item) != None:
                 address1Match = address1Regex.search(item)
                 foundAddress1 = address1Match.group()

        for item in rawData:
             if address2Regex.search(item) != None:
                 address2Match = address2Regex.search(item)
                 foundAddress2 = address2Match.group()

        foundAddress = foundAddress1 + ' ' + foundAddress2
        newComp.add('address', foundAddress.strip())

        if len(newComp.emails) == 0:
            all_companies.remove(newComp) # if there's no email on file, throw out the object

    return all_companies


def get_source(origin, data): # accepts data; 'remote' = (json string) or 'local' = filepath for .txt with string content
    if origin == 'local':
        # Break the whole document into a giant list of blocks
        with open(data, 'r') as f:
            sourceDoc = f.read()
        return sourceDoc

    elif origin == 'remote':
        if isinstance(data, str): # this should be a JSON string
            return data
        else:
            return ''

def get_destination(SOURCE):
    if SOURCE == 'local':
        while True:
            try:
                filepath = input('Please enter desired path: ')
                return filepath

                os.chdir(filepath)
            except:
                print('Invalid path!')
                continue
    elif SOURCE == 'remote':
        return


def generate_wb(all_companies): # creates an excel workbook from the initialized companies and returns the wb
    wb = openpyxl.Workbook()
    sheet = wb['Sheet']

    firstRow = {'A1': 'Company', 'B1' : 'Names', 'C1' : 'Emails', 'D1' : 'Phones', 'E1' : 'Address'}
    for item in firstRow: # writes column labels in first row
        sheet[item] = firstRow[item]

    columnkeys = { # this will be used to map obj attrs to the correct column letter
    'name': 'A',
    'contacts': 'B',
    'emails': 'C',
    'phones': 'D',
    'address': 'E',
    }

    row_index = 2 # begins at the row under the column titles
    for obj in all_companies:
        rows_needed = max( # determines how many rows are needed to display all of this company's info
            len(obj.contacts),
            len(obj.phones),
            len(obj.emails)
            )


        for i in range(rows_needed + 1): # this loop should run as many times as we need rows
            for key in obj.__dict__:
                try: # handles out-of-index errors for shorter lists
                    sheet[columnkeys[key] + str(row_index + i)] = obj.__dict__[key][i] # e.g. sheet['A3'] = obj[key]
                except:
                    continue
        row_index += rows_needed # sets up new starting point for next Company.

    return wb


def export_wb(SOURCE, wb):
    if SOURCE == 'local':
        sheet_destination = get_destination(SOURCE)
        while True:
            try:
                filename = input('Enter filename: ')
                if filename[-5:] != '.xlsx':
                    filename = filename + '.xlsx'
                wb.save(filename)
                break
            except:
                print('Invalid filename!')
                continue
    elif SOURCE == 'remote':
        export_wb('local', wb)


def convert_to_wb(origin, data):
    sourceDoc = get_source(origin, data)
    all_companies = infoScrape(sourceDoc)
    finished_wb = generate_wb(all_companies)
    export_wb(SOURCE, finished_wb)


if __name__ == '__main__':
    convert_to_wb(SOURCE, input("Enter the path of the source file: "))

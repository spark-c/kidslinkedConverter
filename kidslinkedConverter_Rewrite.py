# A rewrite of the original script for better organization and clarity. 02/24/2021

#This program scrapes a formatted document and converts the found text into an organized spreadsheet.
#Contact info must be in blocks with at least one empty line between companies.
#First line of a block must be company name.

#Collin Sparks 11/9/2019
#Last updated 12/4/2019
#python 3.8.0

# Format will be Company(0), Name(1), Email(2), Phone(3), Address(4)

import re, pprint, logging, openpyxl, os, sys
from Company import * # defines the "Company" object
from kidslinked_regex import * # regex definitions


def infoScrape():
    i = 1
    for block in range(len(sourceDoc)):

        Company = 'Company %i' % i
        print(Company)
        collectedInfo = {Company: {   # list for temp storage of found data
                            'name':None}}

        working = sourceDoc[block]

        rawData = working.splitlines() # assumes that each line is a different piece of data to be regex'd

        #COMPANY
        if len(rawData) < 2: # Deletes empty lists
            del rawData
            continue

        while rawData[0] == '': # Deletes beginning empty lines
            del rawData[0]

        if dbaRegex.search(rawData[0]):         # 'dba' = 'doing business as', which needs removed.
            dbaMatch = dbaRegex.search(rawData[0])
            dbaResult = dbaMatch.group(2)
            rawData[0] = dbaResult.strip()

        collectedInfo[Company]['name'] = rawData[0] # assumes that the first line is the company name
        del rawData[0]

        # NAMES
        foundNames = []
        for item in rawData:
             if nameRegex.search(item) != None:
                 foundNames.append(item.strip())

        j = 1
        collectedInfo[Company]['contacts'] = {}
        for person in foundNames:
            collectedInfo[Company]['contacts']['person %i' % j] = person
            j += 1

        # EMAILS
        foundEmails = []
        for item in rawData:
             if emailRegex.search(item) != None:
                 emailMatch = emailRegex.search(item)
                 emailResult = emailMatch.group(1)
                 foundEmails.append(emailResult.strip())

        j = 1
        collectedInfo[Company]['emails'] = {}
        for email in foundEmails:
            collectedInfo[Company]['emails']['email %i' % j] = email
            j += 1

        # PHONES
        foundPhones = []
        for item in rawData:
             if phoneRegex.search(item) != None:
                 phoneMatch = phoneRegex.search(item)
                 phoneResult = phoneMatch.group()
                 foundPhones.append(phoneResult.strip())

        j = 1
        collectedInfo[Company]['phones'] = {}
        for number in foundPhones:
            collectedInfo[Company]['phones']['number %i' % j] = number
            j += 1

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

        collectedInfo[Company]['address'] = foundAddress.strip()

        if len(foundEmails) != 0:
            bigDict.update(collectedInfo) # adds found info to the directory

            i += 1


if not os.path.isfile('./clipboard.txt'):
    input('Place text content into ./clipboard.txt\n***Press ENTER to exit***')
    open('./clipboard.txt', 'a').close()
    sys.exit()



bigDict = {} # this is the main directory



# Break the whole document into a giant list of blocks

#sourceDoc = pyperclip.paste() ****************THIS DOES NOT WORK ON LINUX?????***************
with open(r'./clipboard.txt', 'r') as f:
    sourceDoc = f.read()

#sourceDoc = sourceDoc.split('\r#\n\r\n')
sourceDoc = sourceDoc.split('\n\n')

logging.debug('SOURCEDOC...\nSOURCEDOC...')
logging.debug(pprint.pformat((sourceDoc)))

# Now we process

infoScrape()

print('Compiling complete.\n')


#bigListDebugPrint()

#########EXCEL CODE BELOW

while True:
    try:
        desiredPath = input('Please enter desired path: ')
        os.chdir(desiredPath)
        break
    except:
        print('Invalid path!')
        continue


wb = openpyxl.Workbook()
sheet = wb['Sheet']

firstRow = {'A1': 'Company', 'B1' : 'Names', 'C1' : 'Emails', 'D1' : 'Phones', 'E1' : 'Address'}

for item in firstRow:
    sheet[item] = firstRow[item]

rowIndex = 2 # begins at the row under the column titles
i = 1
for Company in bigDict:

    working = 'Company %s' % i
    contactsize = len(bigDict[working]['contacts'])     #determines how many rows
    phonesize = len(bigDict[working]['phones'])         #are needed to display
    emailsize = len(bigDict[working]['emails'])         #all of this company's info
    rowsneeded = max(contactsize, phonesize, emailsize)

    sheet['A%s' % rowIndex] = bigDict[working]['name']

    j = 0
    for person in bigDict[working]['contacts']:
        sheet['B%s' % str(rowIndex + j)] = bigDict[working]['contacts']['person %s' % str(j+1)]
        j += 1

    j = 0
    for email in bigDict[working]['emails']:
        sheet['C%s' % str(rowIndex + j)] = bigDict[working]['emails']['email %s' % str(j+1)]
        j += 1

    j = 0
    for number in bigDict[working]['phones']:
        sheet['D%s' % str(rowIndex + j)] = bigDict[working]['phones']['number %s' % str(j+1)]
        j += 1

    sheet['E%s' % str(rowIndex)] = bigDict[working]['address']

    rowIndex = rowIndex + rowsneeded # moves "cursor" to next empty row
    i += 1

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

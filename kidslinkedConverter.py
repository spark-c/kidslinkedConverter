#This program scrapes a formatted document and converts the found text into an organized spreadsheet.
#Contact info must be in blocks with at least one empty line between companies.
#First line of a block must be company name.

#Collin Sparks 11/9/2019
#Last updated 12/4/2019
#python 3.8.0


# Format will be Company(0), Name(1), Email(2), Phone(3), Address(4)

import re, pprint, pyperclip, logging, openpyxl, os

logging.disable(logging.CRITICAL)

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


def infoScrape(): # gonna use this function for each info block

    i = 1
    for block in range(len(sourceDoc)):
        
        Company = 'Company %i' % i
        print(Company)
        collectedInfo = {Company: {   # list for temp storage of found data
                            'name':None}}
        
        logging.debug('BLOCK')
        logging.debug(block)

        working = sourceDoc[block]

        logging.debug('WORKING')
        logging.debug(working)
    
        rawData = working.splitlines()

        logging.debug('RAWDATA')
        logging.debug(pprint.pformat(rawData))
        logging.debug('...\n...\n')
              
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

        #NAMES
        foundNames = []
        for item in rawData:
             if nameRegex.search(item) != None:
                 foundNames.append(item.strip())

        j = 1
        collectedInfo[Company]['contacts'] = {}
        for person in foundNames:
            collectedInfo[Company]['contacts']['person %i' % j] = person
            j += 1

        #EMAILS
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

        #PHONES
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

        #ADDRESSES
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

            logging.debug('BIGLIST RESULT:\n %s \n' % Company)
            logging.debug(pprint.pformat(bigDict[Company]))
            logging.debug('\n')

            i += 1
            

def bigListDebugPrint(index = 'all'): # Default value for x is 0
    if index == 'all':
        for x in range(len(bigList)):
            print('PRINTING INDEX... %s \n' % x)
            print('COMPANY: ', bigList[x][0], '\n')
            print('NAMES: ', bigList[x][1], '\n')
            print('EMAILS: ', bigList[x][2], '\n')
            print('PHONES: ', bigList[x][3], '\n')
            print('ADDRESS: ', bigList[x][4], '\n')
            print('\n')
    else:
            print('PRINTING INDEX... %s \n' % index)
            print('COMPANY: ', bigList[index][0], '\n')
            print('NAMES: ', bigList[index][1], '\n')
            print('EMAILS: ', bigList[index][2], '\n')
            print('PHONES: ', bigList[index][3], '\n')
            print('ADDRESS: ', bigList[index][4], '\n')
            print('\n')






#DEFINITIONS


# NAME
nameRegex = re.compile(r'''
(^[a-zA-Z.]+\s    # newline and firstname with space; can include title
[a-zA-Z.]+            # name
\s?[a-zA-Z. ]*$)      # room for an optional middle name, space, ends with newline character
''', re.VERBOSE)


# EMAIL
emailRegex = re.compile(r'([a-zA-Z0-9._+\-]+@[a-zA-Z0-9._+\-]+)')


# PHONE
phoneRegex = re.compile(r'''
(\(?\d{3}\)?\s?-?        # area code, with or without parens or dash or space
\d{3}\s?-?\s?                  # first three digits and dash or no dash
\d{4})                   # final four digits
''', re.VERBOSE)


# ADDRESS
address1Regex = re.compile(r'''
(^\d+\s[a-zA-Z0-9.]+\s[a-zA-Z0-9 .]*)\s?
# 123 w broad street
''', re.VERBOSE)


address2Regex = re.compile(r'''
([a-zA-Z]+,?\s[a-zA-Z]+\s\d{5})
# columbus, OH 12345
''', re.VERBOSE)

# CHECK FOR 'DBA' IN BUSINESS NAME
dbaRegex = re.compile(r'(.*\sdba\s)(.*)')


bigDict = {} # this is the main directory



# Break the whole document into a giant list of blocks

sourceDoc = pyperclip.paste()

sourceDoc = sourceDoc.split('\r\n\r\n')

logging.debug('SOURCEDOC...\nSOURCEDOC...')
logging.debug(pprint.pformat((sourceDoc)))

# Now we proccess

infoScrape()

print('Compiling complete.\nCall "bigDict" for content.')


#bigListDebugPrint()

#########EXCEL CODE BELOW

desiredPath = input('Please enter desired path: ')

os.chdir(desiredPath)

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






wb.save(input('Enter filename: ') + '.xlsx') #could implement code to check whether '.xlsx' is already included




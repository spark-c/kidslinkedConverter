# Definition of the "COMPANY" class of objects for kidslinkedConverter script. 02/24/2021
# Python 3

class Company():

    def __init__(self, name, contacts=[], emails=[], phones=[], address=[]): # made to be updated one piece at a time; requires only name to begin.
        self.name = name
        self.contacts = contacts
        self.emails = emails
        self.phones = phones
        self.address = address

    def add(self, info_type, info):
        if info_type == 'contact':
            self.contacts.append(info)
        elif info_type == 'email':
            self.emails.append(info)
        elif info_type == 'phone':
            self.phones.append(info)
        elif info_type == 'address':
            self.address.append(info)

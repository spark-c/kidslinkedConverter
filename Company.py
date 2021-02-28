# Definition of the "COMPANY" class of objects for kidslinkedConverter script. 02/24/2021
# Python 3

import pprint

class Company():

    def __init__(self, name): # made to be updated one piece at a time; requires only name to begin.
        self.name = [name] # later this will be defined as a list with one item for ease in my excel code.
        self.contacts = []
        self.emails = []
        self.phones = []
        self.address = []

    def add(self, info_type, info):
        if info_type == 'contact':
            self.contacts.append(info)
        elif info_type == 'email':
            self.emails.append(info)
        elif info_type == 'phone':
            self.phones.append(info)
        elif info_type == 'address':
            self.address.append(info)
        elif info_type == 'name':
            self.name = info

    def printattrs(self):
        pprint.pprint(vars(self))

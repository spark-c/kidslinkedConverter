import re

# REGEX DEFINITIONS for kidslinkedConverter script

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

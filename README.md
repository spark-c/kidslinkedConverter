# kidslinkedConverter
# Collin Sparks, created 11/9/2019
# Python 3

The first longer script I wrote; it's very messy in retrospect, but it still does what I need it to do. Has bee rewritten for integration in another project; readme to be cleaned up and updated more thoroughly after.

This was written for a local company to save them hours of manually transferring data between .docx and .xlsx formats.

Takes large volumes of (generally formatted--see below) contact info from the clipboard and generates an excel spreadsheet.

REQUIREMENTS:
1. Contact info blocks must have one piece of data per line (phone#, email, etc)
2. The first line of the block must be the name of the company.
3. Each block must be separated by at least one empty line (\n\n).


# OUTDATED below... after rewrite, this is no longer the correct way to pass source material into script. To be updated...
USE:
1. Use CTRL+A to select all of the text in the document.
2. Use CTRL+C to copy the selected text to the clipboard.
2. Run kidslinkedConverter.py (will convert to .exe in the future)
3. When prompted, follow the on-screen instructions to enter the desired filename and path.

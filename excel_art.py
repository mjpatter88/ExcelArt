'''
Name: excel_art.py
Author: Michael Patterson
Date: June 19, 2013

My original plan was to use the "xlwt" package (http://www.python-excel.org/), but it did not seem to 
support a way to set custom background colors for a cell. I gave up on this approach, but then remembered
that excel has an xml file format. I played around with this a little bit and think I have the format
figured out. My plan is to write python that directly modifies the xml file to get the desired result.

base.xml contains the Excel generated xml with a few colored cells for reference.

Quirks:
Excel handles the background coloring by first creating a style in the <Styles> block, it then uses this
style ID to apply it to a cell in the <Table> block. I'm not sure if there is a limit on the number of
styles, but hopefully not:)

In the <Table> block the numbering is a little weird.
It starts at 1. It is grouped by rows, and they must be in numerical order. If a row or cell id is not
present, then it defaults to one greater than the previous one, or 1 if no previous one.
'''


def read_image():
	pass

def generate_excel_doc():
	pass

def run():
	read_image()
	generate_excel_doc()
	pass



if __name__ == "__main__":
	# get the file name from the command line
	run()
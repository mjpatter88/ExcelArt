'''
Name: excel_art.py
Author: Michael Patterson
Date: June 19, 2013

My original plan was to use the "xlwt" package (http://www.python-excel.org/), but it did not seem to 
support a way to set custom background colors for a cell. I gave up on this approach, but then remembered
that excel has an xml file format. I played around with this a little bit and think I have the format
figured out. My plan is to write python that directly modifies the xml file to get the desired result.

base.xml contains the Excel generated xml with a few colored cells for reference.

I'm using the ElementTree library (part of the standard lib) and more info is here: 
http://www.diveintopython3.net/xml.html

Quirks:
Excel handles the background coloring by first creating a style in the <Styles> block, it then uses this
style ID to apply it to a cell in the <Table> block. I'm not sure if there is a limit on the number of
styles, but hopefully not:)

In the <Table> block the numbering is a little weird.
It starts at 1. It is grouped by rows, and they must be in numerical order. If a row or cell id is not
present, then it defaults to one greater than the previous one, or 1 if no previous one.
'''

import xml.etree.ElementTree as ET

xml = None
DEBUG = True

def read_image():
	pass

def generate_excel_doc():
	xml.write("Test.xml")
	# We need the following lines to prepend the two starting lines to the xml so it's recognized as an Excel file
	f = open("Test.xml")
	temp = f.read()
	f.close()
	f = open("Test.xml", 'w')
	f.write("<?xml version=\"1.0\"?>\n")
	f.write("<?mso-application progid=\"Excel.Sheet\"?>\n")
	f.write(temp)
	f.close()
	pass

def add_style(id, color):
	'''
	This function adds a style to the xml tree. The following shows the format. I believe the styles start at ID=s62
	
	<Style ss:ID="s68">
   		<Interior ss:Color="#5600AD" ss:Pattern="Solid"/>
  	</Style>

	'''

	root = xml.getroot()
	styles = root.find("{urn:schemas-microsoft-com:office:spreadsheet}Styles")
	
	if DEBUG:
		print "Parent Element: ", styles
		print "Number of styles: ", len(styles)
		for child in styles:
			print(child.get("{urn:schemas-microsoft-com:office:spreadsheet}ID"))

	#Append the new style with the correct ID
	id_str = "s" + str(id)
	print id_str
	new_style = ET.SubElement(styles, "{urn:schemas-microsoft-com:office:spreadsheet}Style", 
									{"{urn:schemas-microsoft-com:office:spreadsheet}ID": "s70"})
	#Append the sub-element to the new style
	ET.SubElement(new_style, "{urn:schemas-microsoft-com:office:spreadsheet}Interior", 
									{"{urn:schemas-microsoft-com:office:spreadsheet}Color": "#123456",
									 "{urn:schemas-microsoft-com:office:spreadsheet}Pattern": "Solid"})
	return

def style_cell(row, col, style_id):
	'''
	This function colors the cell at row, col with the specified style. It can only be called once per cell, and the
	style must already exist.
	'''
	pass

def test():
	'''
	A test function that creates a spreadsheet 50x50 with a different color in each cell
	'''
	pass

def run():
	global xml

	xml = ET.parse("base.xml")
	add_style(70, "#123456")

	# read_image()
	generate_excel_doc()

	return 0



if __name__ == "__main__":
	# get the file name from the command line
	run()
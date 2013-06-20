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
num_cols = None
num_rows = None
DEBUG = True

def read_image():
	
	# Possibly set num_cols and num_rows to the #rows/#cols of pixels in the image. Is there a max?
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

def set_wrksht_props():

	# TODO: Set the needed properties for the excel file.
	# 1) The "ExpandedColumnCount" and "ExpandedRowCount" should be num_cols and num_rows
	# 2) The width of each col and height of each row
	pass

def add_style(id, color):
	'''
	This function adds a style to the xml tree. I believe the styles start at ID=s62, mine will start at s63 though.	
	Format:

	<Style ss:ID="s68">
   		<Interior ss:Color="#5600AD" ss:Pattern="Solid"/>
  	</Style>

	'''
	id_str = "s" + str(id)
	root = xml.getroot()
	styles = root.find("{urn:schemas-microsoft-com:office:spreadsheet}Styles")
	
	if DEBUG:
		print "Parent Element: ", styles
		print "Number of styles: ", len(styles)
		for child in styles:
			print(child.get("{urn:schemas-microsoft-com:office:spreadsheet}ID"))

	#Append the new style with the correct ID
	new_style = ET.SubElement(styles, "{urn:schemas-microsoft-com:office:spreadsheet}Style", 
									{"{urn:schemas-microsoft-com:office:spreadsheet}ID": id_str})
	#Append the sub-element to the new style
	ET.SubElement(new_style, "{urn:schemas-microsoft-com:office:spreadsheet}Interior", 
									{"{urn:schemas-microsoft-com:office:spreadsheet}Color": color,
									 "{urn:schemas-microsoft-com:office:spreadsheet}Pattern": "Solid"})
	return

def style_cell(row, col, style_id):
	'''
	This function colors the cell at row, col with the specified style. It must be called in \
	row-major order so the elements are added in the correct order.
	Format:

	<Row ss:Index="3" ss:AutoFitHeight="0">
    	<Cell ss:Index="4" ss:StyleID="s63"/>
    </Row>
	'''
	style_id_str = "s" + str(style_id)
	root = xml.getroot()
	wksht = root.find("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
	table = wksht.find("{urn:schemas-microsoft-com:office:spreadsheet}Table")

	if DEBUG:
		print "Worksheet Element: ", wksht
		print "Table Element: ", table
		print "Number of rows: ", len(table)
		for child in table:
			print(child.get("{urn:schemas-microsoft-com:office:spreadsheet}Index"))

	# The following code looks complicated, but it does one of three things
	# 1: Update an existing cell with the new styleID
	# 2: Add a cell to an existing row with the new style ID
	# 3: Add a row to the table with a new cell with the new style ID

	added = False
	row_index = 0
	# If the specified row exists, then just add the column (checking first that it doesn't exist)
	for tab_row in table.iter("{urn:schemas-microsoft-com:office:spreadsheet}Row"):
		row_index = int(tab_row.get("{urn:schemas-microsoft-com:office:spreadsheet}Index"))
		if row_index == row:	# We want to add to an existing row
			for tab_cell in tab_row.iter("{urn:schemas-microsoft-com:office:spreadsheet}Cell"):
				cell_index = int(tab_cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index"))
				if cell_index == col:	# We want to update the style of this element
					tab_cell.set("{urn:schemas-microsoft-com:office:spreadsheet}StyleID", style_id_str)
					added=True
			if not added:
				# add element with new style
				ET.SubElement(tab_row, "{urn:schemas-microsoft-com:office:spreadsheet}Cell", 
									{"{urn:schemas-microsoft-com:office:spreadsheet}Index": str(col),
									 "{urn:schemas-microsoft-com:office:spreadsheet}StyleID": style_id_str})
				added = True

	# If not, add the row and then the column (checking that no greater or equal column exists)
	if not added:
		if row_index < row:
			# add row element
			tab_row = ET.SubElement(table, "{urn:schemas-microsoft-com:office:spreadsheet}Row", 
											{"{urn:schemas-microsoft-com:office:spreadsheet}Index": str(row)})
			# add cell element with new style
			ET.SubElement(tab_row, "{urn:schemas-microsoft-com:office:spreadsheet}Cell", 
									{"{urn:schemas-microsoft-com:office:spreadsheet}Index": str(col),
									 "{urn:schemas-microsoft-com:office:spreadsheet}StyleID": style_id_str})
			added = True
		else:
			print "Row to add: ", row
			print "Last existing row: ", row_index
			raise Exception("Cannot add this row in this order.")
	return

def test():
	'''
	A test function that creates a spreadsheet 50x50 with a different color in each cell
	'''
	global xml

	xml = ET.parse("base.xml")
	set_wrksht_props()
	add_style(63, "#210672")
	add_style(64, "#960028")
	add_style(65, "#007241")
	add_style(66, "#A68A00")
	add_style(67, "#00AF64")
	add_style(68, "#FFD500")
	style_cell(1, 1, 63)
	style_cell(1, 2, 64)
	style_cell(1, 3, 65)
	style_cell(2, 1, 66)
	style_cell(3, 1, 67)
	style_cell(4, 1, 68)
	generate_excel_doc()
	return

def run():
	global xml

	xml = ET.parse("base.xml")
	read_image()
	set_wrksht_props()
	generate_excel_doc()

	return 0



if __name__ == "__main__":
	# get the file name from the command line
	# run()
	test()
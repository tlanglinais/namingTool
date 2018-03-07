# Relevant Info
# 1. Type of Report
# 2. Customer
# 3. Location
# 4. Sample Point
# 5. Sample Date

import os
from tkinter import filedialog
import openpyxl
import functions
import convert

# Promp File Select
fileName = filedialog.askopenfilename()
# print(f'\n\nOpening {fileName}')

# Need to retain the file extension for renaming
# I split by '.' because i don't know how long the file extension is: Ex. .pdf or .xlsx
fileSplit = fileName.split('.')

# Add a period back to the last element of the split name
fileExtension = '.' + fileSplit[len(fileSplit) - 1]


data_map = {
		# 'Bacteria Culture Analysis': {'abbreviation': 'BCA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		'Microbial Incubation Report': {'abbreviation': 'BCA', 'multipleFlag': 0, 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'B8', 'Sample Date': 'K7'},
		'Corrosion Coupon Analysis Report': {'abbreviation': 'CCA', 'multipleFlag': 1, 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A11', 'Sample Date': 'J5'},
		'Cold Finger Analysis': {'abbreviation': 'CFA', 'multipleFlag': 0, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': '', 'Sample Date': 'H5'},
		# 'Chloride Analysis': {'abbreviation': 'CLA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		'Comprehensive Crude Oil Analysis': {'abbreviation': 'COA', 'multipleFlag': 1, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A12', 'Sample Date': 'H12'},
		# 'Corrosion Selection Test': {'abbreviation': 'CST', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		'Comprehensive Oilfield Water Analysis': {'abbreviation': 'CWA', 'multipleFlag': 0, 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'H5'},
		'Iron and Manganese Analysis': {'abbreviation': 'FeMnA', 'multipleFlag': 1, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		'Iron, Manganese, & Chloride Analysis': {'abbreviation': 'FeMnA', 'multipleFlag': 1, 'Customer': 'C6', 'Location': 'D7', 'Sample Point': 'A10', 'Sample Date': 'J5'},
		# 'Gas Analysis': {'abbreviation': 'GSA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		'Membrane Filter Analysis': {'abbreviation': 'MFA', 'multipleFlag': 0, 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5'},
		'Oil & Grease in Water by Infrared Analysis': {'abbreviation': 'OGA', 'multipleFlag': 1, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A10', 'Sample Date': 'H5'},
		# 'Pipe Failure Analysis': {'abbreviation': 'PFA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		'Scale Coupon Analysis Report': {'abbreviation': 'SCA', 'multipleFlag': 1, 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A10', 'Sample Date': 'K5'},
		'Quantitative Solids Identification': {'abbreviation': 'QSI', 'multipleFlag': 0, 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5'},
		# 'Solids ID Analysis': {'abbreviation': 'SIDA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Scale Inhibitor Residual': {'abbreviation': 'SIR', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Total Oil and Grease': {'abbreviation': 'TOG', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''}
		}

# # Check to see if its a .xls spreadsheet
# if fileExtension == '.xls':
# 	wb = convert.open_xls_as_xlsx(fileName)
# 	sheet = wb.active

# else:
# load .xlsx workbook
wb = openpyxl.load_workbook(fileName)
sheet = wb.active

# print(f"C1 = {sheet['C1'].value}")
# print(f"D1 = {sheet['D1'].value}")

# find if the report name is in cell C1 or D1
if sheet['C1'].value in data_map:
	reportName = sheet['C1'].value
elif sheet['D1'].value in data_map:
	reportName = sheet['D1'].value
else:
	print('---Error: Cannot find the header---')

# Grab customer name and format for renaming
customer = functions.format_name(sheet[data_map[reportName]['Customer']].value)

# Grab location and format for renaming
location = functions.format_name(sheet[data_map[reportName]['Location']].value)

# Grab sample point and format for renaming
# ----------------------------------------------------------NEED TO CHECK IF THERE ARE MULTIPLE SAMPLE POINTS
if data_map[reportName]['multipleFlag']: 
	# Need to see if there are more than one sample points
	sample1 = data_map[reportName]['Sample Point']
	# Chop up the cell name and increment to 1 cell below
	sample2 = 'A' + str(int(sample1[1:]) + 1)
	if sheet[sample2].value != 'None':
		samplePoint = 'Multiple_'

	else:
		print("Something went wrong with the 'Multiple' name")
else:
	samplePoint = functions.format_name(sheet[data_map[reportName]['Sample Point']].value)

# Grab sample date and format for renaming
sampleDate = functions.format_date(sheet[data_map[reportName]['Sample Date']].value)

# Grab the report abbreviation from the reports dictionary
reportAbr = data_map[reportName]['abbreviation']

newName = reportAbr + '_' + customer + location + samplePoint + sampleDate + fileExtension

os.rename(fileName, 'C:\\Users\\tsl\\desktop\\Lab Data\\Renamed\\' + newName)

print(f'\nRenamed:\n{fileName}\nto:\n{newName}\n\n')
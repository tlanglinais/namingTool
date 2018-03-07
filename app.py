# Relevant Info
# 1. Type of Report - Index = 13
# 2. Customer - Index = 23
# 3. Location - Index = 29
# 4. Sample Point - Index = 33
# 5. Sample Date - Index = 26

import os
import datetime
from tkinter import filedialog
import openpyxl

# Promp File Select
fileName = filedialog.askopenfilename()
print(f'Opening {fileName}')

# Need to retain the file extension for renaming
# I split by '.' because i don't know how long the file extension is: Ex. .pdf or .xlsx
fileSplit = fileName.split('.')

# Add a period back to the last element of the split name
fileExtension = '.' + fileSplit[len(fileSplit) - 1]



reports = {'Bacteria Culture Analysis': 'BCA',
		'Microbial Incubation Report': 'BCA',
		'Corrosion Coupon Analysis Report': 'CCA',
		'Cold Finger Analysis': 'CFA',
		'Chloride Analysis': 'CLA',
		'Comprehensive Crude Oil Analysis': 'COA',
		'Corrosion Selection Test': 'CST',
		'Complete Water Analysis': 'CWA',
		'Comprehensive Oilfield Water Analysis': 'CWA',
		'Iron and Manganese Analysis': 'FeMnA',
		'Iron, Manganese, & Chloride Analysis': 'FeMnA',
		'Gas Analysis': 'GSA',
		'Membrane Filter Analysis': 'MFA',
		'Oil & Grease in Water by Infrared Analysis': 'OGA',
		'Oil and Grease Analysis Report': 'OGA',
		'Pipe Failure Analysis': 'PFA',
		'Scale Coupon Analysis': 'SCA',
		'Solids ID Analysis': 'SIDA',
		'Quantitative Solids Identification': 'SIDA',
		'Scale Inhibitor Residual': 'SIR',
		'Total Oil and Grease': 'TOG'
		}

data_map = {
		'Microbial Incubation Report': {'multipleFlag': 'N', 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'B8', 'Sample Date': 'K7'},
		#Corrosion Coupon can have multiple sample points!!!
		'Corrosion Coupon Analysis Report': {'multipleFlag': 'Y', 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A11', 'Sample Date': 'J5'},
		'Cold Finger Analysis': {'multipleFlag': 'N', 'Customer': 'C6', 'Location': 'C7', 'Sample Point': '', 'Sample Date': 'H5'},
		#Corrosion Coupon can have multiple sample points!!!
		'Comprehensive Crude Oil Analysis': {'multipleFlag': 'Y', 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A12', 'Sample Date': 'H12'},
		'Comprehensive Oilfield Water Analysis': {'multipleFlag': 'N', 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'H5'},
		'Comprehensive Crude Oil Analysis': {'multipleFlag': 'N', 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A12', 'Sample Date': 'H12'},
		'Iron and Manganese Analysis': {'multipleFlag': 'N', 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# Iron, Manganese can have multiple sample points!!!
		'Iron, Manganese, & Chloride Analysis': {'multipleFlag': 'Y', 'Customer': 'C6', 'Location': 'D7', 'Sample Point': 'A10', 'Sample Date': 'J5'},
		'Membrane Filter Analysis': {'multipleFlag': 'N', 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5'},
		# Oil & Grease can have multiple sample points!!!
		'Oil & Grease in Water by Infrared Analysis': {'multipleFlag': 'Y', 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A10', 'Sample Date': 'H5'},
		# Scale Coupon Analysis can have multiple sample points!!!
		'Scale Coupon Analysis': {'multipleFlag': 'Y', 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A10', 'Sample Date': 'K5'},
		'Quantitative Solids Identification': {'multipleFlag': 'N', 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5'},

		# 'Bacteria Culture Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Chloride Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Corrosion Selection Test': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Complete Water Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Gas Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Oil and Grease Analysis Report': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Pipe Failure Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Solids ID Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Scale Inhibitor Residual': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Total Oil and Grease': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''}
		}

# load workbook
wb = openpyxl.load_workbook(fileName)
sheet = wb.active

# print(f"C1 = {sheet['C1'].value}")
# print(f"D1 = {sheet['D1'].value}")

# find if the report name is in cell C1 or D1
if sheet['C1'].value in reports:
	reportName = sheet['C1'].value
elif sheet['D1'].value in reports:
	reportName = sheet['D1'].value
else:
	print('Error: Report Not Found in Current Dictionary')

# Grab customer name and format for renaming
customer = format_name(sheet[data_map[reportName]['Customer']].value)

# Grab location and format for renaming
location = format_name(sheet[data_map[reportName]['Location']].value)

# Grab sample point and format for renaming
# ----------------------------------------------------------NEED TO CHECK IF THERE ARE MULTIPLE SAMPLE POINTS
samplePoint = format_name(sheet[data_map[reportName]['Sample Point']].value)

# Grab sample date and format for renaming
sampleDate = format_date(sheet[data_map[reportName]['Sample Date']].value)

# Grab the report abbreviation from the reports dictionary
reportAbr = reports[reportName]

newName = reportAbr + '_' + customer + location + samplePoint + sampleDate + fileExtension

os.rename(fileName, 'C:\\Users\\tsl\\desktop\\Lab Data\\Renamed\\' + newName)
# os.move(newName, 'C:\\Users\\tsl\\desktop\\Lab Data\\Renamed')
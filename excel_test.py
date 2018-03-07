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

def check_space(var):
	"""Checks to see if passed string has an empty space before or after the string"""
	# Check to find a beginning space
	if var[0] == ' ':
		var = var[1:]
	# Check to find an ending space
	if var[len(var) - 1] == ' ':
		var = var[:len(var) - 1]
	return var

def format_date(date):
	"""Accepts in a date and formats it according to the PICS naming convention"""

	# print(f'Old Date: {date}')
	# print(f'Date type: {type(date)}')

	# Check if date is a datetime object
	if isinstance(date, datetime.date):
		
		# variables come over as integers
		# must convert to string before formatting
		month = preceding_zero(str(date.month))
		day = preceding_zero(str(date.day))
		year = str(date.year)[2:]

	# Check if date is a string
	elif isinstance(date, str):
		# Check to see if date is formatted like: MM/DD/YYYY
		if '/' in date:
		# Splits the date up by the / character
			month, day, year = date.split('/')
			month = preceding_zero(month)
			day = preceding_zero(day)

			# Slice off the first 2 digits of the year
			year = year[2:]
			

		# Check to see if date is formatted like: January 1, 2000
		elif ',' in date:

			months_map = {'January': '01',
						'February': '02',
						'March': '03',
						'April': '04',
						'May': '05',
						'June': '06',
						'July': '07',
						'August': '08',
						'September': '09',
						'October': '10',
						'November': '11',
						'December': '12'
						}

			# Split the date up by spaces
			month, day, year = date.split(' ')
			month = months_map[month]
			
			# Slice off the first 2 digits of the year
			year = year[2:]

			# Remove the comma from the second list element(the day)
			if(len(day)) == 3:
				day = preceding_zero(day[:2])

			elif(len(day)) == 2:
				day = preceding_zero(day[:1])
			else:
				print("Something's wrong with this date...")

	else:
		print('Date is not a valid format')
		return 0

	# Return the formatted date
	return year + month + day


def preceding_zero(date):
	"""Accepts a month or day, to check if it needs to have a preceding 0"""

	# Checks to see if the date is only 1 digit
	if len(date) == 1:
		date = '0' + date
	return date

def format_name(data):
	"""Accepts a string to be formatted for the file renaming"""

	# print(f'Data type: {type(data)}')

	if isinstance(data, str):
		return data + '_'

	else:
		return ''


reports = {'Bacteria Culture Analysis': 'BCA',
		'Microbial Incubation Report': 'BCA',
		'Corrosion Coupon Analysis Report': 'CCA',
		'Cold Finger Analysis': 'CFA',
		'Chloride Analysis': 'CLA',
		'Crude Oil Analysis': 'COA',
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
		'Microbial Incubation Report': {'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'B8', 'Sample Date': 'K7'},
		#Corrosion Coupon can have multiple sample points!!!
		'Corrosion Coupon Analysis Report': {'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A11', 'Sample Date': 'J5'},
		'Comprehensive Oilfield Water Analysis': {'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'H5'},
		'Comprehensive Crude Oil Analysis': {'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A12', 'Sample Date': 'H12'},
		'Iron and Manganese Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# Iron, Manganese can have multiple sample points!!!
		'Iron, Manganese, & Chloride Analysis': {'Customer': 'C6', 'Location': 'D7', 'Sample Point': 'A10', 'Sample Date': 'J5'},
		'Membrane Filter Analysis': {'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5'},
		# Oil & Grease can have multiple sample points!!!
		'Oil & Grease in Water by Infrared Analysis': {'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A10', 'Sample Date': 'H5'},
		# Scale Coupon Analysis can have multiple sample points!!!
		'Scale Coupon Analysis': {'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A10', 'Sample Date': 'K5'},
		'Quantitative Solids Identification': {'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5'},

		# 'Bacteria Culture Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Cold Finger Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Chloride Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
		# 'Crude Oil Analysis': {'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': ''},
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
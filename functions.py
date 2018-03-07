# File to store the functions for testing

import datetime

def check_space(var):
	"""Checks to see if passed string has an empty space before or after the string"""
	# Check to find a beginning space
	if isinstance(var, str):
		if var[0] == ' ':
			var = var[1:]
		# Check to find an ending space
		if var[len(var) - 1] == ' ':
			var = var[:len(var) - 1]
		return var
	else:
		print(f'check_space cannot iterate over {type(var)}')

def format_name(data):
	"""Accepts a string to be formatted for the file renaming"""
	# print(f'Data type: {type(data)}')

	prohibited_characters = ['~', '#', '%', '&', '*', '{', '}', '\\', ':', '<', '>', '?', '/', '+', '|', '"']

	if isinstance(data, str):
		for i in range(len(prohibited_characters)):
			data = data.replace(prohibited_characters[i], '')

		return check_space(data) + '_'

	else:
		return ''

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
		print(f'{date} is not a valid date format')
		return ''

	# Return the formatted date
	return year + month + day


def preceding_zero(date):
	"""Accepts a month or day, to check if it needs to have a preceding 0"""

	# Checks to see if the date is only 1 digit
	if len(date) == 1:
		date = '0' + date
	return date
infile = input('What file do you want to open? ')
# infile = open(infile, 'r+')
infile = open('test.txt', 'r+')

infile.read()

line = 1
date = ''
for i in infile:
	if i == 11:
		date = infile.readline(i)


infile = open('test.txt', 'r+')
text = infile.read()
len(text)
lines = []
lines = text.split('\n')

for i in range(len(lines)):
	lines.remove('')



import os
os.chdir('C:\\Users\\tsl\\desktop\\test')

# File 1
text = open('test.txt', 'r').read()
text = text.split('\n')
len(text)
# 1198

# File 2
text2 = open('test2.txt', 'r').read()
text2 = text2.split('\n')
len(text2)
# 1195

# File 3
text3 = open('test3.txt', 'r').read()
text3 = text3.split('\n')
len(text3)
# 1204

# Relevant Info
# 1. Type of Report - Index = 13
# 2. Customer - Index = 23
# 3. Location - Index = 29
# 4. Sample Point - Index = 33
# 5. Sample Date - Index = 26

# report = check_space(text[13])
# customer = check_space(text[23])
# location = check_space(text[29])
# samplePoint = check_space(text[33])
# sampleDate = check_space(text[26])


import os
from tkinter import filedialog
# from tkinter import *

filename = filedialog.askopenfilename()
text = open('test.txt', 'r').read()
text = text.split('\n')

reports = {'Bacteria Culture Analysis': 'BCA',
		'Corrosion Coupon Analysis': 'CCA',
		'Cold Finger Analysis': 'CFA',
		'Chloride Analysis': 'CLA',
		'Crude Oil Analysis': 'COA',
		'Corrosion Selection Test': 'CST',
		'Complete Water Analysis': 'CWA',
		'Iron and Manganese Analysis': 'FeMnA',
		'Gas Analysis': 'GSA',
		'Membrane Filter Analysis': 'MFA',
		'Oil and Grease Analysis': 'OGA',
		'Pipe Failure Analysis': 'PFA',
		'Scale Coupon Analysis': 'SCA',
		'Solids ID Analysis': 'SIDA',
		'Scale Inhibitor Residual': 'SIR',
		'Total Oil and Grease': 'TOG'
		}

print(reports[check_space(text[13])] + '_' + 
		check_space(text[23]) + '_' + 
		check_space(text[29]) + '_' + 
		check_space(text[33]) + '_' + 
		format_date(text[26]))

def check_space(var):
	"""Checks to see if passed string has an empty space before the string"""
	if var[0] == ' ':
		var = var[1:]
	return var

def format_date(date):
	"""Accepts in a date and formats it to JMs specifications"""

	# Splits the date up by the / character
	month, day, year = date.split('/')
	month = preceding_zero(month)
	day = preceding_zero(day)

	# Slice off the first 2 digits of the year
	year = year[2:]

	return year + month + day


def preceding_zero(date):
	"""Accepts a month or day, to check if it needs to have a preceding 0"""

	# Checks to see if the date is only 1 digit
	if len(date) == 1:
		date = '0' + date
	return date



 
filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))

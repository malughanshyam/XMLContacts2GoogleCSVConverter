#!/usr/bin/python

'''
	Created on Feb 14, 2016
	@author		: Ghanshyam Malu
	@desc     	: Convert contacts in XML format to Google CSV
	@Usage    	: Execute the python file
					$ python ContactsXML2CSV.py <xmlFileName>
	@Version  	: Uses Python 2.7
'''

import sys,csv,os,time,traceback
import xml.etree.ElementTree as ET

def usage():
	print "\n","-"*30
	print "Usage: "
	print "-"*30,"\n"
	print "python",sys.argv[0], "<xmlFile>\n"
	print "-"*30,"\n"
	sys.exit()
	
def parseXML(fileName):
	print "Parsing", fileName
	try:
		tree = ET.parse(fileName)
		root = tree.getroot()
		
		# Get all contact nodes
		allParsedContacts = []

		for contact in root.findall("contact"):		
			
			
			
			#structuredName Tag
			displayName = contact.findtext('structuredName/displayName',default='')
			givenName = contact.findtext('structuredName/givenName', default='')
			familyName = contact.findtext('structuredName/familyName', default='')
			prefixName = contact.findtext('structuredName/prefixName', default='')
			suffixName = contact.findtext('structuredName/suffixName', default='')
			middleName = contact.findtext('structuredName/middleName', default='')
			
			# Nickname
			nickName = contact.findtext('nickName', default='')

			#Birthday
			birthday = ''
			events = contact.find('Events')
			allEvents = list(events.iter("event"))
			for event in allEvents:
				if event.findtext('type') == 'birthday' and event.findtext('start_date') != None:
					birthdayFromEpoch = float(event.findtext('start_date'))
					birthday = time.strftime('%Y-%m-%d',  time.gmtime(birthdayFromEpoch/1000.))
			
			# Create a new Dictionary
			newContactDict = {}
			newContactDict['Name']=displayName
			newContactDict['Given Name'] = givenName
			newContactDict['Family Name'] = familyName
			newContactDict['Name Prefix'] = prefixName
			newContactDict['Name Suffix'] = suffixName
			newContactDict['Initials'] = middleName
			newContactDict['Birthday'] = birthday
			newContactDict['Nickname'] = nickName
			
			
			print newContactDict,"\n"
			
			'''
			newContactDict['Name']='Puneet Loya'
			newContactDict['Given Name'] = 'Puneet'
			newContactDict['Family Name'] = 'Loya'
			newContactDict['Birthday'] = '1990-06-03'
			newContactDict['E-mail 1 - Type'] = '* Other'
			newContactDict['E-mail 1 - Value'] = 'loya.puneet@gmail.com'
			newContactDict['E-mail 2 - Type'] = 'Other'
			newContactDict['E-mail 2 - Value'] = 'puneetloya123@yahoo.co.in ::: loyapuneet@gmail.com'
			newContactDict['Phone 1 - Type'] = 'Home'
			newContactDict['Phone 1 - Value'] = '+1 812-391-9840'
			newContactDict['Phone 2 - Type'] = 'Mobile'
			newContactDict['Phone 2 - Value'] = '+91 81 47 834462'
			newContactDict['Phone 3 - Type'] = 'Other'	
			newContactDict['Phone 3 - Value'] = '+91 96 20 309099'	
			'''
			
			
		
	except :
		print "Error parsing the file", fileName, ":\n", sys.exc_info()[0],"\n", sys.exc_info()[1]
		print sys.exc_info()
		# traceback.print_stack()
		traceback.print_exc(file=sys.stdout)
		#sys.exit()
	else:
		return root
	

def exportCSV(fileName):
	'''
	Export to CSV
	'''
	global GoogleCSVFieldNames
	newFileName = os.path.splitext(fileName)[0]+'.csv'
	print newFileName
	try:
		with open(newFileName, 'w') as csvfile:

			writer = csv.DictWriter(csvfile, fieldnames=GoogleCSVFieldNames, lineterminator='\n')
			writer.writeheader()
			newContactDict = {}
			newContactDict['Name']='Puneet Loya'
			newContactDict['Given Name'] = 'Puneet'
			newContactDict['Family Name'] = 'Loya'
			newContactDict['Birthday'] = '1990-06-03'
			newContactDict['E-mail 1 - Type'] = '* Other'
			newContactDict['E-mail 1 - Value'] = 'loya.puneet@gmail.com'
			newContactDict['E-mail 2 - Type'] = 'Other'
			newContactDict['E-mail 2 - Value'] = 'puneetloya123@yahoo.co.in ::: loyapuneet@gmail.com'
			newContactDict['Phone 1 - Type'] = 'Home'
			newContactDict['Phone 1 - Value'] = '+1 812-391-9840'
			newContactDict['Phone 2 - Type'] = 'Mobile'
			newContactDict['Phone 2 - Value'] = '+91 81 47 834462'
			newContactDict['Phone 3 - Type'] = 'Other'	
			newContactDict['Phone 3 - Value'] = '+91 96 20 309099'	
			writer.writerow(newContactDict)
			
	except:
		print "Error exporting CSV:", ":\n", sys.exc_info()[0],"\n", sys.exc_info()[1]

def main():
	if len(sys.argv) != 2:
		usage()
	
	#Global
	global GoogleCSVFieldNames
	GoogleCSVFieldNames = ['Name','Given Name','Family Name','Name Prefix','Name Suffix','Initials','Nickname','Birthday','Notes','E-mail 1 - Type','E-mail 1 - Value','E-mail 2 - Type','E-mail 2 - Value','Phone 1 - Type','Phone 1 - Value','Phone 2 - Type','Phone 2 - Value','Phone 3 - Type','Phone 3 - Value','Website 1 - Type','Website 1 - Value']
		
	parseXML(sys.argv[1])
	#exportCSV(sys.argv[1])

if __name__ == "__main__": 
	main()


'''
Name,Given Name,Family Name,Name Prefix,Name Suffix,Initials,Nickname,Birthday,Notes,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,Phone 1 - Value,Phone 2 - Type,Phone 2 - Value,Website 1 - Type,Website 1 - Value


'Name','Given Name','Family Name','Name Prefix','Name Suffix','Initials','Nickname','Birthday','Notes','E-mail 1 - Type','E-mail 1 - Value','E-mail 2 - Type','E-mail 2 - Value','Phone 1 - Type','Phone 1 - Value','Phone 2 - Type','Phone 2 - Value','Website 1 - Type','Website 1 - Value'
'''
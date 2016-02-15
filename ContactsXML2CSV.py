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
	'''
		Parse the Contacts XML and fetch all the contact details
		@fileName : XML File path
		@return : List of dictionaries for every contact
	'''
	global maxPhoneCount, maxEmailCount, maxWebsiteCount
	print "\n\tProcessing contacts from", fileName,".....",
	try:
		tree = ET.parse(fileName)
		root = tree.getroot()
		
		# Get all contact nodes
		allParsedContacts = []

		for contact in root.findall("contact"):		
			
			
			
			#structuredName Tag
			displayName = contact.findtext('structuredName/displayName',default='').encode('utf-8')
			givenName = contact.findtext('structuredName/givenName', default='').encode('utf-8')
			familyName = contact.findtext('structuredName/familyName', default='').encode('utf-8')
			prefixName = contact.findtext('structuredName/prefixName', default='').encode('utf-8')
			suffixName = contact.findtext('structuredName/suffixName', default='').encode('utf-8')
			middleName = contact.findtext('structuredName/middleName', default='').encode('utf-8')
			
			# Nickname
			nickName = contact.findtext('nickName', default='').encode('utf-8')

			# Birthday
			birthday = ''
			events = contact.find('Events')
			allEvents = list(events.iter("event"))
			for event in allEvents:
				if event.findtext('type') == 'birthday' and event.findtext('start_date'):
					birthdayFromEpoch = float(event.findtext('start_date'))
					birthday = time.strftime('%Y-%m-%d',  time.gmtime(birthdayFromEpoch/1000.)).encode('utf-8')
			
			# Notes
			note = contact.findtext('note', default='').encode('utf-8')
			
			# Emails
			emails = contact.find('emails')
			allEmails = list(emails.iter("email"))
			emailDict = {}
			for i in range(len(allEmails)):
				if allEmails[i].findtext('address') :
					emailDict['E-mail '+ str(i+1) + ' - Type'] = allEmails[i].findtext('type',default='Other').title().encode('utf-8')
					emailDict['E-mail '+ str(i+1) + ' - Value'] = allEmails[i].findtext('address',default='').encode('utf-8')			
			
			# Phone
			phones = contact.find('phones')
			allPhones = list(phones.iter("phone"))
			phoneDict = {}
			for i in range(len(allPhones)):
				if allPhones[i].findtext('number') :
					phoneDict['Phone '+ str(i+1) + ' - Type'] = allPhones[i].findtext('type',default='Other').title().encode('utf-8')
					phoneDict['Phone '+ str(i+1) + ' - Value'] = allPhones[i].findtext('number',default='').encode('utf-8')	
					
					
			# Website
			WebSites = contact.find('WebSites')
			allWebSites = list(WebSites.iter("website"))
			websiteDict = {}
			for i in range(len(allWebSites)):
				if allWebSites[i].findtext('url') :
					websiteDict['Website '+ str(i+1) + ' - Type'] = allWebSites[i].findtext('type',default='Other').title().encode('utf-8')
					websiteDict['Website '+ str(i+1) + ' - Value'] = allWebSites[i].findtext('url',default='').encode('utf-8')	
					
			# Create a new Dictionary
			newContactDict = {}
			newContactDict['Name']=displayName
			newContactDict['Given Name'] = givenName
			newContactDict['Family Name'] = familyName
			newContactDict['Name Prefix'] = prefixName
			newContactDict['Name Suffix'] = suffixName
			newContactDict['Additional Name'] = middleName
			newContactDict['Birthday'] = birthday
			newContactDict['Nickname'] = nickName
			newContactDict['Notes'] = note
			
			# Add emails
			for k,v in emailDict.iteritems():
				newContactDict[k] = v

			# Add Phone No.s
			for k,v in phoneDict.iteritems():
				newContactDict[k] = v

				
			# Add WebSites
			for k,v in websiteDict.iteritems():
				newContactDict[k] = v
				
			allParsedContacts.append(newContactDict)


			# Update Maximum found counts for the email, phone and website to create CSV header.
			if len(emailDict)/2 > maxEmailCount:
				maxEmailCount = len(emailDict)/2
							
			if len(phoneDict)/2 > maxPhoneCount:
				 maxPhoneCount = len(phoneDict)/2
							
			if len(websiteDict)/2 > maxWebsiteCount:
				maxWebsiteCount = len(websiteDict)/2

			
			'''
			Sample
			newContactDict = {}
			newContactDict['Name']='Sheldon Cooper'
			newContactDict['Given Name'] = 'Sheldon'
			newContactDict['Family Name'] = 'Cooper'
			newContactDict['Birthday'] = '1990-06-03'
			newContactDict['E-mail 1 - Type'] = '* Other'
			newContactDict['E-mail 1 - Value'] = 'Cooper.Sheldon@gmail.com'
			newContactDict['E-mail 2 - Type'] = 'Other'
			newContactDict['E-mail 2 - Value'] = 'Cooper.Sheldon2@gmail.com ::: Cooper.Sheldon3@gmail.com'
			newContactDict['Phone 1 - Type'] = 'Home'
			newContactDict['Phone 1 - Value'] = '+1 812-391-8731'
			newContactDict['Phone 2 - Type'] = 'Mobile'
			newContactDict['Phone 2 - Value'] = '+91 81 47 00000'
			newContactDict['Phone 3 - Type'] = 'Other'	
			newContactDict['Phone 3 - Value'] = '+91 96 20 00000'	
			'''

	except :
		print "Error parsing the file", fileName, ":\n", sys.exc_info()[0],"\n", sys.exc_info()[1]
		traceback.print_exc(file=sys.stdout)
		sys.exit()
	else:
		print "Complete !!"	
		return allParsedContacts
	

def exportCSV(fileName, GoogleCSVFieldNames, allParsedContacts):
	'''
	Export to CSV
	'''	
	newFileName = os.path.splitext(fileName)[0]+'.csv'
	print "\n\tExporting contacts to", newFileName,"....",
	
	try:
		with open(newFileName, 'w') as csvfile:
			writer = csv.DictWriter(csvfile, fieldnames=GoogleCSVFieldNames, lineterminator='\n')
			writer.writeheader()
			
			for contact in allParsedContacts:
				writer.writerow(contact)
		print "Complete !!\n\n\tEnjoy !!!!"	
		
	except:
		print "Error exporting CSV:", ":\n", sys.exc_info()[0],"\n", sys.exc_info()[1]
		traceback.print_exc(file=sys.stdout)
		sys.exit()

def main():
	if len(sys.argv) != 2:
		usage()
	
	
	
	#Global
	global screenWidth, maxEmailCount, maxPhoneCount, maxWebsiteCount
	maxPhoneCount=0
	maxEmailCount=0
	maxWebsiteCount = 0
	
	GoogleCSVFieldNames = ['Name','Given Name','Additional Name','Family Name','Name Prefix','Name Suffix','Nickname','Birthday','Notes']
	
	# Email, Phone, Website fields added after finding the maximum number found.
	# Sample format below:
	#'E-mail 1 - Type','E-mail 1 - Value','E-mail 2 - Type','E-mail 2 - Value',
	#'Phone 1 - Type','Phone 1 - Value','Phone 2 - Type','Phone 2 - Value','Phone 3 - Type','Phone 3 - Value',
	#'Website 1 - Type','Website 1 - Value'
	
	screenWidth = 90
	
	print
	print '{:^{screenWidth}}'.format('{:=^{w}}'.format('', w = screenWidth-10), screenWidth=screenWidth)
	print '{:^{screenWidth}}'.format('{:^{w}}'.format('Welcome to XML to Google Contacts CSV Converter', w = screenWidth-10), screenWidth=screenWidth)
	print '{:^{screenWidth}}'.format('{:=^{w}}'.format('', w = screenWidth-10), screenWidth=screenWidth)

	# Parse the given XML File and fetch all contact details
	allParsedContacts = parseXML(sys.argv[1])
	
	# Update GoogleCSVFieldNames based on maxEmailCount, maxPhoneCount, maxWebsiteCount
	for i in range(maxEmailCount):
		GoogleCSVFieldNames.append('E-mail '+ str(i+1) + ' - Type')
		GoogleCSVFieldNames.append('E-mail '+ str(i+1) + ' - Value')					

	for i in range(maxPhoneCount):
		GoogleCSVFieldNames.append('Phone '+ str(i+1) + ' - Type')
		GoogleCSVFieldNames.append('Phone '+ str(i+1) + ' - Value')					

	for i in range(maxWebsiteCount):
		GoogleCSVFieldNames.append('Website '+ str(i+1) + ' - Type')
		GoogleCSVFieldNames.append('Website '+ str(i+1) + ' - Value')					
	
	# Export to Google CSV
	exportCSV(sys.argv[1], GoogleCSVFieldNames, allParsedContacts)

if __name__ == "__main__": 
	main()
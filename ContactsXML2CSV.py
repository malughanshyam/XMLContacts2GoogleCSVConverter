#!/usr/bin/python

'''
	Created on Feb 14, 2016
	@author		: Ghanshyam Malu
	@desc     	: Convert contacts in XML format to Google CSV
	@Usage    	: Execute the python file
					$ python ContactsXML2CSV.py <xmlFileName>
	@Version  	: Uses Python 2.7
'''

import sys,csv,os
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
		print root
		
		
	except:
		print "Error parsing the file", fileName
		sys.exit()
	else:
		return root
	
	
def exportCSV(fileName):
	newFileName = os.path.splitext(fileName)[0]+'.csv'
	print newFileName
	
	with open(newFileName, 'w') as csvfile:
		fieldnames = ['first_name', 'hell','last_name']
		fieldnames = ('first_name', 'hell','last_name')
		fieldnames.append("fmasdf")
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

		writer.writeheader()
		writer.writerow({'first_name': 'Baked', 'last_name': 'Beans'})
		writer.writerow({'first_name': 'Lovely', 'last_name': 'Spam'})
		writer.writerow({'first_name': 'Wonderful', 'last_name': 'Spam'})
	

def main():
# display some lines
	if len(sys.argv) != 2:
		usage()
		
	parseXML(sys.argv[1])
	exportCSV(sys.argv[1])

if __name__ == "__main__": main()
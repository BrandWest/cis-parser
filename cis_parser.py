#Imports
from PyPDF2 import PdfReader
import os, sys, getopt, re
import csv, openpyxl 
#Tkinter GUI creation
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfile
tk.Tk().withdraw() 


def write_file(control):
	"""
	write_file Writes to the output file

	The control dictionary is written to the CSV and saves it to the outputFile

	:param control: The Dictionary storing the extract text
	:type control: dictionary
	"""
	with open (outputFile, "a", newline='') as csvFile:
		writer = csv.DictWriter(csvFile, fieldnames=control.keys())
		writer.writerow(control)



def checkForControl(text, outputFile):
	"""
	checkForControl The regex logic to locate specific cases from the CIS benchmarks

	The Regex is used to determine if the line provided from "text" is one of the locations from the CIS benchmarks Title, Profile Applicability, Description, Rationale, Impact, Audit, Remediation, Default Value, CIS Controls
	Some additional regex is used for specific edge cases between the tested documents.
	:param text: The file reader object containing the text data extracted from the PDF document
	:type text: file reader object
	:param outputFile: location to the output file 
	:type outputFile: string
	"""

	#Flag variables
	nextParagraphFlag = 0 	# Provides identification for continued append to line variable
	nextControl = 0 		# First control or not variable
	ack = 0 				# Checks to see the number of "acknowlegements" found, used to skip the TOC and beginning data.
	contentFlag = 0 		# Provides a flag to skip the title sections 
	
	#The Control header fields dictionary
	control = {
		'title' : '',
		'profile' : '',
		'description' : '',
		'rationale' : '',
		'impact' : '',
		'audit' : '',
		'remediation' : '',
		'defaultValue' : '',
		'references' : '',
		'cisControls' : ''
	}

	#Numerous Regexs to check specific edge/cases
	title2 = re.compile("(\(|.*)(Automated|Manual)\)$") 			# Checks if the next line contains the (Automated or Manual) words
	title = re.compile("^(\"|)(\d*\.){1,}\d.+ (\([A-Z]\d\)|).*") 	# Checks to see if the words are apart of the title
	profile = re.compile("Profile Applicability( |:)")				# Identifies the profile applicablity section
	description = re.compile("Description:")						# Identifies the Description Section
	rationale = re.compile("Rationale:")							# Identifes the Rationale Section
	impact = re.compile("Impact:")									# Identifies the Impact Sectino
	audit = re.compile("Audit:")									# Identifies the Audit section
	remediation = re.compile("Remediation:")						# Identifies the Remediation section
	defaultValue = re.compile("Default Value:")						# Identifies the Default Value Section
	references = re.compile("References:")							# Identifes the References Section
	cisControls = re.compile("CIS Controls:")						# Identifes the CIS Controls Section
	acknowledgements = re.compile("Acknowledgements")				# Identifies the Acknowledgements to bypass the TOC and other garbage
	pages = re.compile("^(\"|)\d* \| Page ")						# Identifes the page and number format


	# Go through each line in the text one by one
	for line in text:
		#Fix some formatting in lines
		line = line.strip()
		#Removes bullet points
		line = re.sub(r"\u25CF*", r"", line)
		# Removes any leading spaces
		line = line.lstrip(' ')
		# Removes any leading double quotes
		line = line.lstrip('"')

		# Removes the page numbers from each page
		if re.search(pages, line):
			line = re.sub(r"(\"|)\d* \| Page ", "", line)
			line = line.lstrip(' ')

		# Used to track the number of "Ack" counts and omits title page and TOC when it reaches 2
		if re.search(acknowledgements, line):
			ack += 1
		
		# Checks to see if there are 4 or more . If there is, pass these as they are garbage (Could be omitted now using ACK variable)
		if re.search("[\.]{4,}.*\d*",line):
			continue

		# Writes the final contorl after the Summary Table has been found
		elif re.search("Summary Table", line) and ack == 2:
			write_file(control)
			break

		#Searches for the title and will only work if ACK is 2 and contentFlag is 0 indicating new control
		elif re.search(title, line) and ack == 2 and contentFlag == 0:
			#New Control writes the old control dictonary through Write_File(control)
			if nextControl == 1:
				write_file(control)
				#Clears the values of the control dictionary
				control = {key: "" for key in control}
				#Writes in title
				control['title'] = line
			#If next control is not 1, it will be the first time finding a control
			else:
				control['title'] = line
				index = 'title'
				nextControl = 1

			"""
			This elif statement will check each line of the file to see if it is in a new area or not

			If the line is not a title, then set the contentFlag, and check each situation using regex. If there is no match, change nextParagraphFlag to 1. Each new area found will set the new index to its name.
			"""
		elif not re.search(title, line) and ack == 2:
			contentFlag == 1
			if re.search(title2, line):
				index = 'title'
				nextParagraphFlag = 1
			elif re.search(profile, line):
				index = 'profile'

			elif re.search(description, line):
				index = 'description'

			elif re.search(rationale, line):
				index = 'rationale'

			elif re.search(impact, line):
				index = 'impact'
			
			elif re.search(audit, line):
				index = 'audit'
			
			elif re.search(remediation, line):
				index = 'remediation'
			
			elif re.search(defaultValue, line):
				index = 'defaultValue'
			
			elif re.search(references, line):
				index = 'references'

			elif re.search(cisControls, line):
				index = 'cisControls'
				contentFlag = 0

			else:
				nextParagraphFlag = 1

		"""
		Used to append to line if there was no identified new control/area

		If nextParagraphFlag is 1, that indicates that no new area has been found. Append to the most recently found control using "Index"
		"""
		if nextParagraphFlag == 1 and nextControl == 1 and ack == 2:
			line1 = control[index]
			line = line1 + " " + line
			control[index] = line
			nextParagraphFlag = 0



def pdfParser(sourceFile, outputFile):
	"""
	pdfParser Extracts the text of the provided sourceFile

	The PDF extraction provides the text of the file to the text variable. Additional functinoality to remove some of the features of PDFs.

	:param sourceFile: The provided source file
	:type sourceFile: string
	:param outputFile: the location to save the data
	:type outputFile: string

	:param partsOfPage: The list to store the portions of the page from the visitor_body function
	:type partsOfPage: list
	:param text: The saved data of the extracted PDF text.
	:type text: string
	"""
	partsOfPage = []
	text = ''


	def visitor_body(text, cm, tm, fontDict, fontSize):
		"""
		visitor_body Used to remove portions of the PDF (headers and footers)

		:param text: The text from the extracted text
		:type text: string
		:param cm: The length from the top header
		:type cm: int
		"""
		y = tm[5]
		if y > 50 and y < 720:
			partsOfPage.append(text)

	reader = PdfReader(sourceFile)
	totalPages = reader.getNumPages()
	for pageNumber in range(totalPages):
		page = reader.getPage(pageNumber)
		
		page.extract_text(visitor_text=visitor_body)
		text += ''.join(partsOfPage)
		text += "\n"
		partsOfPage = []
	
	#Temp file created to store the data	
	with open( "temp.txt", "w") as tempFile:
		tempFile.writelines(text)
	#Temp file read to read the data
	with open("temp.txt", "r") as tempFile:
		"""
		checkForControl Function to perform the regex work

		:param tempFile: The saved temporary file reader object
		:type tempFile: File Reader
		:param outputFile: The location to save the parsed PDF
		:type outputFile: string
		"""
		checkForControl(tempFile, outputFile)
	

def create_file (outputFile, flag):
	"""
	
	create_file Creates a CSV file

	Creates a CSV file with the dictionary fields as header files, saved in the outputFile location

	:param outputFile: The file to be saved to.
	:type outputFile: string
	:param flag: Flag to describe the saving function (CSV or XLSX) -=Currently only CSV option is available=-
	:type flag: string

	"""
	fields = {
		'title' : '',
		'profile' : '',
		'description' : '',
		'rationale' : '',
		'impact' : '',
		'audit' : '',
		'remediation' : '',
		'defaultValue' : '',
		'references' : '',
		'cisControls' : ''
	}

	if flag == 0:
		#write CSV
		with open (outputFile, 'w', newline = '') as csvFile:
			csvWriter = csv.DictWriter(csvFile, fieldnames=fields.keys())
			csvWriter.writeheader()

	# elif flag == 1:
	# 	#write XLSX
	# 	workbook = openpyxl.Workbook()

	# 	workbook.active.append(fields)
	# 	worksheet = workbook.active
	# 	worksheet.freeze_panes="A2"
	# 	worksheet.print_title_rows='1.1'
	# 	workbook.save(outputFile)


		

#Help menu
	"""
	Provide the usage menu for the program.
	"""
def help():
	print("Usage...")
	print("cis_parser.py -f <output file> -s <source file>")
	print("-f : Destination filepath and name (xlsx, csv)")
	print("-s : Source File location (PDF)")
	print("-h : Displays this help menu")
	print("-g : Opts for the use of the GUI file selection. This option requires no further arguments")
	print("-l : 0 for CSV or 1 for XLSX")

#get arguments
def parse_args(argv):
	"""
	Parses arguments provided by the users

	The commandline arguments provided by the user are parsed and asigned to variables to be used throughout the code.

	Parameters
	----------
	arg1 : list
		Contains the command line arguments

	Returns 
	-------
	string : sourceFile
		The specific pdf file (CIS Benchmark)
	string : outputFile
		The specific file to save to (csv)

	"""
	#Declare working vars
	sourceFile = '' 
	outputFile = ''
	flag = 0
	try:
		opts, args = getopt.getopt(argv, "hgdl::s:f:t")
	except:
		help()
	for opt, arg in opts:
		#help function
		if opt in ['-h']:
			help()
		#use command line over GUI
		elif opt in ['-g']:
			sourceFile = askopenfilename(title="Select the source file...")
			outputFile = asksaveasfile(title="Select the destination files...", filetypes=(("CSV Files", "*.csv"), ("XLSX Files", "*.xlsx")))

		#source file
		elif opt in ['-s']:
			sourceFile = arg
			if not os.path.exists(sourceFile):
				print("Please check your path")		

		elif opt in ['-f']:
				outputFile = arg				

		#used for testing
		elif opt in ['-t']:
			exit(1)

		elif opt in ['-l']:
			create_file(outputFile, flag)

	return sourceFile, outputFile

#main
if __name__ == '__main__':
	"""Main function gathering commandline arguments"""
	argv = sys.argv[1:]
	sourceFile, outputFile,  = parse_args(argv)
	pdfParser( sourceFile, outputFile)
import bs4
import os
import pickle
import re
import time
import unicodedata
import urllib.parse
import urllib.request
import xlrd
import xml.etree.ElementTree


# Needs to be bound to os.getcwd
path = r'C:\Users\User\Documents\'


class Parser():
	'''Define an object to parse the original data.
	
	This class will simply parse the original information into small 
	functions tha will be called later if needed.
	Since the information received throughout spreadsheets may vary, 
	each function will get only the necessary information, without the 
	need to modify the script every time we recieve a document with 
	different formating.
	'''
	
	def __init__(self, info, indice):
		'''Initiates the class with a list containing raw data 
		and the respective indices in the spreadsheet.
		
		Args:
			info: a list with all the information scrambled that needs 
				to be fetched.
			indice: a list with ordered indices that determine the correct 
				order of the information.   
		'''

		self.info = info
		self.indice = indice

	def name(self):	
		return self.info[self.indice[0]]
			
	def addresse(self):	
		return self.info[self.indice[1]] if self.indice[1] != -1 else '' 
		
	def address(self):
		'''Remove any comma in the string.'''
		data = self.info[self.indice[2]]

		# Need to check for the presence of the number column
		# The comma will be removed or not based on that value
		number = self.indice[3]
		formatted = data.replace(',', '') if number != -1 else data

		return formatted if self.indice[2] != -1 else '' 
		
	def number(self):
		'''If present, number will be formated to simulate a generic address.

		Later, the number will be passed as a string to a function that will 
		join and parse the address. It wouldn't be necessary if we already 
		have a separate number column, but as we may encounter some addresses 
		that don't have any number (even if the spreadsheet has a column for it), 
		then we need to certify a similar formating for all the addresses.

		Returns:
			A stirng with the number and an added comma.

		Exception:
			ValueError: if number isn't present, the negative value will 
				search for the last value in the column, and that value 
				may not even be a integer.
		'''

		# It's necessary to transform to an integer, otherwise
		# it may return a floating point number

		try:
			data = int(self.info[self.indice[3]])
		except ValueError:
			data = self.info[self.indice[3]]

		formatted = ', ' + str(data)
		return formatted if self.indice[3] != -1 else ''
	
	def complement(self):
		return self.info[self.indice[4]] if self.indice[4] != -1 else ''
	
	def city(self):
		return self.info[self.indice[5]] if self.indice[5] != -1 else ''
		
	def postal(self):
		return self.info[self.indice[6]] if self.indice[6] != -1 else ''


def generator(folder, message):
	'''Iterate through every row in the spreadsheet.
	
	Open the desired spreadsheet and iterate through every row and 
	column, generating a temporary list with all the cells in the 
	current row.
	
	Yields:
		A list with all the information contained in that row.
	'''

	directory = path + '\\Spreadsheets' + folder
	archive = load(directory, 'xls|xlsx', message)

	workbook = xlrd.open_workbook(archive[0])
	
	for sheet in workbook.sheets():
	
		# Specify the number of rows and columns 
		cols, rows = sheet.ncols, sheet.nrows
		
		for row in range(1,rows):
		
			# Generate the row content
			cell = [sheet.cell_value(
				rowx=row,
				colx=x
			) for x in range(cols)]
							
			yield cell, row, rows - 1, archive[1]


def index():
	'''Generate a list with indices.
	
	Generates a list with all the respective indices of the spreadsheet,
	or -1 if the given column doesn't hold any value.	
	It needs to ask the user about the order of the indices.
	
	Returns:
		A list with each index or -1 if not present. The list will have 
		the following order:
		
		[0] Name of the reciever
		[1] Name of the addresse if given
		[2] Street name or full adress
		[3] Street number if has it's own column
		[4] Complement if has it's own column
		[5] Name of the city
		[6] Zip code
	'''

	# Specify the order of the columns as needed
	contents = ('Receiver, Addresse, Address, '
			    'Number, Complement, City, '
			    'Zip Code')
	
	print('Enter the number of each one of the following columns.',
		  'If not present, type -1:\n' + contents)
	
	order = input()
	
	# Define a list with each index
	indexes = [int(x) for x in order.split(' ')]
	
	return indexes

	
def request(url, post):
	'''Fire a request to the given url and returns the whole page.'''
	req = urllib.request.urlopen(url, post).read()
	page = req.decode('utf-8', 'ignore')
	
	return page

## -------------------- Google Maps Geocode API --------------------

def google(address):
	'''Fetch and parse the information from Google Maps Geocode API.
	
	The data will be fetched from a simple request. The recieved info will be 
	bound to a variable (as a string) and will contain all the response in 
	XML format.
	
	Args:
		address: the raw address from the spreadsheet.
		
	Raises:
		AttributeError: while looping through the results of the xml 
			response, if the children doesn't have a 'type' tag, 
			it will raise an exception and try to search for the 
			'location' tag.
		TypeError: after fail in the search for a 'type' tag, it will attempt 
			to search for a 'location' tag. And if still don't find anything, 
			will raise a TypeError beacause one of the tags that match 
			doesn't have any children element (<location_type>).
			
	Returns:
		A list containing all the information fetched by Google. The list 
		will have a determined order as follow:
		
		info = [
			street_number, route,
			administrative_area_level_2, administrative_area_level_1,
			lat, lng
		]
		
		If the response from Google it's not positive, then will return 
		the response message itself.
	'''
	
	# Series of functions that will be matched and fetched as pairs
	# for each keyword
	# Because both short and long name are needed, 
	# we have to specify in wich keyword we need them
	
	def findStreet(element):
		return re.search('route', element.text)
	
	def getStreet(element):
		name = element.find('long_name')
		street = name.text
		return street
		
	def findNumber(element):
		return re.search('number', element.text)
		
	def getNumber(element):
		name = element.find('long_name')
		number = name.text
		return number
	
	def findCity(element):
		return re.search('administrative_area_level_2', element.text)
		
	def getCity(element):
		name = element.find('long_name')
		city = name.text
		return city
		
	def findState(element):
		return re.search('administrative_area_level_1', element.text)
	
	def getState(element):
		name = element.find('short_name')
		state = name.text
		return state
	
	# Pack of functions that will match and then return a specific 
	# element
	functions = [
		(findStreet, getStreet),
		(findNumber, getNumber),
		(findCity, getCity),
		(findState, getState)
	]

	link = 'https://maps.googleapis.com/maps/api/geocode/xml?address='
	
	# Need to encode the address to be a valid url
	encode = urllib.parse.quote_plus(address)
	url = link + encode
	
	# The XML content as a string
	req = request(url, None)	
	
	# Generate log files
	directory = path + '\\Logs\\' + r'{}.xml'.format(address)
	with open(directory, 'w') as f:
		f.write(req)
	
	root = xml.etree.ElementTree.fromstring(req)
	
	status = root.find('status')
	response = status.text
	
	# Look for the status of the response
	# If negative, return the response message
	if response == 'OK':

		data = []

		for result in root.findall('result'):

			info = []

			for children in result:
		
				try:
					type = children.find('type')							
					# Search for the keyword and return it's value if found 
					for pattern, match in functions:
						if pattern(type):					
							info.append(match(children))
									
				# Search for location	
				except AttributeError:
			
					try:
						location = children.find('location')
						# Find lat and lng (the only children)				
						for coordinates in location:
							info.append(coordinates.text)
										
					except TypeError:
						continue			
		
			data.append(info)

		return data, response
	
	else:
		return response	
		
## -------------------- Correios parser --------------------

def correios(data):
	'''Send information to Correios web page and return a valid zip code.
	
	The information returned by Google will be used as the data to be 
	sent in this request. Because the info is already formated, there's no 
	need to edit or parse the data, only alocate the correct indexes of 
	the list.
	
	Args:
		data: data will be a list with the needed information to perform 
			a post request to the Correios web page.
			
	Returns:
		If more than one zip code is found, it will return a list containing 
		all possible matches. The user is the one that has to decide wich 
		of the found numbers is the correct one. If only one is found, then 
		it returns the same list, but with only one item long. The list will 
		contain information as follow:
		
		info = [
			street, neighborhood,
			city, zip code 
		]
		
		Also, if it's a bad address, may not give any usefull information, 
		in wich case it will return the message found on the web page. 
	'''
	
	link = ('http://www.buscacep.correios.com.br/'
			'sistemas/buscacep/'
			'resultadoBuscaCep.cfm')

	post = {
		'UF': data[3],
		'Localidade': accent(data[2]),
		'Logradouro': accent(data[1]),
		'Numero': data[0],
		'Submit': 'Buscar'
	}
	
	# Need to encode the content of the request
	encode = urllib.parse.urlencode(post).encode()
	
	req = request(link, encode)
	
	# Bind the HTML content to a BeautifulSoup element
	html = bs4.BeautifulSoup(req, 'html.parser')
	
	# Search for the positive response
	find = html.find_all('p')
	
	# The last p tag is the response
	response = find[-1].text
	
	if not re.search('NAO', response):

		# Find the table wich contains all the information
		table = html.find_all('table')
		
		rows = table[0].find_all('tr')
		info = []
		
		# If the table has more than two rows, it means that will 
		# return more than one zip code
		if len(rows) > 2:
			
			for row in rows[1:]:		
				# Generate a list for each of the given zip codes
				zip_code = [x.text for x in row.find_all('td')]
				info.append(zip_code)
		
		# If there's only one zip code, return all the elements of the row
		else:
			zip_code = [x.text for x in rows[1].find_all('td')]
			info.append(zip_code)
		
		return info
	
	# If it's a negative response, returne the message	
	else:
		return response
	
## -------------------- address analisys --------------------

def addressParser(data):
	'''Analyze and parse the address, returning street name, number 
	and complement if present and as separate fields.

	Because some of the addresses have different formating, there is a 
	necessity to analyze each one and return separate values for street 
	name,street number and complement.

	Args:
		data: the original address. May be a string or a list with street 
			name, street number and complement if they are in separate fields.

	Return:
		If doesn't have a number column and can't find more than 
		one item if splitted, will then search for a number and, 
		if still can't find any number, will return false. It will 
		return the address itself if can't split the address between 
		a comma, but only if a number is found. The full form is a list 
		containig the street address, the number and complement, all as 
		separate elements that can be used later.
	'''
	
	sequence = [str(x) for x in data]
	
	pattern = [
		r'\d{1,4}',
		r'(\d{1,4})',
		r'^\s*'
	]
	
	string = ' '.join(sequence).replace(',', '')
	
	# Find a number
	address = string if re.search(pattern[0], string) else False
	
	if address != False:
		street, complement = re.split(
			pattern[0],
			address,
			maxsplit=1
		)
		# Replace any whitespace in the beginning
		complement = re.sub(pattern[2], '', complement)
		# Validate the complement
		complement = False if not complement else complement
		
		number = re.search(pattern[1], address).groups(0)[0]

		return street, number, complement
		
	else:
		return address

## -------------------- ascii converter --------------------	

def accent(string):
	'''Remove any special characters of a string.

	Strings containing special characters can't be passed into a 
	post request and sometimes can't be printed in the terminal.

	Args:
		string: a string that may or may not have any special 
		characters on it.

	Returns:
		word: the same string containing only ASCII characters.
	'''

	word = ''.join(x for x in unicodedata.normalize('NFD', string) if
		unicodedata.category(x) != 'Mn'
	)
	return word
	
## -------------------- web crawler --------------------

def crawler():

	# Generate a list of indexes
	positions = index()
	
	# All the information will be stored here
	data, name = [], None

	message = 'Choose the original mailling:'
	directory = '\\Mailling'
	for row in generator(directory, message):

		name = row[3]
	
		try:
			print('\n{}/{}\n{}'.format(
				row[1],		# Current address
				row[2],		# Total addresses 
				row[0]		# Data
			))
		except UnicodeEncodeError:
			print('\n{}/{}\n{}'.format(
				row[1],		# Current address
				row[2],		# Total addresses 
				'Can\'t decode special character.'		# Data
			))
			
		# Any information regarding the original data will be fetched 
		# by calling this object.	
		delivery = Parser(row[0], positions)
		
		raw_address = [
			delivery.address(),
			delivery.number(),
			delivery.complement()
		]

		formatted_address = addressParser(raw_address)

		# Check if it's a valid address before sending to Google		
		if formatted_address == False:
			# This block of code will execute if it's a bad address
			message = 'Address has no complement.'
			print(message)
			row[0].append(message)
			data.append(row[0])			
			continue
			
		else:
			# Send a formatted address to Google
			address_url = '{}, {} - {}'.format(
				formatted_address[0],	# Street
				formatted_address[1],	# Number
				delivery.city()			# City (from the document)
			)
		
		google_data = google(address_url)

		if type(google_data) == str:
			print(google_data)
			row[0].append(google_data)
			data.append(row[0])
			continue

		print('Response:', google_data[1])
		if len(google_data[0]) > 1:
			for i, address in enumerate(google_data[0]):
				print('{}. {} - {} / {}'.format(
					i,
					address[1],
					address[2],
					address[3]
				))
			msg = 'Google found more than one address. Choose one: '
			user_address = input(msg)
			google_data = google_data[0][int(user_address)]
		else:
			google_data = google_data[0][0]

		if len(google_data) < 6:
			google_data.insert(
				0,
				formatted_address[1]
			)
		
		# Search for the postal code in the Correios web page
		try:
			zip_code = correios(google_data)
		except:
			continue

		# If the Zip Code insn't a list, it menas it's a bad address
		if type(zip_code) == str:
			print(zip_code, '\n')
			row[0].append(zip_code)		
			data.append(row[0])
			continue

		# If the address return more than one ZIP Code, ask the user 
		# to choose the correct one 
		elif len(zip_code) > 1:

			for i, number in enumerate(zip_code):
				try:
					print('{}. {} - {} - {} {}'.format(
						i,
						number[0],		# Street
						number[1],		# Neighborhood
						number[2],		# City
						number[3]		# ZIP Code
					))
				except UnicodeEncodeError:
					pass
				
			pick = input('Choose a Zip Code: ')
			postal_code = zip_code[int(pick)][3]
		
		else:
			postal_code = zip_code[0][3]

		complement = formatted_address[2]
		# Format the final address after check if has a complement
		number = '{} / {}'.format(
			google_data[0],
			complement
		)	if complement != False else google_data[0]
		
		# street, number / complement(optional) - city
		final_address = (google_data[1] + ', ' +
			number + ' - ' +
			google_data[2] + '/' +
			google_data[3]
		)

		print('Postal Code: {}'.format(postal_code))

		# All the formatted info that was retrieved
		info = [
			delivery.name(),
			final_address,
			postal_code,
			delivery.postal(),
			google_data[2],		# City
			google_data[3],		# State
			google_data[4],		# Lat
			google_data[5],		# Lng
		]
		
		# Check and add the addresse to the final list if it's present 
		addresse = delivery.addresse()
		
		if positions[1] != -1:
			info.insert(1, addresse)
		else:
			pass
		
		data.append(info)
		
		time.sleep(2)

	save(data, '\\GeoInfo', name)

## -------------------- save binary file --------------------

def save(data, folder, name):
	'''Pickle the passed data to a file.

	To increase the general workflow of the script, 
	a binary file will hold the information if needed.'''
	
	directory = path + '\\Database' + folder
	db = r'{}\{}.pickle'.format(directory, name)
	
	with open(db, 'wb') as f:
		pickle.dump(data, f)
		
	print('Data succefully generated! Saved at:', db)
	
## -------------------- open binary file --------------------

def launch(path):
	
	db = r'{}'.format(path)
	
	with open(db, 'rb') as f:
		data = pickle.load(f)

	return data

## -------------------- print directory --------------------

def load(path, extension, message):

	pattern = r'\.{}$'.format(extension)

	listing = [x for x in os.listdir(path) if re.search(pattern, x)]
	
	print('\n{}'.format(message))
	
	for i, item in enumerate(listing):
		print('{}. {}'.format(i, item))
	
	pick = input()

	name = listing[int(pick)]

	split = r'\.\w+$'

	return r'{}\{}'.format(path, name), re.split(split, name)[0]
	
## -------------------- sort zip-codes --------------------

def sort():
	'''Simply puts the data into a spreadsheet.'''
	#workbook = xlwt.workbook(encoding='utf-8')
	#sheet = workbook.add_sheet

	def comparison(number, code):
		return number[:len(code)] == code

	data = load()

	names = [
		'CDA Cristal',
		'CDA Farrapos',
		'CDA Farroupilha',
		'CDA Iguatemi',
		'CDA Jardim Botânico',
		'CDA Menino Deus',
		'CDA Teresópolis',
		'CDB Protásio Alves'
	]
	codes = [
		[908, 919],
		[902, 910, 911, 912],
		[9003, 9004],
		[913, 914, 915],
		[9001, 9002, 906],
		[901],
		[917],
		[904, 905]
	]

	agents = dict(zip(names, codes))
	info = []

	for item in data:

		# Add bad addresses to the begining
		if item[-1] == False:
			info.insert(0, item)
			continue

		item_zip = item[-6]

		for agent in agents:

			# Test for equalty
			test = [comparison(
				str(item_zip),
				str(code)
			) for code in agents[agent]]
		
			if True in test:
				print('{} => {}'.format(
					agents[agent], item[-6]
				))
			else:
				pass


def compare(item, base):
		return item == base
		
## -------------------- match labels-addresses --------------------
		
def match():

	message = [
		'Choose the database:',
		'Choose the spreadsheet with the labels:'
	]
	db = path + '\\Database\\GeoInfo'
	archive = load(db, 'pickle', message[0])
	data = launch(archive[0])

	labels = generator('\\Labels', message[1])

	for x in labels:
		#data[0]
		#labels[0][2]
		for i, y in enumerate(data):
			check = compare(str(y[0]), x[0][2])
			data[i] = [int(x[0][1])] + y if check else data[i]
	
	save(data, '\\Labels', archive[1])

	#return data
	
## -------------------- match delivery status --------------------
	
def events():

	message = [
		'Choose the database:',
		'Choose the spreadsheet with the delivery history:'
	]
	db = path + '\\Database\\Labels'
	archive = load(db, 'pickle', message[0])
	
	data = launch(archive[0])

	status = [x for x in generator('\\History', message[1])]
	
	for x in status:
		for i, y in enumerate(data):
			data[i] = y + [x[0][8]] if x[0][0] == str(y[0]) else data[i]
			
	save(data, '\\History', archive[1])
	
## -------------------- html/js generator --------------------

def html():

	message = 'Choose the database:'
	directory = path + '\\Database\\History'
	archive = load(directory, 'pickle', message)
	
	db = launch(archive[0])
	
	web = '\\Markers\\Generated' + '\\{}.html'.format(archive[1])

	for i, item in enumerate(db):
	
		source = path + '\\Markers\\source.html' if i == 0 else path + web
		 
		f =  open(source, 'r', encoding='utf-8')
		data = f.readlines()
		f.close()
	
		pattern = r'var locations'

		index = [data.index(x) for x in data if re.search(pattern, x)]
		
		info = ['    ', item, ',\n']
		
		new = data[:index[0]+1] + info + data[index[0]+1:]
		
		markers = open(path + web, 'w', encoding='utf-8')
		
		for x in new:
			markers.write(str(x))
		markers.close()
		
	print('File succefully generated!')

def htmlFromGeoInfo():

	message = 'Choose the database:'
	directory = path + '\\Database\\GeoInfo'
	archive = load(directory, 'pickle', message)
	
	db = launch(archive[0])
	
	web = '\\Markers\\Generated' + '\\{}.html'.format(archive[1])

	for i, item in enumerate(db):
	
		source = path + '\\Markers\\source.html' if i == 0 else path + web
		 
		f =  open(source, 'r', encoding='utf-8')
		data = f.readlines()
		f.close()
	
		pattern = r'var locations'

		index = [data.index(x) for x in data if re.search(pattern, x)]
		
		info = ['    ', item, ',\n']
		
		new = data[:index[0]+1] + info + data[index[0]+1:]
		
		markers = open(path + web, 'w', encoding='utf-8')
		
		for x in new:
			markers.write(str(x))
		markers.close()
		
	print('File succefully generated!')


def main():
	'''Manage all the process.'''


	def option(data):
		'''Enumerate and print all the options give.'''
		for i, item in enumerate(data):
			line = '{}. {}'.format(i, item)

			print(line)


	functions = [
		crawler,
		'Find ZIP Code and Geodata',
		match,
		'Bind labels to addresses',
		events,
		'Generate the status of each delivery',
		html,
		'Generate markers in Google Maps',
		htmlFromGeoInfo,
		'Generate markers without history'		
	]
	print('\nChoose an action to perform:')

	option(functions[1::2])

	action = input()
	launch = functions[::2][int(action)]()

	main()


if __name__ == '__main__':

	main()
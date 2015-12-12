#Programmer:
#Date:
#Description: Primary access class for OSS Code Repository
#Filename: oss_access.py

import xlrd
import sets

class OSS:
	groupArray 		= []
	titleArray 		= []
	codeArray 		= []
	filenameArray 	= []
	
	CurrentModeData = (groupArray, titleArray, codeArray, filenameArray)
	book = xlrd.open_workbook("OSS Code Repo.xls")
	sheetCount = len(book.sheet_names())
	
	def loadMode(self):
		sheet = self.book.sheet_by_name('OSS')
		return sheet
	
	
	# Populate data from the resource
	def populateData(self):
		sheet = self.loadMode()
		print "Populating OSS Repository data..."
		groupIndex = sheet.col_values(0)
		titleIndex = sheet.col_values(1)
		codeIndex  = sheet.col_values(2)
		filenameIndex = sheet.col_values(3)
		self.CurrentModeData = (self.groupArray, self.titleArray, self.codeArray, self.filenameArray)
		print "Groups..."
		for item in groupIndex:
			self.CurrentModeData[0].append(item)
		print "Titles..."
		for item in titleIndex:
			self.CurrentModeData[1].append(item)
		print "Code..."
		for item in codeIndex:
			self.CurrentModeData[2].append(item)
		print "Filenames..."
		for item in filenameIndex:
			self.CurrentModeData[3].append(item)
		print "Done."
			
	def listGroups(self):
		groupItems = self.CurrentModeData[0]
		groupItems.sort()
		#for item in groupItems:
			#print item
		#print "Count: ", len(groupItems)
		return groupItems

	def listTitles(self):
		titleItems = self.CurrentModeData[1]
		titleItems.sort()
		#for item in titleItems:
			#print item	
		#print "Count: ", len(titleItems)
		return titleItems

	def listCode(self):
		codeItems = self.CurrentModeData[2]
		#for item in codeItems:
			#print item
		#print "Count: ", len(codeItems)
		return codeItems
		
	def listFilenames(self):
		filenameItems = self.CurrentModeData[3]
		#for item in filenameItems:
			#print item
		#print "Count: ", len(filenameItems)
		return filenameItems
		
	def search(self, query):
		masterResults = []
		searchResults = []
		self.populateData()
		
		groupIndexSearchResults = []
		groupIndexSearchResults = [ t for t in self.listGroups() if query in t ]
		groupIndexSearchResults = sets.Set(groupIndexSearchResults)
		#groupIndexSearchResults = sets.Set(groupIndexSearchResults)
		for group in groupIndexSearchResults:
			group_ref_number = self.CurrentModeData[0].index(group)
			searchResults.append([group, 
								self.CurrentModeData[1][group_ref_number], 
								#self.CurrentModeData[2][group_ref_number], 
								self.CurrentModeData[3][group_ref_number]])
		
		titleIndexSearchResults = []
		titleIndexSearchResults = [ t for t in self.listTitles() if query in t ]
		titleIndexSearchResults = sets.Set(titleIndexSearchResults)
		for title in titleIndexSearchResults:
			title_ref_number = self.CurrentModeData[1].index(title)
			searchResults.append([self.CurrentModeData[0][title_ref_number],
								title,
								#self.CurrentModeData[2][title_ref_number],
								self.CurrentModeData[3][title_ref_number]])
								  
		
		filenameIndexSearchResults = []
		filenameIndexSearchResults = [ t for t in self.listFilenames() if query in t ]
		filenameIndexSearchResults = sets.Set(filenameIndexSearchResults)
		for filename in filenameIndexSearchResults:
			filename_ref_number = self.CurrentModeData[3].index(filename)
			searchResults.append([self.CurrentModeData[0][filename_ref_number],
								self.CurrentModeData[1][filename_ref_number],
								filename])
								#self.CurrentModeData[2][filename_ref_number],
		
		print "Found", len(searchResults), "entries."
		return searchResults
		
	def listIndexedCount(self):
		pass
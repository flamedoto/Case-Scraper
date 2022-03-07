import time
from tqdm import tqdm
import sys
import random
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import *
import re
import os
import urllib.request
import requests
import pandas as pd
import math
from geopy.geocoders import Nominatim



class PublicCase():
	#Case Types that has to be scraped
	casetype = ['possession','eviction']
	#US Proxy IP PORT
	PROXY = "45.82.245.34:3128"
	geolocator = Nominatim(user_agent="geo")

	#Main URl
	#URL = 'https://public.courts.in.gov/mycase/#/vw/CaseSummary/eyJ2Ijp7IkNhc2VUb2tlbiI6IkRhc0ZabFVBYUZBSExmd1RsY28tZ0ZwemFVMkRuREZWOXlzeG5qUGotZVkxIn19'
	#URL = 'https://www.google.com/webhp?hl=en&sa=X&ved=0ahUKEwiI7fzAgensAhWIXsAKHXCoAjQQPAgI'

	## Defining options for chrome browser
	options = webdriver.ChromeOptions()
	#ssl certificate error ignore
	options.add_argument("--ignore-certificate-errors")
	#Adding proxy
	options.add_argument('--proxy-server=%s' % PROXY)
	Browser = webdriver.Chrome(executable_path = "chromedriver",options = options)

	#Excel file declaration
	ExcelFile = pd.ExcelWriter('data.xlsx')

	#Global variable deceleration
	TotalCase = 0
	TotalCaseDone = 0
	#Total rows in excel file
	Rows = 0
	LastCaseID = "" 
####
	

	def ExcelColorGray(self,s):
		#reutrn excel row with gray of total len 24
		return ['background-color: gray']*24
	
	def ExcelColor(self,s):
		#reutrn excel row with yellow of total len 24
		return ['background-color: yellow']*24

	def addressfilter(self,addr):
#		addr = """C/O Daniel L. Russello
#McNevin & McInness, LLP
#5442 S. East Street, Suite C-14
#Indianapolis, IN 46227"""

		#Spliting address by  new line so we can seperate all the variables
		addr = addr.split("\n")
		address = ""
		#print(addr)

		#Splitting string by new line to seperate address mailing name statezipcity
		addr1 = addr[-1].split(',')
		#First index will City
		city = addr1[0]
		#Split last index by space in which last index will be zip code and will index wil lbe state
		zipcode = addr1[-1].lstrip().split(' ')[-1]
		state = addr1[-1].lstrip().split(' ')[0]

		#removing city state zip line from the array
		addr.pop(len(addr)-1)


		#iterating array to get address
		for i in range(len(addr)):
			#Geolocator geo code will return complete address if the address provided is correct that way we can find that the index of array is address of mailing name
			location = self.geolocator.geocode(addr[i])
			#if provided value is not address it will raise an error if it does not we will store that address in adddress variable remove address index from array and break the loop
			try:
				demovar = location.address
				address = addr[i]
				addr.pop(i)
				#print(address)
				break
			except Exception as e:
				#print(str(e))
				pass
		#all the remaining indexes will mailing name
		mailingname = "".join(addr)
		
		return mailingname,address,city,state,zipcode




	def getinput(self):
		excel_data_df = pd.read_excel('input.xlsx',header=None)
		i = 0
		casenumbers = []
		for data in excel_data_df.values:
			#if its first iteration skip it, because its the header
			if i == 0:
				i += 1
				continue
			#Appending case number found in excel file to array
			casenumbers.append(data[0])
			i += 1
		#log
		print("Total Input Search queries found : ",len(casenumbers))


		#self.TotalCase = len(casenumbers)
		self.TotalCaseDone = 0

		#return all the case number found in excel file
		return casenumbers



	def searchcase(self):

		#calling get input function, function will Extract all inputs from Input excel file
		casenumbers = self.getinput()
		#search query url
		ur = 'https://public.courts.in.gov/mycase/#/vw/Search'
		caselen = 0
		for case in casenumbers:
			print("Searching for Case Number : ",case)
			self.Browser.get(ur)

			#Find the input text file of case number in the form
			casefield = self.Browser.find_element_by_xpath("//input[@id='SearchCaseNumber']")
			#Entering case number in the text field
			casefield.send_keys(case)

			time.sleep(1)
			#Find submit button
			submitbutton = self.Browser.find_element_by_xpath("//button[@class='btn btn-default']")

			#Submit the search query
			submitbutton.click()
			time.sleep(5)

			#search result function will calculate total result and iterate over all the found pages
			self.searchresults()


			caselen += 1
			#log
			print("Case queries done "+str(caselen)+" out of ",len(casenumbers))


	def searchresults(self):
		#Find total result found text i.e '1 to 20 of 577'
		totalresult = self.Browser.find_element_by_xpath("//span[@data-bind='html: dpager.Showing']").text
		#extract all numbers from '1 to 20 of 577' using regex
		totalresult = re.findall(r'\d+', totalresult)

		self.TotalCase = int(max(totalresult))
		#log
		print("Total Search result found : ",max(totalresult))
		#dividing the max number from regex output by total result per page
		totalresult = int(math.ceil(int(max(totalresult)) / 20))
		print("Total Pages :",totalresult)

		#loop till total pages
		for tot in range(totalresult):
			#Finding search result per page
			results = self.Browser.find_elements_by_xpath("//a[@class='result-title']")

			#Calling function will take parameter of all search results , This function will click on each search result one by one and scrape data from it
			self.searchresultiterate(results)

			#Find and click on next page button
			nextbutton = self.Browser.find_element_by_xpath("//button[@title='Go to next result page']").click()
			time.sleep(3)
			#log
			print("Pages Done "+str(tot+1)+" Out of ",totalresult)


	def searchresultiterate(self,results):
		#Iterating over all search result per page
		for i in range(len(results)):
			#Click on each search result if stale element exception find search result from page again 
			try:
				results[i].click()
			except StaleElementReferenceException:
				results = self.Browser.find_elements_by_xpath("//a[@class='result-title']")
				results[i].click()

			#calling data extraction function, this function will extract all required data from the Case Page
			self.DataExtraction()

			#Previous page through js code
			self.Browser.execute_script("window.history.go(-1)")
			time.sleep(2)
			#log
			self.TotalCaseDone += 1
			print("Result(s) Scraped "+str(i+1)+" Out of "+str(len(results))+" Total Cases Scraped : "+str(self.TotalCaseDone)+" / ",self.TotalCase)

	def DataExtraction(self):
		#self.Browser.get(self.URL)
		time.sleep(4)

		#Finding first table in which Case Number is present (Case Detail Table)
		casetypevar = self.Browser.find_elements_by_xpath('//div[@class="col-xs-12 col-sm-8 col-md-6"]//table//tr')

		#Finding Third table in which have to search that case is either eviction or possession or niether
		eptable = self.Browser.find_element_by_xpath("//table[@class='event-list table table-condensed table-hover']").text

		#calling function posevicheck (possession eviction check), this function will check if eviction or possession text is being used in this table
		proceed = self.poseviccheck(eptable)

		#If Eviction or Possession is present in the table then proceed is True
		if proceed == True:
			#Finding All Parties dropdowns
			partydetail = self.Browser.find_elements_by_xpath("//table[@class='ccs-parties table table-condensed table-hover']//span[@class='small glyphicon glyphicon-collapse-down']")
			#totlen variable is used for how many drop is being clicked
			totlen = 0
			uc = []
			#iterating over all the dropdowns found
			for pd in partydetail:
				#Click each and every one the of them if error means not clickable then skip it
				try:
					pd.click()
					totlen += 1
				except:
					totlen += 1
					#uc = index of unclickable divs
					uc.append(totlen)
					pass
			#Finding Table of parties
			pct = self.Browser.find_elements_by_xpath("//table[@class='ccs-parties table table-condensed table-hover']//tr")
			#Calling Fucntion partiescase takes parameter, Party Table,Case Detail Table,total len multiply by 2
			self.partiescase(pct,casetypevar,totlen*2,uc)
		#else its False Which means case is not a eviction case
		else:
			#Calling function Case details takes parameter , case detail table, This function will scrape all the required details from table i.e Case Number and will return 6 variables
			casenumber,court,type1,filed,status,statusdate = self.casedetails(casetypevar)
			#Defining 14 variable for excel file
			#tenetname,mailingname,address,city,state,zipcode,attorneyname,aa,propertyowner,mailingnameplain,mailingaddress,mailingcity,mailingstate,mailingzip= "","","Not an Eviction case","","","","","","","","","","",""
			#Calling excel write function will take 20 parameters of all excel col required by user, this function will write data in excel and save it
			self.ExcelWrite(casenumber,"","","","","","Not an Eviction Case","","","","","","","","","","","","","","","","","","Not an Eviction Case")




	def partiescase(self,pct,casetypevar,totlen,uc):
		#Calling function Case details takes parameter , case detail table, This function will scrape all the required details from table i.e Case Number and will return 6 variables
		casenumber,court,type1,filed,status,statusdate = self.casedetails(casetypevar)
		#This variable will use to skip one iteration after other
		skip = False
		#Variable for attorney counts
		countatt = 0
		#Variable for address counts
		count = 0
		#variable for attornet address counts
		countattadd = 0
		itercount = 0
		for i in range(totlen):
			#Defining 14 variable for excel file
			tenetname,mailingname,address,city,state,zipcode,attorneyname,aa,propertyowner,mailingnameplain,mailingaddress,mailingcity,mailingstate,mailingzip,attorneymailingname,attorneyzipcode,attorneycity,attorneystate= "","","","","","","","","","","","","","","","","",""
			#if Skip is true which means loop this skip last iteration
			if skip == True:
				#skip False and skip iteration
				skip = False
				continue
			#if skip is False
			else:
				skip = True
			itercount += 1
			#if current index is the index of unclickable div then skip the iteration
			if itercount in uc:
				continue

			#if defending is present in the table row as the text
			if 'Defendant' in pct[i].text:
				#Remove defendant from the table row remaining text will be tenet name
				tenetname = pct[i].text.replace('Defendant','').lstrip()
				#if Address is present in the next row to the defendant
				if 'Address' in pct[i+1].text:
					#Find address span tag as address raw text, state zip city mailing name will be in that text too
					addr = pct[i+1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[count].text
					#Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
					mailingname,address,city,state,zipcode = self.addressfilter(addr)
				else:
					#if address is not present then decrease 1 from address count variable
					count -= 1
				#If defendant has attorney
				if 'Attorney' in pct[i+1].text:
					try:
						#Find all the attorney in the parties table take of text of current index attorney
						attorneyname = pct[i+1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
						#if attorney is not pro see
						if 'Pro Se' not in attorneyname:
						#Find all the attorney addresses in the parties table take of text of current index attorney
							attorneyad = pct[i+1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")[countattadd].text
							
							#Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
							attorneymailingname,aa,attorneycity,attorneystate,attorneyzipcode=self.addressfilter(attorneyad)
						else:
							#if it is doesnt have the address

							#else decrease one from attorney address counts varaible
							countattadd -= 1

					except NoSuchElementException:
						pass
				else:
				#else decrease one from attorney counts variable and attorney address counts varaible
					countatt -= 1
					countattadd -= 1
			#if Plaintiff is present in the table row as the text
			elif 'Plaintiff' in pct[i].text:
				#Remove Plaintiff from the table row remaining text will be property owner
				propertyowner = pct[i].text.replace('Plaintiff','').lstrip()


				#if Address is present in the next row to the Plaintiff
				if 'Address' in pct[i+1].text:					#Find address span tag as address raw text, state zip city mailing name will be in that text too
					addr = pct[i+1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAddr']")[count].text
					
					#Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
					mailingnameplain,mailingaddress,mailingcity,mailingstate,mailingzip=self.addressfilter(addr)

				else:
					#if address is not present then decrease 1 from address count variable
					count -= 1

				#If Plaintiff has attorney
				if 'Attorney' in pct[i+1].text:
					try:
						#Find all the attorney in the parties table take of text of current index attorney
						attorneyname = pct[i+1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAtty']")[countatt].text
						#if attorney is not pro see
						if 'Pro Se' not in attorneyname:
						#Find all the attorney addresses in the parties table take of text of current index attorney
							attorneyad = pct[i+1].find_elements_by_xpath("//span[@aria-labelledby='labelPartyAttyAddr']")
							#Split the address by new line
							attorneyad = attorneyad[countattadd].text
							#Calling Function address filter takes parameter raw address, this function will seperate mailing name address city name zipcode state from raw address and return it
							attorneymailingname,aa,attorneycity,attorneystate,attorneyzipcode=self.addressfilter(attorneyad)
						else:
							#if it is doesnt have the address

							#else decrease one from attorney address counts varaible
							countattadd -= 1

					except NoSuchElementException:
						pass
				else:
				#else decrease one from attorney counts variable and attorney address counts varaible
					countatt -= 1
					countattadd -= 1
			else:
				#adding 1 to each variable, address count, attorney count, attorney address acount
				count += 1
				countatt += 1
				countattadd += 1
				continue


			#Calling excel write function will take 20 parameters of all excel col required by user, this function will write data in excel and save it
			self.ExcelWrite(casenumber,court,type1,filed,status,statusdate,tenetname,mailingname,address,city,state,zipcode,aa,attorneyname,propertyowner,mailingnameplain,mailingaddress,mailingcity,mailingstate,mailingzip,attorneymailingname,attorneyzipcode,attorneycity,attorneystate,"")
			#adding 1 to each variable, address count, attorney count, attorney address acount
			count += 1
			countatt += 1
			countattadd += 1





	def poseviccheck(self,eptable):
		#default proceed True means eviction or possession found in table
		proceed = True

		#Iterate through array which was define in the start 
		for c in self.casetype:

			if c in eptable.lower():
				proceed = True
			else:
				proceed = False

		return proceed


	def casedetails(self,casetypevar):
		#required Variables
		casenumber = ''
		court = ''
		type1 = ''
		filed = ''
		status =''
		statusdate = ''

		#iterating table rows (tr) of table 
		for cases in casetypevar:
			#if case number is present in it remove case number text from the string and add it to variable
			if 'case number' in cases.text.lower():
				casenumber = cases.text.replace(' ','').strip('CaseNumber')
			#if court is present in it remove court text from the string and add it to variable
			elif 'court' in cases.text.lower():
				court = cases.text.strip('Court').lstrip()
			#if type is present in it remove type text from the string and add it to variable
			elif 'type' in cases.text.lower():
				type1 =cases.text.replace('Type','')
			#if filed is present in it remove filed text from the string and add it to variable
			elif 'filed' in cases.text.lower():
				filed =cases.text.replace('Filed','')
			#if status is present in it
			elif 'status' in cases.text.lower():
				#Split status by comma(,) last index will be status and first will be status date always
				t = cases.text.replace('Status','').split(',')
				status = t[-1]
				statusdate = t[0]


		#returning all the required varialbes
		return casenumber.strip(),court.strip(),type1.strip(),filed.strip(),status.strip(),statusdate.strip()






	def ExcelWrite(self,casenumber,court,type1,filed,status,statusdate,tenetname,mailingname,address,city,state,zipcode,aa,attorneyname,propertyowner,mailingnameplain,mailingaddress,mailingcity,mailingstate,mailingzip,attorneymailingname,attorneyzipcode,attorneycity,attorneystate,eviction):
		df = pd.DataFrame({"Case Number": [casenumber],"Status": [status],"Township": [court],"Type": [type1],"Filed Date": [filed],"Status Date": [statusdate],"Tenant Name": [propertyowner],"Mailing Name": [mailingnameplain],"Mailing Address": [mailingaddress],"Mailing City": [mailingcity],"Mailing State": [mailingstate],"Mailing Zip": [mailingzip],"Property Owner": [tenetname],"Property Mailing Name": [mailingname],"Property Address": [address],"Property City": [city],"Property State": [state] ,"Property Zip": [zipcode] ,"Attorney Name": [attorneyname],"Attorney Address": [aa],"Attorney Mailing Name": [attorneymailingname], "Attorney State": [attorneystate],"Attorney City": [attorneycity],"Attorney Zip": [attorneyzipcode]})
		#if case is not an eviction case
		if eviction == "Not an Eviction Case":
			#add yellow color to the row
			df = df.style.apply(self.ExcelColor, axis=1)
		#If first entry in excel
		if self.Rows == 0:
			df.to_excel(self.ExcelFile,index=False,sheet_name='Data')
			self.Rows = self.ExcelFile.sheets['Data'].max_row
			self.LastCaseID = casenumber
		else:
			#if this is the new case add a new line to excel before adding case data to excel
			if self.LastCaseID != casenumber:
				#creating empty dataframe of element len 24
				df1 = pd.DataFrame({"Case Number": [""],"Status": [""],"Township": [""],"Type": [""],"Filed Date": [""],"Status Date": [""],"Property Owner": [""],"Property Mailing Name": [""],"Property Address": [""],"Property City": [""],"Property State": [""] ,"Property Zip": [""] ,"Tenant Name": [""],"Mailing Name": [""],"Mailing Address": [""],"Mailing City": [""],"Mailing State": [""],"Mailing Zip": [""],"Attorney Name": [""],"Attorney Address": [""],"Attorney Mailing Name": [""], "Attorney State": [""],"Attorney City": [""],"Attorney Zip": [""]})
				#applying color to the row axis 1 = row
				df1 = df1.style.apply(self.ExcelColorGray,axis=1)
				#df1 = df1.style.set_properties(**{'height': '300px'})
				#writing colored row to excel
				df1.to_excel(self.ExcelFile,index=False,sheet_name='Data',header=False,startrow=self.Rows)
				self.Rows = self.ExcelFile.sheets['Data'].max_row
				#then writing data
				df.to_excel(self.ExcelFile,index=False,sheet_name='Data',header=False,startrow=self.Rows)
				self.Rows = self.ExcelFile.sheets['Data'].max_row
			else:
				df.to_excel(self.ExcelFile,index=False,sheet_name='Data',header=False,startrow=self.Rows)
				self.Rows = self.ExcelFile.sheets['Data'].max_row
			self.LastCaseID = casenumber


		self.ExcelFile.save()





a = PublicCase()
#a.DataExtraction()
a.searchcase()
#print(a.addressfilter(""))


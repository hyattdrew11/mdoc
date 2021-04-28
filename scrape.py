import calendar
from datetime import date
import os
import platform
import sys
import urllib.request
import string
import time
import pdfkit
import json
import shutil
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import csv

print("START SCRIPT")

chrome_options = webdriver.ChromeOptions()
settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--kiosk-printing')

CHROMEDRIVER_PATH = '/Users/drewhyatt/chrome/chromedriver'
driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)


download_path = '/Users/drewhyatt/Downloads' # Path where browser save files
new_path = os.getcwd() + '/storage' # Path where to move file


url = 'https://www.ms.gov/mdoc/inmate/'
resultsUrl = 'https://www.ms.gov/mdoc/inmate/Search/GetSearchResults'
inmateUrl ='https://www.ms.gov/mdoc/inmate/Search/GetDetails/'
inmateIDS = []
az = list(string.ascii_lowercase)
# az = ('a')

searchInput = 'LastName'
inmateRows = []
columns = ["Inmate ID","Race","Sex","Date of Birth","Height","Weight","Complexion","Build","Eye Color","Hair Color","Entry Date","Location","UNIT","Location Change Date","Number of Sentences"]

def testCSV():
	colMatch = ["Inmate ID:","Race:","Sex:","Date of Birth:","Height:","Weight:","Complexion:","Build:","Eye Color:","Hair Color:","Entry Date:","Location:","UNIT:","Location Change Date:","Number of Sentences:"]
	driver.get(url)
	time.sleep(2)
	driver.get("https://www.ms.gov/mdoc/inmate/Search/GetDetails/232293")
	mainContent = driver.find_element_by_id("PageContentBody")
	tables = mainContent.find_elements_by_tag_name("table")
	inmateRow = columns
	for x in tables:
		trs = x.find_elements_by_tag_name("tr")
		for tr in trs:
			try:
				tds = tr.find_elements(By.TAG_NAME, "td")
				for i in tds:
					for index, col in enumerate(colMatch):
						if col in i.text:
							inmateRow[index] = i.text
			except IndexError:
				pass

	print(inmateRow)

def writeRow(inmateID):
	# print('===================================================================================')
	# print("WRITING INMATE ROW")
	colMatch = ["Inmate ID:","Race:","Sex:","Date of Birth:","Height:","Weight:","Complexion:","Build:","Eye Color:","Hair Color:","Entry Date:","Location:","UNIT:","Location Change Date:","Number of Sentences:"]
	mainContent = driver.find_element_by_id("PageContentBody")
	tables = mainContent.find_elements_by_tag_name("table")
	inmateRow = columns
	inmateRow[0] = inmateID
	for x in tables:
		trs = x.find_elements_by_tag_name("tr")
		for tr in trs:
			try:
				tds = tr.find_elements(By.TAG_NAME, "td")
				for i in tds:
					for index, col in enumerate(colMatch):
						if col in i.text:
							inmateRow[index] = i.text
							# print(i.text)
			except IndexError:
				pass
	# print(inmateRow)
	# print('===================================================================================')
	return inmateRow

def scrape():
	with open('storage/scrape_ledger.csv', 'w') as f:
		# using csv.writer method from CSV package
		write = csv.writer(f)
		write.writerow(columns)
		debugCount = 0 
		for x in az:
			tmpArray = []
			print("GETTING NEW LETTER")
			driver.get(url)
			inputElement = driver.find_element_by_id(searchInput)
			inputElement.send_keys(x)
			inputElement.send_keys(Keys.ENTER)
			table_id = driver.find_element_by_id('SearchResults')
			trs = table_id.find_elements_by_tag_name("tr")
			count = 2
		
			for x in trs:
				try:
					count += 1
					tds = trs[count].find_elements(By.TAG_NAME, "td")
					anchor = tds[0].find_elements_by_tag_name("a")
					for i in anchor:
						txt = i.get_attribute('innerHTML')
						tmpArray.append(txt)
						inmateIDS.append(txt)
				except IndexError:
					pass


			for x in tmpArray:
				details = inmateUrl+x
				driver.get(details)
				# WRITE INMATE INFO TO CSV
				row = writeRow(x)
				write.writerows([row])
				# inmateRows.append(row)
				driver.execute_script('window.print();')
				# debugCount += 1
				if debugCount > 5:
					break

				new_filename = x+'.pdf' # Set the name of file
				timestamp_now = time.time() # time now
				# Now go through the files in download directory
				for (dirpath, dirnames, filenames) in os.walk(download_path):
					for filename in filenames:
						if filename.lower().endswith(('.pdf')):
							full_path = os.path.join(download_path, filename)
							timestamp_file = os.path.getmtime(full_path) # time of file creation
							# if time delta is less than 15 seconds move this file
							if (timestamp_now - timestamp_file) < 15: 
								full_new_path = os.path.join(new_path, new_filename)
								os.rename(full_path, full_new_path)

		# ZIP STORAGE FOLDER
		shutil.make_archive(str(date.today()), 'zip', new_path)
		print("COMPLETE SCRAPE")
		f.close()
		exit()

if __name__ == '__main__':
	scrape()
	# testCSV()
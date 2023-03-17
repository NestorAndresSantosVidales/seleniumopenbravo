'''
*!
 * @SPMBOT Version 1.0
 *
 * This Bot was written to download a range SPM reports of Web Scrapping on OpenBravo website,
 * The Bot use a list of SPM id and Oracle Credentials to extract the report of a range of dates. 
 
 * Written by Nestor Santos https://nestorandressantosvidales.com/
 
 *
 * BSD license, all text here must be included in any redistribution.
 */


'''

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time 


theusername = ""
thepassword = ""
thedate     = "08-25-2022"
spmlist =['17245129', '16054315', '11141576', '18562641', '11109741', '10401851', '11383764', '11080382', '11769672', '11106598', '17486799', '10759875', '17357773', '11320992', '20914518', '12238297', '16182538', '21946159', '11780666', '11100204', '20617500', '15330195', '16624598', '15317099', '18394869', '17539390', '22031089', '11780486', '17473399', '17783773', '12640307', '11143667', '10963492', '11248226', '15962792', '15330159', '11095395', '17490265', '11568002', '11742984', '13078359', '16418685', '10859760', '10892952', '16625035', '11216022', '18393208', '10485508', '16578265', '17484937', '10918153', '11824649', '11338081', '10779727', '10896337', '10891842', '17224502', '11569994', '11300625', '10484178', '10889968', '11823225', '20893365', '21980406', '11570948', '10886259', '10724982', '21189295', '12212298', '15877918', '21726582', '16586941', '15464322', '11253462', '10890196', '17767837', '16885654', '18397480', '17361769', '11571443', '17458935', '10887019', '11164435', '11334382', '10810046', '11821892', '13150281', '12276282', '17210569', '21968122', '17329764', '17221107', '18394880', '21927086', '15333468', '21971960', '17263116', '11756883', '17826216', '21937444', '11311333', '17227977', '17224459', '11740696', '15293533', '17236029', '11766429', '17427617', '15291184', '16112630', '17211380', '16577406', '10999223', '11770825', '11726822', '11568164']


# Set path Selenium
CHROMEDRIVER_PATH = '/usr/local/bin/chromedriver'
s = Service(CHROMEDRIVER_PATH)
WINDOW_SIZE = "1920,1080"

options = Options()
options.add_argument("start-maximized")

download_dir = "/Users/admin/Desktop/spmbot/pathToDownloadDir"
chrome_options = webdriver.ChromeOptions()
preferences = {"download.default_directory": download_dir ,
               "directory_upgrade": True,
               "safebrowsing.enabled": True }
chrome_options.add_experimental_option("prefs", preferences)

for thespm in range(len(spmlist)):

	theurl="https://csaap.oracle.com/oalapp/web/spm/?tabId=8F9CA53965834621900034A7C378127E&criteria=%7B%22operator%22%3A%22and%22%2C%22_constructor%22%3A%22AdvancedCriteria%22%2C%22criteria%22%3A%5B%7B%22fieldName%22%3A%22searchKey%22%2C%22operator%22%3A%22iContains%22%2C%22value%22%3A%22"+ str(spmlist[thespm])+ " %22%2C%22_constructor%22%3A%22AdvancedCriteria%22%7D%5D%7D"
	
	driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

	print("Starting SPM: " +str(spmlist[thespm]) );
	# Get the response and print title
	driver.get(theurl)
	print("Login Loaded!")
	time.sleep(10)
	username = driver.find_element("id", "sso_username")
	password = driver.find_element("id", "ssopassword")
	username.send_keys(theusername)
	password.send_keys(thepassword)
	driver.find_element("id","signin_button").click()
	print("Login Passed :)")
	print("Loading OpenBravo Website...")

	#time.sleep(60)

	WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "isc_EA"))).click()

	#computation = driver.find_element("id","isc_EA").click()
	print("Computation Tab Clicked")

	time.sleep(50)

	wait = WebDriverWait(driver, 60)

	todate = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div/div/div/div[3]/div[1]/div/div[2]/div/table/tbody/tr/td[4]/div/nobr/span/table/tbody/tr/td[1]/input'))).send_keys(thedate)
	#08-25-2022 - 08-30-2022
	print("Choose Date Option Clicked")

	time.sleep(5)
	todate = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div/div/div/div[3]/div[1]/div/div[2]/div/table/tbody/tr/td[4]/div/nobr/span/table/tbody/tr/td[1]/input'))).send_keys(Keys.ENTER)


	time.sleep(200)

	download = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div/div[5]/div[45]/table/tbody/tr/td'))).click()
	print("Downloading Report...")



	time.sleep(120)
	print("SPM report downloaded :)")




	driver.close()
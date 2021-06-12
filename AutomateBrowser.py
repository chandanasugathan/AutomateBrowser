from selenium.common import exceptions
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import datetime
import math
import imagehash
from PIL import Image

REPO_PATH = 'C:\\Users\\cvelloth\\Downloads'
chrome_driver = os.path.join(REPO_PATH, "chromedriver.exe")
chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--enable-cookies")

chrome_options.add_experimental_option("prefs", {"profile.default_content_setting_values.cookies": 0})
chrome_options.add_experimental_option("prefs", {"profile.cookie_controls_mode": 0})
chrome_options.add_experimental_option("prefs", {"profile.block_third_party_cookies": False})
chrome_options.add_experimental_option("prefs", {"security.cookie_behavior": 0})
browser = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)
wait = WebDriverWait(browser, 100)
presence = EC.presence_of_element_located
visible = EC.visibility_of_element_located
clickable = EC.element_to_be_clickable


def play_youtube_video(url = "https://www.youtube.com/watch?v=-CmadmM5cOk"):
	test_duration = 235
	test_interval = 30
	test_stop_time = 60
	total_checks = 2
	
	browser.execute_script("window.open('about:blank','youtubetab');")
	browser.switch_to.window("youtubetab")
	browser.get(url)
	time.sleep(5)
	wait.until(visible((By.ID, "video-title")))
	
	video_duration = browser.execute_script("return document.getElementById('movie_player').getDuration()")
	print("Video duration (sec): {0}".format(video_duration))	
	print("\nStarting tests: {0}".format(datetime.datetime.now()))
	idx = 0
	has_ad = False
	total_skip_ad = 0
	ad_word = check_ad(has_ad, total_skip_ad)
	if has_ad :
		print (ad_word)
	else :
		skip_ad(has_ad, total_skip_ad)
	test_stop_time = time.time() + test_duration    
	total_checks = math.floor(test_duration / test_interval)
	while True:
		if time.time() > test_stop_time:
			break
		player_status = browser.execute_script("return document.getElementById('movie_player').getPlayerState()")
		
		if player_status == -1:
			wait.until(visible((By.XPATH, "//button[@class='ytp-large-play-button ytp-button']")))
			browser.find_element_by_xpath("//button[@class='ytp-large-play-button ytp-button']").click()
		video_start_time = time.time()
		video_stop_time = video_start_time + video_duration
		interval_time = time.time() + test_interval
		
		
		if interval_time < test_stop_time < video_stop_time:
			test_window(idx, total_checks)
			interval_time = time.time() + test_interval
			idx += 1
		time.sleep(test_interval)
	

def check_ad(has_ad, total_skip_ad):
	try:
		browser.find_element_by_xpath("//button[@class='ytp-ad-skip-button ytp-button']").click()
		has_ad = True
		total_skip_ad += 1
		print("{} ads skipped".format(total_skip_ad))
		return browser.find_element_by_xpath("//button[@class='ytp-ad-skip-button ytp-button']").text
	except:
		return None

def skip_ad(has_ad, total_skip_ad):
	try:
		has_ad = False		
		browser.find_element_by_class_name("//button[@class='ytp-ad-skip-button ytp-button']").click()
		total_skip_ad += 1
		print("{} ads skipped".format(total_skip_ad))
		time.sleep(1)
	except:
		return None

def test_window(idx, total_checks):
	print("\nChecking results of Test Interval [{0}/{1}]...".format(idx + 1, total_checks))
	try:
		try:
			screenshot = "screenshot" + "-" + str(idx) + ".png"
			browser.save_screenshot(screenshot)
		except NoSuchElementException:
			error = "Web server down!"
			print ("FAIL: {0} : {1}".format(datetime.datetime.now(), error))
			teardown()
		if idx > 0:
			prev_screenshot = "screenshot" + "-" + str(idx - 1) + ".png"
			prev_hash = imagehash.average_hash(Image.open(prev_screenshot))
			curr_hash = imagehash.average_hash(Image.open(screenshot))
			if prev_hash == curr_hash:
				print("{0} and {1} appear similar!".format(prev_screenshot, screenshot))
				error = "Browser Window Hang or Black Frame detected!"
				print("FAIL: {0} : {1}".format(datetime.datetime.now(), error))
				raise
			else:
				os.remove(prev_screenshot)

	except exceptions.WebDriverException:
		error = "Browser window closed!"
		print("FAIL: {0} : {1}".format(datetime.datetime.now(), error))
		raise
	print("Screenshot Check Success: {0}".format(datetime.datetime.now()))
	
def teams_login():
	browser.get("https://www.office.com/")
	time.sleep(5)
	try :
		wait.until(visible((By.ID, "hero-banner-sign-in-to-office-365-link")))
		browser.find_element_by_xpath('//*[@id="hero-banner-sign-in-to-office-365-link"]').click()
		EMAILFIELD = (By.ID, "i0116")
		PASSWORDFIELD = (By.ID, "i0118")
		NEXTBUTTON = (By.ID, "idSIButton9")
		TEAMSBUTTON = (By.XPATH, '//*[@id="ShellSkypeTeams_link"]')
		COOKIEBLOCK = (By.XPATH,'//*[@id="cookie-controls-toggle"]')
		SETTINGSBUTTON = (By.XPATH, '//*[@id="settings-menu-button"]')
		SETTINGSDROPDOWN = (By.XPATH, "//button[@class='ts-sym left-align-icon']")
		email = ""
		password = ""
		wait.until(clickable(EMAILFIELD)).send_keys(email)
		wait.until(clickable(NEXTBUTTON)).click()
		wait.until(clickable(PASSWORDFIELD)).send_keys(password)
		wait.until(clickable(NEXTBUTTON)).click()
		wait.until(clickable(NEXTBUTTON)).click()
		
		browser.execute_script("window.open('about:blank','secondtab');")
		browser.switch_to.window("secondtab")
		browser.get('chrome://newtab')
		wait.until(clickable(COOKIEBLOCK)).click()
		#time.sleep(5)
		browser.close()
		browser.switch_to.window(browser.window_handles[0])
		wait.until(clickable(TEAMSBUTTON)).click()
		time.sleep(40) #Wait for teams page to load
		
		chwd = browser.window_handles
		p = browser.current_window_handle
		for w in chwd:
		#switch focus to child window
			if(w!=p):
				browser.switch_to.window(w)
		
		
		time.sleep(10)		
		wait.until(clickable(SETTINGSBUTTON)).click()
		wait.until(clickable(SETTINGSDROPDOWN)).click()			
		time.sleep(10)
		
		#teardown()
	except :#(exceptions.TimeoutException, exceptions.WebDriverException):
		print("Exception Occurred")
		teardown()


def officeapps_login():
	browser.get("https://www.office.com/")
	mainwindow = browser.current_window_handle
	time.sleep(5)
	try :
		SIGNINBUTTON = (By.ID, "hero-banner-sign-in-to-office-365-link")
		wait.until(visible(SIGNINBUTTON)).click()
		#browser.find_element_by_xpath('//*[@id="hero-banner-sign-in-to-office-365-link"]').click()
		EMAILFIELD = (By.ID, "i0116")
		PASSWORDFIELD = (By.ID, "i0118")
		NEXTBUTTON = (By.ID, "idSIButton9")
		#Teams Buttons Xpaths
		TEAMSBUTTON = (By.XPATH, '//*[@id="ShellSkypeTeams_link"]')
		COOKIEBLOCK = (By.XPATH,'//*[@id="cookie-controls-toggle"]')
		SETTINGSBUTTON = (By.XPATH, '//*[@id="settings-menu-button"]')
		SETTINGSDROPDOWN = (By.XPATH, "//button[@class='ts-sym left-align-icon']")
			
		
		email = ''
		password = ''
		wait.until(clickable(EMAILFIELD)).send_keys(str(email))
		wait.until(clickable(NEXTBUTTON)).click()
		wait.until(clickable(PASSWORDFIELD)).send_keys(password)
		wait.until(clickable(NEXTBUTTON)).click()
		wait.until(clickable(NEXTBUTTON)).click()
		
		browser.execute_script("window.open('about:blank','secondtab');")
		browser.switch_to.window("secondtab")
		browser.get('chrome://newtab')
		wait.until(clickable(COOKIEBLOCK)).click()
		#time.sleep(5)
		browser.close()
		browser.switch_to.window(browser.window_handles[0])
		wait.until(clickable(TEAMSBUTTON)).click() #Click for teams to start.		
		time.sleep(40) #Wait for teams page to load
		
		chwd = browser.window_handles
		for w in chwd:
		#switch focus to child window
			if(w!=mainwindow):
				browser.switch_to.window(w)	
		
		time.sleep(10)		
		wait.until(clickable(SETTINGSBUTTON)).click()
		wait.until(clickable(SETTINGSDROPDOWN)).click()			
		time.sleep(10)
		
		#Opening Word from office 365 account
		browser.execute_script("window.open('about:blank','wordtab');")
		browser.switch_to.window("wordtab")
		browser.get('https://www.office.com/launch/word?auth=2&username='+str(email)+'&login_hint='+str(email))
		FILEOPTION = (By.XPATH, '//*[@id="recommended-mru-mru_item_1"]')
		wait.until(visible(FILEOPTION)).click() #Click the first recommended file
		time.sleep(3) #Wait before opening new window
		
		#Opening Excel from office 365 account
		browser.execute_script("window.open('about:blank','exceltab');")
		browser.switch_to.window("exceltab")
		browser.get('https://www.office.com/launch/excel?auth=2&username='+str(email)+'&login_hint='+str(email))
		FILEOPTION = (By.XPATH, '//*[@id="recommended-mru-mru_item_1"]')
		wait.until(visible(FILEOPTION)).click() #Click the first recommended file
		time.sleep(10)	#Wait before opening new window
		
		#Opening One Note from office 365 account		
		browser.execute_script("window.open('about:blank','onenotetab');")
		browser.switch_to.window("onenotetab")
		browser.get('https://www.office.com/launch/onenote?auth=2&username='+str(email)+'&login_hint='+str(email))
		FILEOPTION = (By.XPATH, '//*[@id="unpinned-mru-mru_item_0"]')
		wait.until(visible(FILEOPTION)).click() #Click the first file in list
		time.sleep(3) #Wait before opening new window
		
		#Opening Outlook from office 365 account
		browser.switch_to.window(mainwindow)
		OUTLOOKBUTTON = (By.XPATH, '//*[@id="ShellMail_link"]')
		wait.until(clickable(OUTLOOKBUTTON)).click()
		time.sleep(3) #Wait before opening new window
		
		#Take Screenshot of all open tabs
		chwd = browser.window_handles
		tab_index = 0
		for w in chwd:
			browser.switch_to.window(w)
			time.sleep(3) #Wait before taking screenshot
			screenshot = "screenshot_tab-"+str(tab_index+1)+".png"
			browser.save_screenshot(screenshot)
			tab_index = tab_index+1
			
		
	except Exception as error:
		print("Exception Occurred")
		raise error		
	
def teardown():
	browser.close()
	browser.quit()
	
	
if __name__ == "__main__":
	try:
		try :
			officeapps_login() #Replace with function teams_login() if need to check only login into office 365 and teams app.
		except Exception as error:
			print ("Error Launching Teams")
			raise error
		try :
			play_youtube_video()
		except Exception as error:
			print ("Error Verifying youtube video")
			raise
		teardown()
	except Exception as error:
		print (error)
		teardown()
	
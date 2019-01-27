
from selenium import webdriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as condi
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pickle
from datetime import datetime
import xlsxwriter
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import MoveTargetOutOfBoundsException
from os import path
import glob
from selenium.webdriver.common.action_chains import ActionChains

import time

workbook = xlsxwriter.Workbook('d:/twitterPostData.xlsx')
dataSource = workbook.add_worksheet('DataSource')
watcherSource = workbook.add_worksheet('WatcherSource')
watchingData = workbook.add_worksheet('WatchingData')
tempData = workbook.add_worksheet('TempData')

watcherSource.write(49, 0, "1")
workbook.close()

time.sleep(0.2)

workbook = xlsxwriter.Workbook('d:/twitterPostData.xlsx')
dataSource = workbook.add_worksheet('DataSource')
watcherSource = workbook.add_worksheet('WatcherSource')
watchingData = workbook.add_worksheet('WatchingData')
tempData = workbook.add_worksheet('TempData')

profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", 'D:\\insight')
profile.set_preference("browser.helperApps.neverAsk.saveToDisk",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, "
                       "application/x-excel, application/x-msexcel, application/excel, "
                       "application/vnd.ms-excel")

driver = webdriver.Firefox(firefox_profile=profile)
wait = WebDriverWait(driver, 1)
element = wait.until(condi.presence_of_all_elements_located)

driver.get("https://twitter.com/login")

#It's optional to load twitter cookies
try:
    cookies = pickle.load(open("twCookies.pkl", "rb"))
    for cookie in cookies:
        driver.add_cookie(cookie)
except(OSError, IOError, FileNotFoundError) as exception:
   #handling errors

driver.get('https://twitter.com/Bitcoin/status/1001306047926095873')

c = 0

# while c < 1:
#     buttonCandidate = driver.find_element_by_css_selector(
#         '.ThreadedConversation-showMoreThreadsButton.u-textUserColor')
#     print(buttonCandidate)
#     if buttonCandidate is not None:
#         c = c + 1
#         buttonCandidate.click()

countThreads = 0
countThreadsEX = 0
# while c < 1:
#    mySpinner: WebElement = driver.find_element_by_class_name('spinner')
#    if mySpinner is not None:
#        c = c + 1
#        try:
#            time.sleep(1)
#        except:
tempthreadcount = 0
stopcounter = 0
tempthreadcountex = 0
stopcounterex = 0


while c < 1:
    allthreads = driver.find_elements_by_class_name('ThreadedConversation')
    allthreadsEX = driver.find_elements_by_class_name('ThreadedConversation--loneTweet')
    print('stopcounter ', stopcounter)
    print('stopcounterex', stopcounterex)

    if countThreads > 0:
        if tempthreadcount == allthreads[countThreads - 1]:
            stopcounter = stopcounter + 1
            if stopcounter > 10 and stopcounterex > 10:
                c = 1
        else:
            stopcounter = 0

        tempthreadcount = allthreads[countThreads - 1]
        driver.execute_script('arguments[0].scrollIntoView();', allthreads[countThreads-1])

        time.sleep(0.1)

    if countThreadsEX > 0:
        if tempthreadcountex == allthreadsEX[countThreadsEX - 1]:
            stopcounterex = stopcounterex + 1
            if stopcounter > 10 and stopcounterex > 10:
                c = 1
        else:
            stopcounterex = 0

        tempthreadcountex = allthreadsEX[countThreadsEX - 1]
        driver.execute_script('arguments[0].scrollIntoView();', allthreadsEX[countThreadsEX - 1])

        time.sleep(0.1)

    countThreads = 0
    countThreadsEX = 0
    for thread in allthreads:
        countThreads = countThreads + 1
        print('countThreads ', countThreads)
    for threadex in allthreadsEX:
        countThreadsEX = countThreadsEX + 1
        print('countThreadsEX', countThreadsEX)
    time.sleep(0.1)

allthreads = driver.find_elements_by_class_name('ThreadedConversation')
allthreadsEX = driver.find_elements_by_class_name('ThreadedConversation--loneTweet')

i = 0

for thread in allthreads:
    plist: WebElement = thread.find_elements_by_tag_name('p')
    for p in plist:
        watcherSource.write(i, 0, p.get_attribute('outerHTML'))
        i = i + 1

for threadex in allthreadsEX:
    plist: WebElement = threadex.find_elements_by_tag_name('p')
    for p in plist:
        watcherSource.write(i, 0, p.get_attribute('outerHTML'))
        i = i + 1

workbook.close()

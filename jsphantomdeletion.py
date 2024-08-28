import ntpath
import logging
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import xlrd
import xlwt
from tempfile import TemporaryFile
import glob
import os
import win32com.client
import time
import urllib2



myname = ntpath.basename(__file__).split('.')[0]
module_logger = logging.getLogger(myname)

USER = 'zvamsre'
PASS = 'Mar@2018'
EMAIL='roshini.shekar@wipro.com'
COMMENT='Please Deactivate this ID from the account'


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders.Item("Ericsson-ISIT-Deletion") # "6" refers to the index of a folder - in this case,
                                            # the inbox. You can change that number to reference
                                            # any other folder
messages = inbox.Items
message = messages.GetFirst()
rec_time = message.CreationTime
body_content = message.body
subj_line = message.subject
l=len(body_content)
Final_body=body_content[0:(l-1)]
Final_body=Final_body.split(';')
print Final_body

Email_ID = Final_body[-1]
print Email_ID

COMMENT="Please Deactivate Immediately"

def deletion():
    if not sys.warnoptions:
        warnings.simplefilter("ignore")
    #----Chrome---#Headless
    #options = webdriver.ChromeOptions()
    #options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
    #options.add_argument('--headless')
    #CHROMEDRIVER_PATH = "C:\Python27\Creation\chromedriver.exe"
    #browser = webdriver.Chrome(CHROMEDRIVER_PATH, chrome_options=options)
    #print("Chrome Headless Browser Invoked")
    #----PhantomJs---#
    browser = webdriver.PhantomJS()
    module_logger.info("PhantomJs Invoked")
    
    url = 'https://login2.internal.ericsson.com/login/V2_0/login.fcc?TYPE=33619969&REALMOID=06-1157e95f-1734-4756-b912-af0e0bc92b32&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=$SM$cauIEyZNipFMYNggDqRIvQLV%2fZrbIWSUs36Y5EAqFNodhbStPPG5XMiaYr%2flfk%2b6&TARGET=$SM$http%3a%2f%2fisignum%2einternal%2eericsson%2ecom%2fexternal%2f#isignum/home' 
    browser.get(url)

    
    assert 'Ericsson - Enterprise Sign On' in browser.title
    user = browser.find_element_by_xpath('//*[@id="user"]')
    user.send_keys(USER)
    password = browser.find_element_by_xpath('//*[@id="password"]')
    password.send_keys(PASS)
    
    browser.find_element_by_css_selector("#IMAGE1").click()
    module_logger.info("Logging in")
    delay = 35 # seconds
    try:
        myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, 'Home')))
        module_logger.info("logged in")
        module_logger.info("Page is ready!")
    except TimeoutException:
        module_logger.info("Loading took too much time!")
    assert 'ISIGNUM | External Identity Management' in browser.title

    emailid_for_deletion = browser.find_element_by_css_selector("body > div.eaContainer-applicationHolder > div > div.eaiSignum-mainContentScrollable > div > div > div > div > div.eaiSignum-wHome-userSection > div.eaiSignum-wHome-userSection-searchWidgetHolder > div > form > div.eaiSignum-wSearch-generalSearchSection > div:nth-child(4) > div.eaiSignum-wSearch-form-inputHolder > span > input")
    emailid_for_deletion.send_keys(Email_ID)
    module_logger.info("Email id entered is found!")
    browser.find_element_by_css_selector("body > div.eaContainer-applicationHolder > div > div.eaiSignum-mainContentScrollable > div > div > div > div > div.eaiSignum-wHome-userSection > div.eaiSignum-wHome-userSection-searchWidgetHolder > div > form > div.eaiSignum-wSearch-generalSearchSection > div:nth-child(4) > button > span").click()
    browser.implicitly_wait(10) 
    browser.find_element_by_css_selector("body > div.eaContainer-applicationHolder > div > div.eaiSignum-mainContentScrollable > div > div > div > div > div.eaiSignum-wHome-userSection > div.eaiSignum-wHome-userSection-tableDiv > div > div > div.elISIGNUMLib-wTableLib-tableHolder > div > div.elTablelib-Table-wrapper.eb_scrollbar > table > tbody > tr > td:nth-child(10) > a:nth-child(1) > i").click()
    browser.find_element_by_css_selector("body > div.eaContainer-applicationHolder > div > div.eaiSignum-mainContentScrollable > div > div > div > div > div.eaiSignum-wUserManagement-contentHolder > div > form > div.eaiSignum-wUserDetails-userDataSection > div.eaiSignum-wUserDetails-block.eaiSignum-wUserDetails-deactivateTextsBlock > div.eaiSignum-wUserDetails-deactivateCheckbox > input").click()
    comment = browser.find_element_by_css_selector("body > div.eaContainer-applicationHolder > div > div.eaiSignum-mainContentScrollable > div > div > div > div > div.eaiSignum-wUserManagement-contentHolder > div > form > div.eaiSignum-wUserDetails-userDataSection > div.eaiSignum-wUserDetails-block.eaiSignum-wUserDetails-commentBlock > div.eaiSignum-wUserDetails-field > div > span.elISIGNUMLib-wCustomTextarea > textarea")
    comment.send_keys(COMMENT)
    #browser.find_element_by_css_selector("body > div.eaContainer-applicationHolder > div > div.eaiSignum-mainContentScrollable > div > div > div > div > div.eaiSignum-wUserManagement-contentHolder > div > form > div.eaiSignum-wUserDetails-buttonHolder > button.eaiSignum-wUserDetails-submitBtn.ebBtn.ebBtn_large.ebBtn_color_darkBlue").click()
    module_logger.info("ID deactivation request raised!")

def main():
    module_logger.setLevel(logging.DEBUG)
    # create file handler which logs even debug messages
    fh = logging.FileHandler('C:\Python27\XIDDELETION\ci_dev_infra.log')
    fh.setLevel(logging.DEBUG)
    # create console handler with a higher log level
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    # create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    fh.setFormatter(formatter)
    # add the handlers to logger
    module_logger.addHandler(ch)
    module_logger.addHandler(fh)
    
    deletion()
    


if __name__ == '__main__':
    main()
    

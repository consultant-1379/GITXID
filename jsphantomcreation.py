import ntpath
import logging
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
import sys
import xlrd
import xlwt
from tempfile import TemporaryFile
import glob
import os
import win32com.client 
import time
import urllib2
import csv
from pandas import ExcelWriter
import warnings


myname = ntpath.basename(__file__).split('.')[0]
module_logger = logging.getLogger(myname)

USER = 'zvamsre'
PASS = 'Mar@2018'

def fetchattachmentfromoutlook():
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders.Item("Ericsson-ISIT-Creation") # "6" refers to the index of a folder - in this case,
                                            # the inbox. You can change that number to reference
                                            # any other folder
    messages = inbox.Items
    message = messages.GetLast()
    rec_time = message.CreationTime
    body_content = message.body
    subj_line = message.subject
    l=len(body_content)
    Final_body=body_content[0:(l-2)]
    Final_body=Final_body.split(';')
    


    ofile  = open('C:\Python27\Creation\CreationOfXid.csv', "wb")
    writer = csv.writer(ofile, delimiter=';')
    writer.writerow(["Company ID","Company name","Managed unit","Employee number","First name","Last name","Initials","Name affix","Email address",
                     "Fixed phone number","Extension","Mobile phone number","End of assignment","User country","User state","User city","Comment"])

    Comments = Final_body[-1]
    City = Final_body[-2]
    State = Final_body[-3]
    Country = Final_body[-4] 
    End_Of_Ass = Final_body[-5]
    Mobile = Final_body[-6]
    Extention = Final_body[-7]
    Fixed_Phone_No = Final_body[-8]
    Email = Final_body[-9]
    Name_aff = Final_body[-10]
    Initials = Final_body[-11]
    Last_name = Final_body[-12]
    First_name = Final_body[-13]
    Emp_no = Final_body[-14]
    Managed_unit = Final_body[-15]
    Comp_name = Final_body[-16]
    Comp_id = Final_body[-17][-3:]
    writer.writerow([Comp_id,Comp_name,Managed_unit,Emp_no,First_name,Last_name,Initials,Name_aff,Email,Fixed_Phone_No,Extention,Mobile,End_Of_Ass,Country,State,City,Comments])
    module_logger.info("CSV Generated..")



    
def creation():
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

    browser.find_element_by_xpath('//*[@id="IMAGE1"]').click()
    module_logger.info("Logging in")
    delay = 35 # seconds
    try:
        myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, 'Home')))
        module_logger.info("logged in")
        module_logger.info("Page is ready!")
    except TimeoutException:
        module_logger.info("Loading took too much time!")
        browser.quit()
        sys.exit()
    
    assert 'ISIGNUM | External Identity Management' in browser.title
    
    users=WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div/div[2]/span[1]'))).click()#Users
    module_logger.info("users")
    
    actions=WebDriverWait(browser, 10)
    checkbox=actions.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[1]/div/div/div/div[2]/div/div/a[3]/span')))
    browser.execute_script("arguments[0].click();",checkbox)#Bulk Actions
    module_logger.info("bulk actions")
    
    browser.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/div[1]/div[2]/input").send_keys(glob.glob('C:\Python27\Creation\CreationOfXid.csv'))#choosefile
    #browser.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/div[1]/div[2]/input").send_keys(glob.glob('C:\Users\SR376502\Documents\creationOfXid.csv'))#choosefile

    module_logger.info("CSV File has been uploaded to ISIGNUM")
    
    browser.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/div[4]/button[1]").click()#confirm

    
    if browser.find_elements_by_xpath('//*[@class="elISIGNUMLib-wEditableCell-input ebInput elISIGNUMLib-wEditableCell-input_invalid"]'):
        module_logger.info("Request Denied!!!")
        module_logger.info("===========================")
        browser.quit()
    else:
        #browser.find_element_by_xpath("/html/body/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/form/div[5]/div[2]/button[1]").click()#submit
        module_logger.info("Request Raised!!!")
        module_logger.info("===========================")
        browser.quit()

    

def main():
    module_logger.setLevel(logging.DEBUG)
    # create file handler which logs even debug messages
    fh = logging.FileHandler('C:\Python27\Creation\ci_dev_infra.log')
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
    fetchattachmentfromoutlook()
    creation()
    


if __name__ == '__main__':
    main()


    

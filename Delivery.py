
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from datetime import datetime
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
import getpass
USER_NAME = getpass.getuser()


def get_filefolder_path(name):
    currentpath = os.path.realpath(__file__)
    print(currentpath)
    # add_to_startup(currentpath)

    patharr = currentpath.split('\\')
    patharr[len(patharr)-1] = name
    ptr = ('\\').join(patharr)
    return ptr

def logger(downfile):
    now = datetime.now()
    log = now.strftime("[%Y/%m/%d %H:%M:%S] <{}> downloaded\n").format(downfile)
    log_file_path = get_filefolder_path('log.log')
    with open(log_file_path, 'a') as logfile:
        logfile.write(log)
        logfile.close()

def makecustompath(header, tag):
    prefix = 'C:/Users/' + USER_NAME + '/Downloads/OpAvailExportedData ('

    backfix = ').xlsx'

    # Source path 
    source = None
    ptr = "C:/Users/" + USER_NAME + "/Downloads/OpAvailExportedData.xlsx"

    existflag = True
    id = 1 
    while existflag:
        source = ptr
        ptr = prefix + str(id) + backfix
        existflag = os.path.exists(ptr)
        id += 1
    # Destination path 

    now = datetime.now()

    download_path = get_filefolder_path(tag)
    try:
        os.mkdir(download_path)
    except:
        pass
    filename = header + '_' + tag + '_' + now.strftime("%Y-%m-%d") + '.xlsx'

    destination1 = download_path + '\\' + filename
    destination = destination1.replace("\\", "/")
    print(destination)
    # Copy the content of 
    # source to destination 
    dest = shutil.copyfile(source, destination) 
    os.remove(source)
    logger(filename)

def init_driver():

    driver = webdriver.Chrome(ChromeDriverManager().install())

    return driver

def get_url(driver, url):
    driver.get(url)
    time.sleep(3)
    driver.find_element_by_xpath('//*[@id="WebSplitter1_tmpl1_ContentPlaceHolder1_ddlCycleDD"]/div/table/tbody/tr/td[1]/input').clear()
    time.sleep(2)
    selection = 'EVENING'
    search_input_element = driver.find_element(by=By.XPATH, value='//*[@id="WebSplitter1_tmpl1_ContentPlaceHolder1_ddlCycleDD"]/div/table/tbody/tr/td[1]/input')
    search_input_element.send_keys(selection)
    driver.find_element_by_xpath('//*[@id="WebSplitter1_tmpl1_ContentPlaceHolder1_HeaderBTN1_btnDownload"]').click()


def main(tag):
    driver = init_driver()
    link = 'https://pipeline2.kindermorgan.com/Capacity/OpAvailPoint.aspx?code=' + tag
    get_url(driver, link)
    time.sleep(3)
    print('excel downloaded')
    driver.close()
    makecustompath('Delivery',tag)
   


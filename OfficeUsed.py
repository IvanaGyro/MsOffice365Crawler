#-*- coding:utf8 -*-
import csv
import pickle
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

defaultAccountUrl = 'https://portal.office.com/account/'
defaultAccount404Url = 'https://portal.office.com/account/404error'
defaultLogin404Url = 'https://login.microsoftonline.com/error404'
defaultAccountFn = 'officeAcct.csv'
defaultCookiesFn = 'office_cookies.pkl'
defaultInstalledFn = 'office_used.csv'
driverPath = 'chromedriver.exe'
maxTimeout = 10

csvFields = {
    'usr':'usr',
    'pwd':'pwd',
    'installsUsedVal':'installed number'
}


class wait_installsUsedVal_load(object):
    def __init__(self, locator=(By.ID, 'installsUsedVal')):
        self.locator = locator

    def __call__(self, driver):
        return len(EC._find_element(driver, self.locator).text)


def load_officeAcct(fn):
    accountLs = []
    with open(fn, 'r') as fd:
        csvReader = csv.DictReader(fd)
        for line in csvReader:
            accountLs.append(dict(line))
    return accountLs


def load_officeCookies(fn):
    with open(fn, 'rb') as fd:
        return pickle.load(fd)


def login_office_account(driver, usr, pwd, accountUrl=defaultAccountUrl, waitLoading=True):
    driver.get(accountUrl)
    driver.find_element_by_id('cred_userid_inputtext').send_keys(usr)
    driver.find_element_by_id('cred_password_inputtext').send_keys(pwd)
    driver.find_element_by_id('cred_sign_in_button').click()
    if waitLoading:
        WebDriverWait(driver, maxTimeout).until(wait_installsUsedVal_load()) #installsUsedVal is loaded by AJAX


def get_office_account_cookies(driver, usr, pwd, accountUrl=defaultAccountUrl, login404Url=defaultLogin404Url):
    login_office_account(driver, usr, pwd, accountUrl=accountUrl, waitLoading=False)
    cookies = driver.get_cookies()

    #delete cookie with accountUrl
    driver.delete_all_cookies()

    #delete cookie with the login page
    driver.get(login404Url)
    driver.delete_all_cookies()

    return cookies


def save_installsUsedValLs(installsUsedValLs, fn):
    with open(fn, 'w') as fd:
        csvWriter = csv.DictWriter(fd, [csvFields['usr'], csvFields['installsUsedVal']], dialect='unix')
        csvWriter.writeheader()
        csvWriter.writerows(installsUsedValLs)


def save_all_account_cookies(accountFn=defaultAccountFn, outputFn=defaultCookiesFn, accountUrl=defaultAccountUrl, login404Url=defaultLogin404Url):
    driver = webdriver.Chrome(executable_path=driverPath)
    accountLs = load_officeAcct(accountFn)
    cookiesLs = []
    for account in accountLs:
        cookiesLs.append({'usr':account[csvFields['usr']], 'cookies':get_office_account_cookies(driver, account[csvFields['usr']], account[csvFields['pwd']], accountUrl=accountUrl, login404Url=login404Url)})
    driver.close()

    with open(outputFn, 'wb') as fd:
        pickle.dump(cookiesLs , fd)


def save_all_installsUsedVal_with_login(accountFn=defaultAccountFn, outputFn=defaultInstalledFn, accountUrl=defaultAccountUrl, login404Url=defaultLogin404Url):
    driver = webdriver.Chrome(executable_path=driverPath)
    accountLs = load_officeAcct(accountFn)
    installsUsedValLs = []
    for account in accountLs:
        login_office_account(driver, account[csvFields['usr']], account[csvFields['pwd']], accountUrl=accountUrl)
        installsUsedVal = driver.find_element_by_id('installsUsedVal').text
        installsUsedValLs.append({csvFields['usr']:account['usr'], csvFields['installsUsedVal']:installsUsedVal})
        
        #delete cookie with accountUrl
        driver.delete_all_cookies()

        #delete cookie with the login page
        driver.get(login404Url)
        driver.delete_all_cookies()
    driver.close()

    save_installsUsedValLs(installsUsedValLs, outputFn)


def save_all_installsUsedVal_with_cookies(cookieFn=defaultCookiesFn, outputFn=defaultInstalledFn, accountUrl=defaultAccountUrl, account404Url=defaultAccount404Url):
    driver = webdriver.Chrome(executable_path=driverPath)
    cookies = load_officeCookies(cookieFn)
    installsUsedValLs = []
    for account in cookies:
        #add cookies
        driver.get(account404Url) #can only add cookies for the same domain
        for cookie in account['cookies']:
            driver.add_cookie(cookie)

        #get installsUsedVal
        driver.get(accountUrl)
        WebDriverWait(driver, maxTimeout).until(wait_installsUsedVal_load()) #installsUsedVal is loaded by AJAX
        installsUsedVal = driver.find_element_by_id('installsUsedVal').text
        installsUsedValLs.append({csvFields['usr']:account['usr'], csvFields['installsUsedVal']:installsUsedVal})

        driver.delete_all_cookies()

    driver.close()

    save_installsUsedValLs(installsUsedValLs, outputFn)






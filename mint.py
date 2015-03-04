'''
@author: Kyle Dove

Purpose: Log into Mint/ Pull Account Information/ Write to Excel

Modification History:

DATE           Author               Description   
===========    ===============      =============
Aug 10 2014    Kyle Dove            Created.
Aug 11 2014    Kyle Dove            Added Logging.
                                    Desperately needs code cleanup.
                                    Completed for testing.
Aug 19 2014    Kyle Dove            Modified Wait time for site
                                    so that the site has time to
                                    refresh.
Sep 12 2014    Kyle Dove            Fixes applied.
Sep 26 2014    Kyle Dove            Updated version of Selenium for FireFox update
                                    Incorporated BeautifulSoup for html parsing.
Dec 12 2014    Kyle Dove            Added 10 second wait to give time for web page
                                    to load.
Dec 18 2014    Kyle Dove            Added functionality for putting accounts in an 
                                    array to reduce repeat code.
Mar 03 2015    Kyle Dove            Added more logging.
                                    Modified to look for nickname of account first.
                                    Then after the name is found, search for balance based on nickname.
                                    Changed to record ALL accounts with nickname. mintArray file no longer needed.

TODO: cleanup
 
'''

import time
import datetime
import re
import logging
import BeautifulSoup
from selenium import webdriver
from xlutils.copy import copy 
from xlrd import open_workbook 
from xlwt import Formula, XFStyle

#setup log
# create logger with 'spam_application'
logger = logging.getLogger('thin_mint')
logger.setLevel(logging.INFO)
logdate = time.strftime("%Y%m%d")
# create file handler which logs even debug messages
fh = logging.FileHandler('mintLog_' + logdate + '.log')
fh.setLevel(logging.INFO)
# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.ERROR)
# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)
# add the handlers to the logger
logger.addHandler(fh)
logger.addHandler(ch)

logger.info('Script Initiated')
mintPath = 'BalanceCopy.xls'

#get credentials
logger.info('Getting Credentials')
separator=":"
fileIN = open('mintCreds', "r")
line = fileIN.readline()

while line:
    sout=line.split(separator)
    user=sout[0]
    passwd=sout[1]
    line = fileIN.readline()

#get array
logger.info('Getting Array Info')
#fileIN = open('mintArray', "r")
#line = fileIN.readline()
#nums = []
combos = []
balances = []
newBalances = []
assets = []
liabilities = []
totalAsset = 0.0
totalLiability = 0.0
column = 0

while line:
    nums = line.split(separator)
    line = fileIN.readline()

driver = webdriver.Firefox()
driver.get('https://wwws.mint.com/login.event?task=L')
logger.info('Launching Web Browser')

time.sleep(10)
logger.info('Sleeping for 10 seconds')

#Find Username and Password Elements by Id
UsernameElement = driver.find_element_by_id("form-login-username")
PasswordElement = driver.find_element_by_id("form-login-password")

UsernameElement.send_keys(user)
PasswordElement.send_keys(passwd)

logger.info('Submitting Form')
driver.find_element_by_id("submit").click()

time.sleep(10) #Waits 10 seconds so that the site has time to refresh for new data
#TODO: Add while loop, check for refreshing indicator, once all indicators are gone, continue.
logger.info('Sleeping for 10 seconds')
html_source = driver.page_source

logger.info('Initiating Soup')
soup = BeautifulSoup.BeautifulSoup(html_source) 

nicknames = soup.findAll('span', attrs={'class' : 'nickname'})
print(nicknames)

for nickname in nicknames:
    txt = nickname.find(text=True)
    print "Nickname: " + txt.strip()
    parent = nickname.parent
    parent = parent.parent
    for balance in parent.findAll('span', attrs={'class' : 'balance'}):
        print('Balance: ' + balance.text)
        combo = []
        combo.append(txt.strip())
        combo.append(balance.text)
        combos.append(combo)

# Log ALL balances (for now)
counter = 0
for combo in combos:
    counterStr = counter
    counterStr = str(counterStr)
    logger.info(counterStr + ' --- Nickname: ' + combo[0] + ' --- Balane: ' + combo[1])
    balances.append(combo[1])
    counter = counter + 1
    # for item in combo:
    #     print(item)

logger.info('Iterating through balances')
for balance in balances:
    balance = re.sub('[$,]', '', balance)
    balance = float(balance)
    newBalances.append(balance)
    if balance > 0:
        assets.append(balance)
        print(str(balance) + ' is an asset')
    else:
        liabilities.append(balance)
        print(str(balance) + ' is a liability')

#find total assets
for asset in assets:
    totalAsset = totalAsset + asset
logger.info('Total Asset: ' + str(totalAsset))

#find total liabilities
for liability in liabilities:
    totalLiability = totalLiability + liability
logger.info('Total Liability: ' + str(totalLiability))

#Start Excel
logger.info('Starting Excel')
rb = open_workbook(mintPath,formatting_info=True)
r_sheet = rb.sheet_by_index(0)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

num_rows = r_sheet.nrows
strRows = str(num_rows + 1)
StrRowsMinOne = str(num_rows)

logger.info('Number of Rows: ' + str(num_rows))
logger.info('Date to insert: ' + time.strftime("%m/%d/%Y"))

myDate = time.strftime("%m/%d/%Y")

dateStyle = XFStyle()
dateStyle.num_format_str='M/D/YY'

currencyStyle = XFStyle()
currencyStyle.num_format_str = '$#,##0.00'

#row, col, text
w_sheet.write(num_rows,0,datetime.datetime.now(),dateStyle)

for i in range (len(newBalances)):
    column = int(i+1)
    newBalance = newBalances[i]
    w_sheet.write(num_rows,column,newBalance,currencyStyle)

#Total Assets
column = column + 1
w_sheet.write(num_rows,column,totalAsset,currencyStyle)
#Asset Change
column = column + 1
w_sheet.write(num_rows,column,Formula('O' + strRows + '-O' + StrRowsMinOne),currencyStyle)
#Total Liabilities
column = column + 1
w_sheet.write(num_rows,column,totalLiability,currencyStyle)
logger.info('Saving Excel')
wb.save(mintPath)


#logEndTime = time.strftime('%m/%d/%Y %I:%M:%S %p')
logger.info('Finished script. Closing browser.')

#Close Down
driver.close()

print('Done')
'''
@author: Kyle Dove

Purpose: Log into Mint/ Pull Account Information/ Write to Excel

Modification History:

DATE           Author               Description   
===========    ===============      =============
Aug 10 2014    Kyle Dove            Created
Aug 11 2014    Kyle Dove            Added Logging
                                    Desperately needs code cleanup
                                    Completed for testing
Aug 19 2014    Kyle Dove            Modified Wait time for site
                                    so that the site has time to
                                    refresh
Sep 12 2014    Kyle Dove            Fixes applied
Sep 26 2014    Kyle Dove            Updated version of Selenium for FireFox update
                                    Incorporated BeautifulSoup for html parsing
Dec 12 2014    Kyle Dove            Added 10 second wait to give time for web page
                                    to load
Dec 18 2014    Kyle Dove            Add functionality for putting accounts in an 
                                    array to reduce repeat code
 
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
logdate = time.strftime("%Y%m%d")
logfile = 'mintLog_' + logdate + '.log'
logging.basicConfig(filename=logfile,level=logging.INFO)
logStartTime = time.strftime('%m/%d/%Y %I:%M:%S %p')

logging.info('Script Initiated at: ' + logStartTime)
mintPath = 'BalanceCopy.xls'

#get credentials
separator=":"
fileIN = open('mintCreds', "r")
line = fileIN.readline()

while line:
    sout=line.split(separator)
    user=sout[0]
    passwd=sout[1]
    line = fileIN.readline()

#get array
fileIN = open('mintArray', "r")
line = fileIN.readline()
nums = []
newBalances = []

while line:
    nums = line.split(separator)
    line = fileIN.readline()

driver = webdriver.Firefox()
driver.get('https://wwws.mint.com/login.event?task=L')

time.sleep(10)

#Find Username and Password Elements by Id
UsernameElement = driver.find_element_by_id("form-login-username")
PasswordElement = driver.find_element_by_id("form-login-password")

UsernameElement.send_keys(user)
PasswordElement.send_keys(passwd)

driver.find_element_by_id("submit").click()

time.sleep(10) #Waits 10 seconds so that the site has time to refresh for new data
#TODO: Add while loop, check for refreshing indicator, once all indicators are gone, continue.

html_source = driver.page_source

soup = BeautifulSoup.BeautifulSoup(html_source) 
balances = soup.findAll('span', attrs={'class' : 'balance'})
#print balances

#for balance in balances:
#    print balance.text

for i in range(len(nums)):
    index = int(nums[i])
    balances[index] = balances[index].text
    balances[index] = re.sub('[-]', '', balances[index])
    balances[index] = re.sub('[$,]', '', balances[index])
    logging.info(balances[index])
    balances[index] = float(balances[index])
    newBalances.append(balances[index])

#Start Excel
rb = open_workbook(mintPath,formatting_info=True)
r_sheet = rb.sheet_by_index(0)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

num_rows = r_sheet.nrows
strRows = str(num_rows + 1)
StrRowsMinOne = str(num_rows)

logging.info(num_rows)
logging.info(time.strftime("%m/%d/%Y"))

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
w_sheet.write(num_rows,9,Formula('sum(B' + strRows + ':H' + strRows + ')'),currencyStyle)
#Asset Change
w_sheet.write(num_rows,10,Formula('J' + strRows + '-J' + StrRowsMinOne),currencyStyle)
#Total Liabilities
w_sheet.write(num_rows,11,newBalances[len(newBalances) - 1],currencyStyle)

wb.save(mintPath)

logEndTime = time.strftime('%m/%d/%Y %I:%M:%S %p')
logging.info('Script ended at: ' + logEndTime)

#Close Down
driver.close()

print('Done')
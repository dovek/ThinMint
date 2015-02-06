# ThinMint
Python Script that pulls information from Mint.com

This Python script will open a friefox browser window and download financial 
information to be stored in Excel. Here's how to install it.

1. Install Python Using Python 2.7)
2. Dependencies include the following:
    -Beautiful Soup
    -Selenium
    -xlUtils

When mint.py runs it will use the credentials stored in the mintCreds file to
log into mint.com.

Once logged in it will use beautiful soup to put the financial information into
an array. Then it will use the mintArray file to select which elements in the 
array are to be written to the excel file 'BalanceCopy.xls'.
	
One nice feature about this is that if you are lazy like me but you want to 
keep financial records, Windows has a task scheduler that can run this for you. 
I have my script run at the beginning of every month.



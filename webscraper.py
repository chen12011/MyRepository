from pickle import TRUE
from sre_constants import SUCCESS
import requests
from bs4 import BeautifulSoup
import re
import smtplib
from urllib.request import urlretrieve
import PyPDF2 
import urllib3

"""
A script created for a friend who needed to be updated when a specific courtcase popped up in a judges daily schedule. These schedules
are updated daily and are uploaded through PDFs.

This script scrapes a specific court county page looking for a specific string within the PDF files.
Hosted within an ec2 instance and runs hourly.

Removed private data
"""

appPass = "Google app Password"
gmail_user = 'gmail user that will be the sender'

recipient = ["emails"]


fp = requests.get('https://www.cobbcounty.org/courts/state-court/clerk/civil-calendars')
soup = BeautifulSoup(fp.content,features="html.parser")
regExp = re.compile('amazonaws')
num = 0
findString = "string to look for"
regexp1 = re.compile(findString,re.IGNORECASE)
list = []
listoffiles = []
for a in soup.find_all('a', href=re.compile('s3.us-west-2.amazonaws.com')):
    link = requests.get(a['href'])
    filename = 'file'+str(num)
    filepath = 'local filepath within ec2 instance or machine'
    urlretrieve(a['href'],filepath)

    pdfFileObj = open(filepath, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 

    pageObj = pdfReader.getPage(0) 

    pdfText = (pageObj.extractText()) 
    if regexp1.search(pdfText):
        list.append("Success")
        listoffiles.append(a.contents)
        print(SUCCESS)
    else:
        print("FAIL")
    
    pdfFileObj.close() 
    num += 1


email_text_success = "Successful, I have found '" + findString + "' in these files" + str(listoffiles)

if list:
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, appPass)
        server.sendmail(gmail_user, recipient, email_text_success)
        server.close()
    except:
        print('Something went wrong...') 
else:
    print("Nothing")  
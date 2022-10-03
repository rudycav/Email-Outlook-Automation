import win32com.client
from bs4 import BeautifulSoup
import re
import pandas as pd
import xlsxwriter
import emoji



def outlook_data(keywords, row=0, col=0): 
    workbook = xlsxwriter.Workbook('outlookdata.xlsx') #open excel
    sheet = workbook.add_worksheet()

    #connect to outlook inbox
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    
    #create lists for subject and body
    subject_list = [message.Subject for message in messages]
    body_list = [message.Body for message in messages]

    keywords = keywords #look for certain keywords in email subject

    #combine lists into dictionary
    messages_dict = dict(zip(subject_list,body_list))

    #start from cell 0 for row & columns
    row = row 
    col = col

    for key, val in messages_dict.items():
        if keywords in key:
            links = re.findall(r'(https?://\S+)', str(val))[1] #get the first link from each email
            cleanr = re.compile('<.*?>') #compiler for span
            get_span_tag = re.sub(cleanr, '', str(val)) #get span tag for each email
            get_a_tag = re.findall(r'>(.*?)<', str(val)) #get a tag for each email

            sheet.write(row, col, key)
            sheet.write(row, col+1, links)
            row += 1
    workbook.close() #close workbook




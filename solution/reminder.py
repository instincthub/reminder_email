#! python3
# Send Email reminder based on their status on the spreadsheet

import openpyxl
import smtplib,ssl
import sys

# Open the spreadsheet and get the latest dues status.
wb = openpyxl.load_workbook('../duesRecords.xlsx')
sheet = wb['Sheet1']

lastCol = sheet.max_column
latestMonth = sheet.cell(row=1, column=lastCol).value

unpaidMembers = {}
# Check each member's payment status
for r in range(2, sheet.max_row + 1):
    payment = sheet.cell(row=r, column=lastCol).value
    if payment != 'paid':
        name = sheet.cell(row=r, column=1).value
        email = sheet.cell(row=r, column=2).value
        unpaidMembers[name] = email
        # print(name, email, unpaidMembers)

# Create a secure SSL context
domain = 'smtp.gmail.com'
port = 587
input_pw = input('Email Password:')

# Try to log in to server and send email
try:
    server = smtplib.SMTP(domain, port)
    server.ehlo()
    server.starttls()
    server.login('skills@instincthub.com', input_pw)

    # Send out reminder emails.
    for name, email in unpaidMembers.items():
        body = 'Subject: %s dues unpaid.\nDear %s,\nRecords show that you have not paid dues for %s. Please make this payment as soon as possible. Thank you!' % (latestMonth, name, latestMonth)
        print('Sending email to %s...' % email)
        sendmailStatus = server.sendmail('skills@instincthub.com', email, body)

        if sendmailStatus != {}:
            print('There was a problem sending email to %s: %s' % (email, sendmailStatus))
except Exception as e:
    # Print any error messages to stdout
    print(e)
finally:
    server.quit()

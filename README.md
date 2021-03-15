# Sending Reminder Emails to Selected Members in an Excel Spreadsheet Using Python.

Say you have been “volunteered” to track member dues for the Mandatory Volunteerism Club. This is a truly boring job, involving maintaining a spreadsheet of everyone who has paid each month and emailing reminders to those who haven’t. Instead of going through the spreadsheet yourself and copying and pasting the same email to everyone who is behind on dues, let’s— you guessed it—write a script that does this for you.

At a high level, here’s what your program will do:
>- Read data from an Excel spreadsheet.
>- Find all members who have not paid dues for the latest month. Find their email addresses and send them personalized reminders.

##This means your code will need to do the following:
>- Open and read the cells of an Excel document with the openpyxl module.
>- Create a dictionary of members who are behind on their dues.
>- Log in to an SMTP server by calling smtplib.SMTP(), ehlo(), starttls(), and login().
>- For all members behind on their dues, send a personalized reminder email by calling the sendmail() method.
>- Create a new folder on your desktop called 'reminder_email'. Open the folder with your referred editor and create a new file called reminders.py.
---
####Reference: 
Al Sweigart, 2015. Automate the Boring Stuff with Python: Practical Programming for Total Beginners
# Sending Reminder Emails to Unconfirmed Attendance 

- In the lesson, you learned about how to use python openpysl module to access data from excel spreadsheet by targeting the rows and column specific to the mail we are trying to send.

- Also, you learned about how to send email with the python smtplib module which allows us to login our mail server and send a custom message to the un paid members.

## Similarly, you'll replicate what you've learnt in this project. 

>#### TODO: Check each for each participant with the red cell.
> >#### TODO: Store list in a dictionary

>#### TODO: Log in to email account.

>#### TODO: Send out reminder emails.


## Part 1: Access the Excel Sheet
1. Import the appropriate modules. 
2. Load the ```participantsSheet.xlsx``` workbook and store in wb variable
3. Store the first sheet in the ```'sheet'``` variable. 
4. Target the second to the last column and store in a variable like this ``'status_col = sheet['C3']'``.
5. Target the event title in the first row and first column and save the value in a variable like this ``event_title = sheet['A1']'event_title = sheet['A1']``. 
6 Create an empty dictionary called ``unconfirmed = {}``
   
## Part 2: Loop through the spreadsheet
1. Check for participants whose second to the last column contain red cells. Remember you need to run a forloop here ``for r in range(2, sheet.max_row + 1):``.
2. For each row, store the value in a variable called 'participant' (print this variable to make sure you are getting the right value) ``participant = sheet.cell(row=r, column=lastCol).value``. 
3. Still in the loop, create a variable called cell_color and store ``status_col[r].fill.start_color.index`` in it. 
4. Check if the cell contains red (``'red'`` is the same as ``'FFFF0000'`` in hexadecimal value). Check if the variable ``cell_color == 'FFFF0000'``: 
   >store ``name = sheet.cell(row=r, column=1).value``
   
   >store ``email = sheet.cell(row=r, column=2).value``

5. Store the name and email variable in the empty ``unconfirmed`` dictionary you created earlier. 

## Part 3: Login and Send Reminder Email
1. Log into your account:
   ```python
   server = smtplib.SMTP('smtp.gmail.com', 587) server.ehlo()
   server.starttls()
   server.login(' my_email_address@gmail.com ', 'email_password')
   ```
2. Loop through the ``unconfirmed`` dictionary; ``for name, email in unconfirmed.items():``
3. Use ``f''`` string to create your email ``body`` variable and use the message below:
    > Subject: {event_title} Reminder Email.\nDear {name}, \nRecords show that you are yet to confirm your invitation for {event_title} coming up in August. Kindly confrim take allow us to proceed with other arrangements. \n\nThank you 
4. Send your email ``server.sendmail('my_email_address@gmail.com', email, body)`` you can store  this in a ``send_email`` variable, so you can use it to print error message in case email fail to send.
5. ``if send_email != {}`` print the faulty email.


***If you get stuck, google it, if you still can't figure it out, email us at ``skills@instincthub.com``***



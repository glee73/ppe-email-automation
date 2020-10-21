# PPE Email Automation

### Introduction

I created this program when I was volunteering for a nonprofit that
distributed PPE to healthcare workers working with COVID-19 patients.
Since the nonprofit was a startup, there wasn't much infrastructure in
place, and all of the donor-recipient match emails had to be copied and 
pasted by hand. The process was a little tedious, so I decided to save 
some time by automating the email-sending. I intended to distribute this 
as an executable to my team as well, to use until the technology team 
could develop a more robust platform for us. This was my first personal
project!

### Technologies 

Pandas, Numpy, HTML

### Overview

*Logging into email & reading from Excel*

At the start of the script, I coded a few lines to let the user input their
email log-in info, name, and match sheet name. Next, I read in the match sheet
as a Pandas dataframe, which I cleaned to contain only the columns that I
needed.

*Iterating through the matches*

I knew I had to find an efficient way to get through all the donors, so I
decided to turn the email column into a NumPy array. I created a helper function 
to delete the donor emails as I iterated through the array with a while loop
until it was empty. To see if a match was single- or multiple-recipient, I
checked to see if the first item of the array was equal to the second.

*Sending the email*

I created separate helper functions to deal with single- and multiple-recipient
matches and applied them to the chunk of the dataframe with the correct donor
email, which I deep copied using conditional selection. For the most part, I
extracted the necessary data from the chunk and placed them into the email with
f-strings. I coded the email with HTML - written as a string - placed it into
a MIMEMultipart alternative message container, and sent it with SMTPlib.

To account for cases where the recipient email was missing or dysfunctional,
I used a try/except block and added the match to a list where I kept track of
the emails that failed to send.

For the tables in the multi-recipient emails, I created a helper function that
took the relevant columns from the chunk, stylized it, and converted it to
HTML so that it could be placed as an f-string into the email as a table.

Once I finished sending all the emails, I had the script send the user an
email informing them of the sent and unsent emails.

### Instructions for use

Place the match sheet and the Python script in the same folder. Navigate
to the correct directory using the command line and type "python automated.py".
Important notes - the emails in the match sheet must be sorted in alphabetical
order & the user must have Pandas, NumPy, and Python installed (unless the user
is using an .exe version)

### Bugs/shortcomings

This script assumes that all donors have a valid email listed because that was
always the case. I still tried to implement safeguards, but I found it difficult
to do so - if the email is NaN, I cannot place the donor email into the not-sent 
list, nor can I extract any other information because I cannot access the row
index. I tried one approach where I pre-emptively placed all rows with
NaN donor or recipient emails into the not-sent list and removed them from the
dataframe, but that also messed up the row indices. I tried to use .item()
intead of .loc() and it worked on Jupyter but did not work when I ran my tests
on the command line. Also, because I could not use real emails to test, I could
not catch all of the SMTPlib errors.

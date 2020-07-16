# ppe-emails

INTRODUCTION:

I created this program when I was volunteering for a nonprofit that
distributed PPE to healthcare workers working with COVID-19 patients.
Each week, I would receive an Excel sheet of donor-recipient matches, spit
straight out of the nonprofit's matching algorithm, and I would connect
the donors and recipients using a couple of email templates. There were two
kinds of matches: one donor to one recipient and one donor to multiple
recipients. The latter kind was especially frustrating because I had to copy
and paste all the recipients' information into a table - I had a case
where one donor was donating to 19 recipients - so it was tedious and
time-consuming. Because my first sheet took me about 6 hours, I decided to save
some time by automating it.

TECHNOLOGIES: 

Pandas, Numpy, SMTPlib

OVERVIEW:

Logging into email & reading from Excel

At the start of the script, I coded a few lines to let the user input their
email log-in info, name, and match sheet name. Next, I read in the match sheet
as a Pandas dataframe, which I cleaned to contain only the columns that I
needed.

Iterating through the matches

I knew I had to find an efficient way to get through all the donors, so I
decided to turn the email column into a NumPy array. Since multiple-recipient
matches barred me from using a for-loop - the donor's email would appear
multiple times at the front of the array, I created a helper function to delete
the donor emails accordingly and I iterated through the array with a while loop
until it was empty. To see if a match was single- or multiple-recipient, I
checked to see if the first item of the array was equal to the second.

Sending the email

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

INSTRUCTIONS FOR USE

Place the match sheet and the Python script in the same folder. Navigate
to the correct directory using the command line and type "python automated.py".
Important notes - the emails in the match sheet must be sorted in alphabetical
order & the user must have Pandas, NumPy, and Python installed (unless the user
is using an .exe version)

BUGS/SHORTCOMINGS

This script assumes that all donors have a valid email listed because that is
always the case. I still tried to implement safeguards, but I found it difficult
to do so - if the email is NaN, I cannot place the donor email into the not-sent 
list, nor can I extract any other information because I cannot access the row
index. I tried one approach where I pre-emptively placed all rows with
NaN donor or recipient emails into the not-sent list and removed them from the
dataframe, but that also messed up the row indices. I tried to use .item()
intead of .loc() and it worked on Jupyter but did not work when I ran my tests
on the command line. Also, because I could not use real emails to test, I could
not catch all of the SMTPlib errors.

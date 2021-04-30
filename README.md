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

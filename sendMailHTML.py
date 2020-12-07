#!/usr/bin/env python

import datetime
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from os.path import join, dirname
from dotenv import load_dotenv

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

# me == my email address
# you == recipient's email address
me = os.getenv("from", "yourname@yourserver.com")
#you = ["xxx@homecreditcfc.cn"]
you = ["sucre.xu@homecredit.cn"]

# Create message container - the correct MIME type is multipart/alternative.
COMMASPACE = ', '
msg = MIMEMultipart('alternative')
msg['Subject'] = "Subject"
msg['From'] = me
msg['To'] = COMMASPACE.join(you)

# Open a plain text file for reading.  For this example, assume that
# the text file contains only ASCII characters.
fp = open("new.html", 'rb')
# Create a text/plain message
text_HTML = fp.read().decode('utf-8')
fp.close()

# Create the body of the message (a plain-text and an HTML version).
# text = "Daily report"

# Record the MIME types of both parts - text/plain and text/html.
# part1 = MIMEText(text, 'plain')
part2 = MIMEText(text_HTML, 'html')

# Attach parts into message container.
# According to RFC 2046, the last part of a multipart message, in this case
# the HTML message, is best and preferred.
# msg.attach(part1)
msg.attach(part2)

# Send the message via local SMTP server
host = os.getenv('host', "smtp.example.com")
port = os.getenv('port', "25")
s = smtplib.SMTP(host=host, port=port)
# sendmail function takes 3 arguments: sender's address, recipient's address
# and message to send - here it is sent as one string.
s.sendmail(me, you, msg.as_string())
print("Success")
print(os.getenv("from", "yourname@yoursever.com"))
s.quit()

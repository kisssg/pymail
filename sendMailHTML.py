#  update at 2021-01-19 17:21 by sucre
#  show sending progress bar, update at 2020-12-15 14:40 by sucre
import time
print("Reading data from template...",time.strftime("%Y-%m-%d %H:%M:%S"))
import smtplib
import xlrd
import os
import pandas as pd
from progress.bar import Bar

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import join, dirname
from dotenv import load_dotenv

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

# me == my email address
# you == recipient's email address
me = os.getenv("from", "yourname@yourserver.com")
#you = ["xxx@homecreditcfc.cn"]
try:
    book = xlrd.open_workbook(dirname(__file__)+'/template.xlsx')
    sheet = book.sheet_by_name('Sheet1')
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)]
            for r in range(sheet.nrows)]
    data = pd.DataFrame(data, None, data[0])
    book.__exit__
    print("File read completed.",time.strftime("%Y-%m-%d %H:%M:%S"))
except FileNotFoundError:
    print("Template file not found.")
    os._exit(0)
except:
    print("Error reading template file.")
    os._exit(0)
    
result = {}

try:    
    print("Sending emails...",time.strftime("%Y-%m-%d %H:%M:%S"))
    bar=Bar('Process：',max=len(data)-1)
    for i in range(1, len(data)):
        # print(['name','email','msg1','msg2'])
        you = data['mail_to'][i].split(',')

        # Create message container - the correct MIME type is multipart/alternative.
        COMMASPACE = ', '
        msg = MIMEMultipart('alternative')
        # Subject of email
        msg['Subject'] = "Notification {}-{}".format(
            data['contract'][i], xlrd.xldate.xldate_as_datetime(data['date'][i], 0).date())
        msg['From'] = me
        msg['To'] = COMMASPACE.join(you)
        # msg['Cc'] = me
        # Open a plain text file for reading.  For this example, assume that
        # the text file contains only ASCII characters.
        # fp = open("new.html", 'rb')
        # Create a text/plain message
        # text_HTML = fp.read().decode('utf-8')
        # fp.close()
        template = """
        <body>
        Dear {0},<br/>
        <br/>
        Your ID number is {1}.
        </body>
        """
        text_HTML = template.format(data['name'][i], data['id'][i])

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
        host = os.getenv('host')
        port = os.getenv('port')
        s = smtplib.SMTP(host=host, port=port)
        # sendmail function takes 3 arguments: sender's address, recipient's address
        # and message to send - here it is sent as one string.
        sendresult = s.sendmail(me, you, msg.as_string())
        result = {**result, **sendresult}
        bar.next()
    bar.finish()
    if(result == {}):
        print("Mails sent！ ",time.strftime("%Y-%m-%d %H:%M:%S"))
    else:
        for r in result:
            print(r)    
    s.quit()
except:
    print("Something went wrong, check your configuration please.",time.strftime("%Y-%m-%d %H:%M:%S"))

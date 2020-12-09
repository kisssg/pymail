#!/usr/bin/env python

import datetime
import smtplib
import xlrd
import os
import pandas as pd

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
book = xlrd.open_workbook('template.xlsx')
sheet = book.sheet_by_name('Sheet1')
data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
data=pd.DataFrame(data,None,data[0])
book.__exit__
result={}
for i in range(1,len(data)):
    #print(['name','email','msg1','msg2'])
    you = data['接收人邮箱'][i].split(',')

    # Create message container - the correct MIME type is multipart/alternative.
    COMMASPACE = ', '
    msg = MIMEMultipart('alternative')
    # 这里是输入邮件标题
    msg['Subject'] = "音视频核查结果通知{}-{}".format(data['合同号码'][i],xlrd.xldate.xldate_as_datetime(data['外访日期'][i],0).date())
    msg['From'] = me
    msg['To'] = COMMASPACE.join(you)
    # msg['Cc'] = me
    # Open a plain text file for reading.  For this example, assume that
    # the text file contains only ASCII characters.
    # fp = open("new.html", 'rb')
    # Create a text/plain message
    # text_HTML = fp.read().decode('utf-8')
    # fp.close()
    template="""
    <body>
    Dear {0},<br/>
    <br/>
    <table cellpadding="5" cellspacing="0" border="1"> 
    <tr><td>员工工号</td><td>{1}</td></tr>
    <tr><td>质检工号</td><td>{2}</td></tr>
    <tr><td>合同号码</td><td>{3}</td></tr>
    <tr><td>外访日期</td><td>{4}</td></tr>
    <tr><td>外访时间</td><td>{5}</td></tr>
    <tr><td>无效外访申诉结果</td><td>{6}</td></tr>
    <tr><td>此音视频中的错误</td><td>{7}</td></tr>
    <tr><td>说明</td><td>核查出的违规将提出扣罚建议。</td></tr>
    <tr><td>正确指引提示</td><td>请参考《错误行为正确做法指引V2.0》</td></tr>
    <tr><td>如果您对此错误有疑问</td><td>如果您对此错误有疑问，您可以联系您的组长协助查看视频。</td></tr>
    <tr><td>如果您需要申诉</td><td>如果您需要申诉，您需要记录视频中关键行为的时间点，按照申诉邮件模板回复进行申诉。 </td></tr>
    </table><br/>
    Late Collection Quality Control Team<br/>
    后期催收质检团队
    </body>
    """
    text_HTML=template.format(data['员工姓名'][i],data['员工工号'][i],data['QC工号'][i],data['合同号码'][i],xlrd.xldate.xldate_as_datetime(data['外访日期'][i],0).date(),data['外访时间'][i],data['无效外访申诉结果'][i],data['此音视频中的错误'][i])

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
    sendresult=s.sendmail(me, you, msg.as_string())
    result = {**result, **sendresult}  
if(result=={}):
    print("All done successfully!")
else:
    for r in result:
        print(r)
s.quit()

from exchangelib import DELEGATE, IMPERSONATION, Account, Credentials, \
Configuration, NTLM, CalendarItem, Message, \
Mailbox, Attendee, Q, ExtendedProperty, FileAttachment, ItemAttachment, \
HTMLBody, Build, Version
from exchangelib import EWSTimeZone, EWSDateTime, EWSDate
import requests
from datetime import timedelta
import datetime as dt
import re, os, json
import keyring
import traceback
import logging
import time
import getpass
"""
Setting up configuration of exchange server
"""
#serviceId = jsonData["SERVICEID"]
userName = <firstname.lastname@company.com> # or getpass.getuser()
#userPassword = keyring.get_password(serviceId, userName)
userPassword = <Password>
credentials = Credentials(userName, userPassword)
#smtpAddress = jsonData['AccountInfo']['SMTPaddress']
smtpAddress = <mailbox-emailid>
#Server = jsonData['Server']
Server = "outlook.office365.com"

config = Configuration(server=Server, credentials=credentials)
account = Account(primary_smtp_address=smtpAddress, config=config, credentials=credentials, autodiscover=False,
                  access_type=DELEGATE)
account.root.refresh();
inboxItems = account.inbox.all();
print(inboxItems)
unreadcount = account.inbox.unread_count
print(unreadcount)
"""
Sets the timerange within which the mails are extracted
"""
tz=EWSTimeZone.localzone()
two_hours =tz.localize(EWSDateTime.now()) - timedelta(hours=3)
recent_mails=account.inbox.filter(datetime_received__range=(two_hours,tz.localize(EWSDateTime.now()))).order_by('-datetime_received')

print (recent_mails.count())
for msg in recent_mails:
    #StartDate = dt.datetime.now().strftime("%m/%d/%Y %H:%M:%S")
    msgSubject = str(msg.subject)
    msgBody = str(msg.body)
    sub = "Task Distribution"
    if re.search(sub, msgSubject, re.IGNORECASE):
        for attachment in msg.attachments:
            if isinstance(attachment, FileAttachment):
                local_path = os.path.join(os.getcwd(), attachment.name)
                with open(local_path, 'wb') as f:
                    f.write(attachment.content)
                print('Saved attachment to', local_path)
        print("Found")
        break
    else:
        print("Not Found")
        
        
        
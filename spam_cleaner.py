import imaplib
import email
import os
import re
import datetime

date_time = datetime.datetime.now()
date_time_formatted = date_time.strftime('%a %b %d %Y %X')

M = imaplib.IMAP4_SSL('outlook.office365.com')

username = os.environ['OUTLOOK_EMAIL']
password = os.environ['OUTLOOK_SMTP_PASS']

M.login(username, password)

type, data = M.list()

pattern = r'[(\w+|\W+)\s*]+@postmaster\d*.\w+.\w+'


def delete_spam():

    M.select('Inbox')

    typ, msgids = M.search(None, 'SINCE "01-APR-2023"')

    for msgid in msgids[0].split():

        _, data = M.fetch(msgid, '(RFC822)')

        message = email.message_from_bytes(data[0][1])

        if re.search(pattern, str(message.get('From'))):
            print('Deleting.....')
            print(f'From: {message.get("From")}')
            print(f'Subject: {message.get("Subject")}')
            print(f'Deleted On: {date_time_formatted}\n')
            print('----------------------\n')

            M.store(msgid, '+FLAGS', '\\Deleted')

    M.expunge()
    M.close()
    M.logout()


delete_spam()

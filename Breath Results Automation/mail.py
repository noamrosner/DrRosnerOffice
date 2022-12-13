import smtplib
import ssl
from email.message import EmailMessage
from tkinter import filedialog
import imaplib, email, getpass
import os, sys
from bs4 import BeautifulSoup
import warnings
warnings.filterwarnings("ignore")



# info
email_sender = 'SENDEREMAILADDRESS'
email_password = 'PASSWORD'
attachment_dir = "/Path/to/dir/../"
# attachment_dir = filedialog.askdirectory()

def login():
    global con
    email_sender = 'SENDEREMAILADDRESS'
    email_password = 'PASSWORD'
    imap_url = "imap.gmail.com"
    con = imaplib.IMAP4_SSL(imap_url)
    print("Trying to log in")
    con.login(email_sender, email_password)
    print("Login was successful")
    con.select('INBOX')
    print("Inside INBOX")


def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)


def search(key, value, con):
    result, data = con.search(None, key, '"{}"'.format(value))
    return data


def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs


def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()

        # check if there is a file
        if bool(fileName):
            print("downloading the attachment")
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath, 'wb') as f:
                f.write(part.get_payload(decode=True))


def scan():
    global con
    login()
    msgs = get_emails(search('FROM', 'Noam Rosner' , con))
    for msg in msgs:
        raw = email.message_from_bytes(msg[0][1])
        get_attachments(raw)



if __name__ == "__main__":
    scan()

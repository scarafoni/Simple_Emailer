# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

try:
    import win32com.client
except:
    print('warning, no outlook capabilities supported')

import pandas as pd
import argparse
import time
import smtplib
import tqdm
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.encoders import encode_base64
from email.message import EmailMessage
import  os
import traceback

parser = argparse.ArgumentParser()
parser.add_argument('email_csv')
parser.add_argument('text')
parser.add_argument('--ignore', default='na')
parser.add_argument('--debug', action='store_true')
parser.add_argument('--service', default='namecheap', choices=['outlook', 'gmail', 'namecheap'])
parser.add_argument('--attachment', default='none')


def add_attachment(f, message): 
    with open(f, "rb") as fil:
        part = MIMEApplication(
            fil.read(),
            Name=os.path.basename(f)
        )
    # After the file is closed
    part['Content-Disposition'] = 'attachment; filename="%s"' % os.path.basename(f)
    message.attach(part)

    return message

def send_email_namecheap(email, subject, body, args):
    sender_email = 'dan@scarafoni.com'
    receiver_email  = email
    smtp_server = 'mail.privateemail.com'
    port = 465
    login = "dan@scarafoni.com"
    password = open('pw_namecheap.txt','r').read().strip()
    # message = EmailMessage()
    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = f"Dan Scarafoni <{sender_email}>"
    message["To"] = receiver_email
    content = body
    # print(content)
    # message.set_content(content)
    message.attach(MIMEText(content, 'plain'))
    if args.attachment  is not 'none':
        message = add_attachment(args.attachment, message)
    server = smtplib.SMTP_SSL(smtp_server, port)
    server.login(login, password)
    if args.debug:
        e = None
    else:
        try:
            server.send_message(message)
        except:
            print('ERROR- unable to send email')
            traceback.print_exc()
        e = email
    server.quit()
    return e

def send_outlook_email(email, subject, body, args):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= subject
    newmail.To=email
    newmail.Body= body
    # print(f'\t to: {email}')
    # print(f'\t subject: {subject}')
    # print(f'\t {body}')
    if args.debug:
        return None
    else:
        newmail.Send()
        return email

def send_gmail(email, subject, body, args):
    def send_email(subject, body, sender, recipients, password):
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = ', '.join(recipients)
        smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, recipients, msg.as_string())
        smtp_server.quit()
    sender = "dscarafo@u.rochester.edu"
    recipients = [email]
    password = open('pw.txt','r').read().strip()
    # print(f'\t to: {email}')
    # print(f'\t subject: {subject}')
    # print(f'\t {body}')
    if args.debug:
        return None
    else:
        try:
            send_email(subject, body, sender, recipients, password)
        except:
            print('ERROR- unable to send email')
        return email

def main(args):
    df = [str(x) for x in pd.read_csv(args.email_csv)['Email'].tolist()]
    text = open(args.text, 'r').read().split('\n')
    if args.ignore == 'na':
        ignore = []
    else:
        ignore = [str(x) for x in pd.read_csv(args.ignore)['Email'].tolist()]
    
    ignore_add = []
    subject = text[0]
    body = '\n'.join(text[1:])
    pbar = tqdm.tqdm(df)
    for email in pbar:
        if email == 'nan':
            m = f'found nan email- {email}, continuing...'
            continue
        elif email in ignore:
            m = f'ignoring {email} as it\'s already been sent'
            continue
        else:
            m = f'email to send- {email}'
        if args.debug:
            d = 'debug...not sending'
        else:
            d = ''
        pbar.set_description(f'{m}, {d}')

        if args.service == 'outlook':
            e = send_outlook_email(email, subject, body, args)
        elif args.service == 'gmail':
            e = send_gmail(email, subject, body, args)
        else:
            e = send_email_namecheap(email, subject, body, args)
        if not e is None:
            ignore_add.append(e)
        
        if not args.ignore == 'na' and not args.debug:
            new_df = pd.DataFrame(ignore+ignore_add, columns=['Email'])
            new_df.to_csv('ignore.csv', index=False)

        t = 1 if args.debug else 10
        time.sleep(t)

if __name__ == '__main__':
    args = parser.parse_args()
    main(args)
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

parser = argparse.ArgumentParser()
parser.add_argument('email_csv')
parser.add_argument('text')
parser.add_argument('--ignore', default='na')
parser.add_argument('--debug', action='store_true')
parser.add_argument('--service', default='outlook', choices=['outlook', 'gmail'])

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
        else:
            e = send_gmail(email, subject, body, args)
        if not e is None:
            ignore_add.append(e)
        
        time.sleep(3)
        
    if not args.ignore == 'na' and not args.debug:
        print('updatng ignore email list')
        new_df = pd.DataFrame(ignore+ignore_add, columns=['Email'])
        new_df.to_csv('ignore.csv', index=False)


if __name__ == '__main__':
    args = parser.parse_args()
    main(args)
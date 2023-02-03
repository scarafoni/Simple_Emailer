# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import win32com.client
import pandas as pd
import argparse
import time

parser = argparse.ArgumentParser()
parser.add_argument('email_csv')
parser.add_argument('text')

def main(args):
    df = [str(x) for x in pd.read_csv(args.email_csv)['Email'].tolist()]
    text = open(args.text, 'r').read().split('\n')
    subject = text[0]
    body = '\n'.join(text[1:])
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    for email in df:
        if email == 'nan':
            print(f'found nan email[ {email}, continuing...')
            continue
        print('sending email:')
        newmail=ol.CreateItem(olmailitem)
        newmail.Subject= subject
        newmail.To=email
        newmail.Body= body
        print(f'\t to: {email}')
        print(f'\t subject: {subject}')
        print(f'\t {body}')
        # attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
        # newmail.Attachments.Add(attach)
        # To display the mail before sending it
        # newmail.Display() 
        newmail.Send()
        time.sleep(1)
        print()

if __name__ == '__main__':
    args = parser.parse_args()
    main(args)
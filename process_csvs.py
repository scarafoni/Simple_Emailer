"""
processes csvs for gmail and outlook processing
"""
import os
import glob
import pandas as pd

def main():
    # get all csv files and combine into one file
    csvs = glob.glob('input_csvs/*.csv')
    dfs = [pd.read_csv(x) for x in csvs]
    emails = []
    for df in dfs:
        if 'email' in df.columns:
            emails.extend(df['email'].tolist())
        else:
            emails.extend(df['Email'].tolist())
    # first 300 entries are for outlook
    outlook_emails = pd.DataFrame(emails[:250], columns=['Email'])
    outlook_emails.to_csv('outlook_input.csv', index=False)
    
    gmail_emails = pd.DataFrame(emails[250:1950+250], columns=['Email'])
    gmail_emails.to_csv('gmail_input.csv', index=False)

    leftover_emails = pd.DataFrame(emails[1950+250:], columns=['Email'])
    leftover_emails.to_csv('leftover_input.csv', index=False)

if __name__ == '__main__':
    main()
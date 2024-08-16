"""
The vendor I most frequently order books from is GOBI. When I do this, I receive an acknowledgement email.
I wrote this script to extract some of the key information in these acknowledgement emails,
and put it in a spreadsheet so that I didn't have to do it manually.
"""

import win32com.client
import re
import csv
import os
import subprocess

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # The inbox is folder 6
messages = inbox.Items

csv_file = 'books_ordered.csv'
csv_columns = ['Date', 'Title', 'Author', 'Binding', 'Series', 'Supplier', 'Purchase Option', 'Price', 'Fund']
log_file = 'processed_emails.log'

# I'm keeping a log of emails that I have already processed, and only processing the new
# ones each time I run the program.
if os.path.exists(log_file):
    with open(log_file, 'r') as f:
        processed_ids = set(f.read().splitlines())
else:
    processed_ids = set()

# Extracts the info from a field with the simple format "{field_name}: {value}"
def extract_info(chunk, field_name):
    try:
        # Use regex to find the field value
        pattern = re.compile(rf'{field_name}:\s*(.*?)\s*[\n\t]', re.DOTALL)
        match = pattern.search(chunk)
        if match:
            return match.group(1).strip()
        return ''
    except ValueError:
        return ''
    
# For ebooks, there is a table of prices for the various purchase options, plus an indication of
# which purchase option I chose, so I have to extract the whole price table
def extract_price_table(chunk):
    price_table = {}
    table_start = re.search(r'Supplier\s+Purchase Option\s+List Price', chunk)
    if table_start:
        table_lines = chunk[table_start.end():].strip().split('\n')
        for line in table_lines:
            if not line.strip():
                continue
            columns = re.split(r'\t', line.strip())
            if len(columns) >= 3:
                supplier = columns[0].strip().replace('+', '').replace('-', '')
                purchase_option = columns[1].strip().split(' |')[0]
                price = columns[2].strip()
                if re.match(r'^\d+\.\d{2}\s+USD$', price):
                   price_table[(supplier, purchase_option)] = price.replace(' USD', '')
    return price_table

# For print books the price is more straightforwardly listed in the email
def extract_us_list_price(chunk):
    try:
        pattern = re.compile(r'US List:\s*\$?(\d+\.\d{2})\s*USD')
        match = pattern.search(chunk)
        if match:
            return match.group(1).strip()
        return ''
    except ValueError:
        return ''

def process_email_body(email_body, received_date):
    chunks = email_body.split('SELECTION ACKNOWLEDGEMENT')
    book_entries = []
    
    for chunk in chunks[1:]:  # Skip the first chunk as it is before the first selection
        price_table = extract_price_table(chunk)
        us_list_price = extract_us_list_price(chunk)
        
        supplier = extract_info(chunk, 'Supplier')
        purchase_option = extract_info(chunk, 'Purchase Option')
        price = us_list_price if us_list_price else price_table.get((supplier, purchase_option), 'N/A')
        fund = extract_info(chunk, 'Fund')
        
        book_info = {
            'Date' : received_date,
            'Title' : extract_info(chunk, 'Title'),
            'Author' : extract_info(chunk, 'Author'),
            'Binding' : extract_info(chunk, 'Binding'),
            'Series' : extract_info(chunk, 'Series Title'),
            'Supplier' : supplier,
            'Purchase Option' : purchase_option,
            'Price' : price,
            'Fund' : fund
        }

        if book_info['Title']:
            book_entries.append(book_info)
    return book_entries

def process_message(message):
    email_id = message.EntryID
    if message.Class == 43 and message.SenderEmailAddress == "DoNotReply@Ybp.com" and message.Subject == "GOBI Selection Acknowledgements alexmanchester@stanford.edu":
        received_date = message.ReceivedTime.strftime('%Y-%m-%d')
        print(f"Looking at MailItem {email_id} from DoNotReply@Ybp.com on {received_date}")
        if email_id not in processed_ids:
            print(f"MailItem {email_id} not yet processed")
            body = message.Body
            book_entries = process_email_body(body, received_date)
            print("Found book entries")
            for book_info in book_entries:
                writer.writerow(book_info)
                print("Wrote book info")
            processed_ids.add(email_id)
            with open(log_file, 'a') as f:
                f.write(email_id + '\n')
                print("Added email id to processed log")

# Open the CSV file in append mode
with open(csv_file, 'a', newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
    # Write header only if the file is empty
    if csvfile.tell() == 0:
        writer.writeheader()
    for message in messages:
        process_message(message)

#I have to close classic Outlook at the end otherwise strange things happen.
tasklist_output = subprocess.check_output('tasklist', shell=True).decode()
if 'OUTLOOK.EXE' in tasklist_output:
    os.system("taskkill /f /im outlook.exe")

print("Done.")
import imaplib
import email
from openpyxl import load_workbook

fn = 'example.xlsx'
wb = load_workbook(fn)
ws = wb['data']



imap_url = 'imap.gmail.com'
email_address = 
email_password = input('Enter your password: ')

imap = imaplib.IMAP4_SSL(imap_url)
imap.login(email_address, email_password)

imap.select('Inbox')

_, msgnums = imap.search(None, "ALL")

index = 1

for msg in msgnums[0].split():
    _, data = imap.fetch(msg, "(RFC822)")

    message = email.message_from_bytes(data[0][1])

    print(f"Message Number: {msg}")
    print(f"From: {message.get('From')}")
    print(f"To: {message.get('To')}")
    print(f"BCC: {message.get('BCC')}")
    print(f"Date: {message.get('Date')}")
    print(f"Subject: {message.get('Subject')}")

    print("Content:")
    for part in message.walk():
        if part.get_content_type() == 'text/plain':
            print(part.as_string())
            ws[f"A{index}"] = part.as_string()
            wb.save(fn)
            index += 1

imap.close()
wb.close()

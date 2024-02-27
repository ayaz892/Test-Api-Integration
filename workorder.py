import imaplib
import email
import requests
import re

# Email configuration
EMAIL_HOST = 'imap.gmail.com'
EMAIL_USERNAME = 'ayazmuhammad1300@gmail.com'
EMAIL_PASSWORD = 'gtwb kpzf hldx lbrr'
SEARCH_CRITERIA = '(FROM "ayazhasan97@gmail.com" SUBJECT "Work Order")'

# Monday.com configuration
BOARD_ID = '1848752616'
API_KEY = 'eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjMyNjA0NjU4NCwiYWFpIjoxMSwidWlkIjo1NjQxMTgzMSwiaWFkIjoiMjAyNC0wMi0yN1QxNDowMzoxMy41NzVaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjE1MjE4MDgsInJnbiI6ImFwc2UyIn0.qTA_3-46PNBnl9HfUFHMn6ZeJ4VTJCNs4PDfctae4zk'

try:
    # Connect to the email server
    mail = imaplib.IMAP4_SSL(EMAIL_HOST)
    mail.login(EMAIL_USERNAME, EMAIL_PASSWORD)
    mail.select('inbox')

    # Search for emails with specific criteria
    result, data = mail.search(None, SEARCH_CRITERIA)

    for num in data[0].split():
        result, data = mail.fetch(num, '(RFC822)')
        raw_email = data[0][1]
        email_message = email.message_from_bytes(raw_email)

        # Extract relevant information from the email
        payload = None
        for part in email_message.walk():
            if part.get_content_type() == 'text/plain':
                payload = part.get_payload(decode=True).decode()
                break

        if payload:
            work_order_number_match = re.search(r'WO\s+(\d+)', payload)
            purchase_order_number_match = re.search(r'PO\s+(\d+)', payload)

            # Extract attachment filename if needed
            attachment_filename = None
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    attachment_filename = part.get_filename()
                    break

            if work_order_number_match and purchase_order_number_match:
                work_order_number = work_order_number_match.group(1)
                purchase_order_number = purchase_order_number_match.group(1)

                # Post data to Monday.com
                url = f'https://api.monday.com/v2/boards/{BOARD_ID}/items'
                headers = {'Authorization': f'Bearer {API_KEY}', 'Content-Type': 'application/json'}
                payload = {
                    'name': f'WO {work_order_number}',
                    'column_values': [
                        {'title': 'Item', 'text': f'WO {work_order_number}'},
                        {'title': 'Person', 'text': ''},  # Fill in the appropriate person if needed
                        {'title': 'Scheduled Date', 'date': {'date': '2023-12-05'}},  # Adjust the scheduled date as needed
                        {'title': 'WO#', 'text': work_order_number},
                        {'title': 'PO#', 'text': purchase_order_number},
                        {'title': 'Attachment', 'file': attachment_filename}
                    ]
                }
                response = requests.post(url, headers=headers, json={'item': payload})

                if response.status_code == 200:
                    print('Data added successfully to Monday.com')
                else:
                    print('Error adding data to Monday.com:', response.text)

    # Logout from the email server
    mail.logout()

except Exception as e:
    print('An error occurred:', e)

import datetime
from O365 import Account, FileSystemTokenBackend

# Connette ad outlook365 e scarica la lista delle email
# da una cartella specifica
#
# Usage: python test_connect.py
#
# Connect to Microsoft 365 account
credentials = ('client_id', 'client_secret')
token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
account = Account(credentials, token_backend=token_backend)
if not account.is_authenticated:
    # Authenticate if not already authenticated
    account.authenticate(scopes=['basic', 'mailbox_all'])
    account.connection.refresh_token()

# Get the mailbox folder
mailbox = account.mailbox()
folder = mailbox.get_folder(folder_name='Sent Items')

# Get yesterday's date
yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
yesterday = yesterday.strftime('%Y-%m-%d')

# Get all emails sent yesterday
emails = folder.get_messages(query=f"sentDateTime ge {yesterday}")

# Print the subject of each email
for email in emails:
    print(email.subject)
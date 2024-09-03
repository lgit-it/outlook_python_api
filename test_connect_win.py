import win32com.client
import sqlite3
import os

# Connessione a Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Connessione al database SQLite (crea un nuovo file db se non esiste)
db_file = 'emails.db'
conn = sqlite3.connect(db_file)
cursor = conn.cursor()

# Creazione della tabella nel database per memorizzare le email
cursor.execute('''
    CREATE TABLE IF NOT EXISTS emails (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subject TEXT,
        body TEXT,
        sender_name TEXT,
        sender_email TEXT,
        received_time TEXT,
        attachments TEXT
    )
''')

# Selezione della casella di posta (0 per Inbox)
inbox = outlook.GetDefaultFolder(6)  # 6 corrisponde a "Inbox"
messages = inbox.Items

# Iterazione attraverso le email
for message in messages:
    try:
        subject = message.Subject
        body = message.Body
        sender_name = message.SenderName
        sender_email = message.SenderEmailAddress
        received_time = message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
        
        # Gestione degli allegati
        attachments = []
        for attachment in message.Attachments:
            attachments.append(attachment.FileName)
        attachments_str = ', '.join(attachments)
        
        # Inserimento dei dati nel database
        cursor.execute('''
            INSERT INTO emails (subject, body, sender_name, sender_email, received_time, attachments)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (subject, body, sender_name, sender_email, received_time, attachments_str))

    except Exception as e:
        print(f"Errore durante l'elaborazione dell'email: {e}")

# Commit e chiusura della connessione al database
conn.commit()
conn.close()

print(f"Tutte le email sono state salvate in '{os.path.abspath(db_file)}'")

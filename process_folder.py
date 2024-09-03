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
# cursor.execute('''
#     DROP TABLE IF EXISTS emails
# ''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS emails (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        folder_path TEXT,
        subject TEXT,
        body TEXT,
        sender_name TEXT,
        sender_email TEXT,
        received_time TEXT,
        attachments TEXT,
        cc TEXT,
        bcc TEXT,
        importance TEXT,
        categories TEXT,
        is_read INTEGER,
        sensitivity TEXT,
        message_class TEXT,
        internet_headers TEXT,
        conversation_id TEXT,
        entry_id text unique
    )
''')

def process_folder(folder):
    """
    Processa tutte le email in una cartella specifica e chiama la funzione ricorsiva per tutte le sottocartelle.
    """
    
    messages = folder.Items
    
    print ("Processing folder:", folder.FolderPath ,(messages.Count))   
   
    # Iterazione attraverso le email nella cartella corrente 
    for message in messages:
        #se il messaggio non è di tipo IPM.Note, non lo processa
        if not message.MessageClass.startswith("IPM.Note"):
            continue
        
        # se il messaggio esiste già nel database non lo processa
        # Verifica se il messaggio è già stato inserito nel database usando come chiave l'entryID
        cursor.execute("SELECT entry_id FROM emails WHERE entry_id = ?", (message.EntryID,))
        if cursor.fetchone():
            continue
        
        try:
            subject = message.Subject
            body = message.Body
            sender_name = "" if not hasattr(message, 'SenderName') or message.SenderName is None else message.SenderName
            sender_email = "" if not hasattr(message, 'SenderEmailAddress') or message.SenderEmailAddress is None else message.SenderEmailAddress
            received_time = "" if not hasattr(message, 'ReceivedTime') else message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')


            # Cattura dei campi CC e BCC se sono presenti
            
            #verifica la presenza di CC e BCC
            cc = "" if not hasattr(message, 'CC') or message.CC is None else message.CC
            
            bcc = "" if not hasattr(message, 'BCC') or message.BCC is None else message.BCC
                
            # Importanza (Bassa, Normale, Alta)
            importance =  "" if not hasattr(message, 'Importance') else {0: 'Low', 1: 'Normal', 2: 'High'}.get(message.Importance, 'Normal')

            # Categorie (Tag)
            categories = "" if not hasattr(message, 'Categories') else message.Categories

            # Stato di lettura
            is_read = 1 if message.UnRead == False else 0

            # Sensibilità (Normale, Personale, Privato, Confidenziale)
            sensitivity = {0: 'Normal', 1: 'Personal', 2: 'Private', 3: 'Confidential'}.get(message.Sensitivity, 'Normal')

            # Classe del messaggio (es. IPM.Note per email standard)
            message_class = message.MessageClass

            # Headers Internet (opzionale, può essere vuoto)
            internet_headers = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")

            # Identificatore del thread/conversazione
            conversation_id = message.ConversationID
            entry_id = message.EntryID

            # Gestione degli allegati
            attachments = []
            for attachment in message.Attachments:
                if attachment.Type == 1:  # 1 corrisponde a "File"
                    attachments.append(attachment.FileName)
                if attachment.Type == 5:    # 5 corrisponde a "Outlook Item"
                    attachments.append(attachment.DisplayName)
                if attachment.Type == 6:   # 6 corrisponde a "Embedded Message"
                    attachments.append(attachment.DisplayName)
                
            attachments_str = ', '.join(attachments)
            
            # Percorso della cartella
            folder_path = folder.FullFolderPath

            # Inserimento dei dati nel database
            # inserisce i dati solo se non esiste già un record con lo stesso entry_id, altrimenti lo aggiorna
            cursor.execute('''
                INSERT INTO emails (folder_path, subject, body, sender_name, sender_email, received_time, attachments, cc, bcc, importance, categories, is_read, sensitivity, message_class, internet_headers, conversation_id, entry_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(entry_id) DO UPDATE SET
                    folder_path = excluded.folder_path,
                    subject = excluded.subject,
                    body = excluded.body,
                    sender_name = excluded.sender_name,
                    sender_email = excluded.sender_email,
                    received_time = excluded.received_time,
                    attachments = excluded.attachments,
                    cc = excluded.cc,
                    bcc = excluded.bcc,
                    importance = excluded.importance,
                    categories = excluded.categories,
                    is_read = excluded.is_read,
                    sensitivity = excluded.sensitivity,
                    message_class = excluded.message_class,
                    internet_headers = excluded.internet_headers,
                    conversation_id = excluded.conversation_id
            ''', (folder_path, subject, body, sender_name, sender_email, received_time, attachments_str, cc, bcc, importance, categories, is_read, sensitivity, message_class, internet_headers, conversation_id, entry_id))

        except Exception as e:
            print(f"Errore durante l'elaborazione dell'email: {e}")

    # Ricorsione nelle sottocartelle
    for subfolder in folder.Folders:
        process_folder(subfolder)

# Avvio dell'elaborazione a partire dalla Inbox
root_folder = outlook.Folders.Item("luigi.gregori@bancafinint.com")  # Modifica con il nome corretto della tua cartella radice, se necessario
process_folder(root_folder)

# Commit e chiusura della connessione al database
conn.commit()
conn.close()

print(f"Tutte le email e le cartelle sono state salvate in '{os.path.abspath(db_file)}'")

import win32com.client
import json
import requests
import logging
import time
import os

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_sent_emails():
    if os.path.exists('sent_emails.json'):
        with open('sent_emails.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def extract_ai_answer_emails(outlook):
    root_folder = outlook.Folders.Item(1)  # Der erste Account in Outlook
    
    ai_answer_folder = None
    for folder in root_folder.Folders:
        if folder.Name == "AI-Antwort":
            ai_answer_folder = folder
            break

    if not ai_answer_folder:
        logging.warning("Der 'AI-Antwort' Ordner wurde nicht gefunden.")
        return []

    messages = ai_answer_folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sortiere nach Empfangszeit, neueste zuerst
    
    emails = []
    for message in messages:
        if not message.UnRead:
            continue  # Überspringe gelesene Nachrichten
        emails.append({
            "subject": message.Subject,
            "sender": message.SenderEmailAddress,  # E-Mail-Adresse des Absenders
            "body": message.Body[:1000],
            "entryID": message.EntryID
        })
    
    logging.info(f"{len(emails)} ungelesene E-Mails im 'AI-Antwort' Ordner gefunden.")
    return emails

def create_ollama_model(sent_emails):
    modelfile_content = f'''FROM mistral-nemo

SYSTEM """
Du bist ein E-Mail-Assistent. Deine Aufgabe ist es, basierend auf dem Kontext früherer E-Mails,
professionelle und personalisierte Antworten auf neue E-Mails zu generieren.
Berücksichtige den Ton und Stil der früheren E-Mails in deinen Antworten.
"""
'''
    
    url = 'http://localhost:11434/api/create'
    data = {
        "name": "email_assistant",
        "modelfile": modelfile_content,
        "stream": False
    }
    response = requests.post(url, json=data)
    logging.info(response.json())

def query_ollama(email, sent_emails):
    url = 'http://localhost:11434/api/generate'
    
    email_context = "\n\n".join([f"Betreff: {e['subject']}\nAn: {e['recipient']}\nNachricht: {e['body'][:200]}..." for e in sent_emails[:5]])
    
    data = {
        "model": "email_assistant",
        "prompt": f"""Basierend auf dem folgenden Kontext früherer E-Mails:

{email_context}

Generiere eine professionelle und personalisierte Antwort auf diese neue E-Mail:

Betreff: {email['subject']}
Von: {email['sender']}
Nachricht: {email['body']}

Deine Antwort sollte höflich, prägnant und auf den Punkt sein. Beachte den Ton und Kontext der ursprünglichen E-Mail sowie den Stil früherer E-Mails.

Wichtig: Erstelle die Antwort so, als würde sie von YOU@YOURMAILADRESS.com gesendet.

Antwort:""",
        "stream": False
    }
    try:
        response = requests.post(url, json=data)
        response.raise_for_status()
        result = response.json()['response']
        
        # Entferne eventuell generierte E-Mail-Header
        result = result.split("Antwort:")[-1].strip()
        
        logging.info(f"Generierte Antwort: {result[:100]}...")  # Log nur die ersten 100 Zeichen
        return result
    except requests.exceptions.RequestException as e:
        logging.error(f"Fehler bei der Anfrage an Ollama: {e}")
        return None

def create_draft(outlook, email, response):
    drafts_folder = outlook.GetDefaultFolder(16)  # 16 ist der Code für den Entwürfe-Ordner
    mail = drafts_folder.Items.Add()
    mail.Subject = f"RE: {email['subject']}"
    mail.Body = response
    mail.To = email['sender']  # Setze den Absender der Ursprungsmail als Empfänger
    mail.SentOnBehalfOfName = "YOU@YOURMAILADRESS.com"  # Setze den Absender
    mail.Save()
    logging.info(f"Entwurf für '{email['subject']}' im Entwürfe-Ordner erstellt.")

def process_emails(outlook, sent_emails):
    ai_answer_emails = extract_ai_answer_emails(outlook)
    
    for email in ai_answer_emails:
        logging.info(f"Verarbeite E-Mail: {email['subject']}")
        response = query_ollama(email, sent_emails)
        if response:
            create_draft(outlook, email, response)
            # Markiere die verarbeitete E-Mail als gelesen
            original_mail = outlook.GetItemFromID(email['entryID'])
            original_mail.UnRead = False
            original_mail.Save()
            logging.info(f"E-Mail '{email['subject']}' als gelesen markiert.")
        else:
            logging.error(f"Fehler beim Generieren der Antwort für '{email['subject']}'.")

def main():
    sent_emails = load_sent_emails()
    create_ollama_model(sent_emails)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    while True:
        try:
            process_emails(outlook, sent_emails)
            time.sleep(30)  # Warte 30 Sekunden vor der nächstem Check
        except Exception as e:
            logging.error(f"Ein Fehler ist aufgetreten: {e}")
            time.sleep(300)  # Bei einem Fehler 5 Minuten warten

if __name__ == "__main__":
    main()

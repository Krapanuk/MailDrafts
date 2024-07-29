import os
import json
import time
import logging
import requests
import faiss
import torch
from transformers import AutoModel, AutoTokenizer
import win32com.client

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Lade das vortrainierte Modell und den Tokenizer
model_name = "dunzhang/stella_en_1.5B_v5"
tokenizer = AutoTokenizer.from_pretrained(model_name)
model = AutoModel.from_pretrained(model_name)

def load_sent_emails():
    """Lädt die gesendeten E-Mails aus einer JSON-Datei."""
    print("Lade gesendete E-Mails aus 'sent_emails.json'...")
    if os.path.exists('sent_emails.json'):
        with open('sent_emails.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def extract_ai_answer_emails(outlook):
    """Extrahiert ungelesene E-Mails aus dem 'AI-Antwort' Ordner."""
    print("Extrahiere ungelesene E-Mails aus dem 'AI-Antwort' Ordner...")
    root_folder = outlook.Folders.Item(1)  # Der erste Account in Outlook
    
    ai_answer_folder = None
    for folder in root_folder.Folders:
        if folder.Name == "Bekannte Absender":
            for subfolder in folder.Folders:
                if subfolder.Name == "AI-Antwort":
                    ai_answer_folder = subfolder
                    break
            if ai_answer_folder:
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
    """Erstellt ein Ollama-Modell basierend auf den gesendeten E-Mails."""
    print("Erstelle Ollama-Modell basierend auf den gesendeten E-Mails...")
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

def load_faiss_index():
    """Lädt den FAISS-Index, wenn vorhanden."""
    print("Lade FAISS-Index, falls vorhanden...")
    if os.path.exists("emails.index"):
        return faiss.read_index("emails.index")
    return None

def create_faiss_index(sent_emails):
    """Erstellt einen neuen FAISS-Index aus den gesendeten E-Mails."""
    print("Erstelle neuen FAISS-Index aus den gesendeten E-Mails...")
    email_texts = [email['subject'] + " " + email['body'] for email in sent_emails]
    email_vectors = get_embeddings(email_texts)
    
    index = faiss.IndexFlatL2(email_vectors.shape[1])
    index.add(email_vectors)
    faiss.write_index(index, "emails.index")
    
    return index, email_vectors

def get_embeddings(texts):
    """Erstellt Embeddings für die gegebenen Texte."""
    print("Erstelle Embeddings für die Texte...")
    inputs = tokenizer(texts, return_tensors='pt', padding=True, truncation=True)
    with torch.no_grad():
        embeddings = model(**inputs).last_hidden_state.mean(dim=1).numpy()
    return embeddings

def search_emails(query, index, email_vectors, k=5):
    """Sucht nach den relevantesten E-Mails basierend auf der Abfrage."""
    print(f"Suche nach den relevantesten E-Mails für die Abfrage: {query[:50]}...")
    query_vector = get_embeddings([query])
    distances, indices = index.search(query_vector, k)
    return indices[0]

def query_ollama(email, sent_emails, index, email_vectors):
    """Stellt eine Anfrage an das Ollama-Modell, um eine Antwort auf die E-Mail zu generieren."""
    print(f"Stelle eine Anfrage an das Ollama-Modell für die E-Mail: {email['subject']}...")
    url = 'http://localhost:11434/api/generate'
    
    query = email['subject'] + " " + email['body']
    indices = search_emails(query, index, email_vectors)
    email_context = "\n\n".join([f"Betreff: {sent_emails[i]['subject']}\nNachricht: {sent_emails[i]['body'][:200]}..." for i in indices])
    
    data = {
        "model": "email_assistant",
        "prompt": f"""Basierend auf dem folgenden Kontext früherer E-Mails:

{email_context}

Generiere eine professionelle und personalisierte Antwort auf diese neue E-Mail:

Betreff: {email['subject']}
Von: {email['sender']}
Nachricht: {email['body']}

Deine Antwort sollte höflich, prägnant und auf den Punkt sein. Beachte den Ton und Kontext der ursprünglichen E-Mail sowie den Stil früherer E-Mails.

Wichtig: Erstelle die Antwort so, als würde sie von YOU@YOURMAILADRESS.net gesendet.

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
    """Erstellt einen Entwurf im Entwürfe-Ordner von Outlook."""
    print(f"Erstelle Entwurf für die E-Mail: {email['subject']}...")
    drafts_folder = outlook.GetDefaultFolder(16)  # 16 ist der Code für den Entwürfe-Ordner
    mail = drafts_folder.Items.Add()
    mail.Subject = f"RE: {email['subject']}"
    mail.Body = response
    mail.To = email['sender']  # Setze den Absender der Ursprungsmail als Empfänger
    mail.SentOnBehalfOfName = "YOU@YOURMAILADRESS.net"  # Setze den Absender
    mail.Save()
    logging.info(f"Entwurf für '{email['subject']}' im Entwürfe-Ordner erstellt.")

def process_emails(outlook, sent_emails, index, email_vectors):
    """Verarbeitet die E-Mails im 'AI-Antwort' Ordner."""
    print("Verarbeite E-Mails im 'AI-Antwort' Ordner...")
    ai_answer_emails = extract_ai_answer_emails(outlook)
    
    for email in ai_answer_emails:
        logging.info(f"Verarbeite E-Mail: {email['subject']}")
        response = query_ollama(email, sent_emails, index, email_vectors)
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
    """Hauptfunktion zum Ausführen des E-Mail-Assistenten."""
    print("Starte E-Mail-Assistenten...")
    sent_emails = load_sent_emails()
    create_ollama_model(sent_emails)

    # Lade den gespeicherten FAISS-Index und E-Mail-Vektoren
    index = load_faiss_index()
    if index is None:
        logging.info("FAISS-Index nicht gefunden, erstelle einen neuen Index.")
        index, email_vectors = create_faiss_index(sent_emails)
    else:
        email_texts = [email['subject'] + " " + email['body'] for email in sent_emails]
        email_vectors = get_embeddings(email_texts)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    while True:
        try:
            process_emails(outlook, sent_emails, index, email_vectors)
            print("Warte 10 Sekunden vor dem nächsten Check...")
            time.sleep(10)  # Warte 10 Sekunden vor der nächstem Check
        except Exception as e:
            logging.error(f"Ein Fehler ist aufgetreten: {e}")
            print("Ein Fehler ist aufgetreten. Warte 5 Minuten vor dem nächsten Versuch...")
            time.sleep(300)  # Bei einem Fehler 5 Minuten warten

if __name__ == "__main__":
    main()
import win32com.client
import json
from datetime import datetime, timedelta

def extract_sent_emails(max_emails=50, days_back=30):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    sent_folder = outlook.GetDefaultFolder(5)  # 5 ist der Code für gesendete Elemente
    messages = sent_folder.Items
    
    # Sortiere die Nachrichten nach Sendedatum in absteigender Reihenfolge
    messages.Sort("[SentOn]", True)
    
    # Setze einen Datums-Filter für die letzten 30 Tage
    date_cutoff = datetime.now() - timedelta(days=days_back)
    messages = messages.Restrict("[SentOn] >= '" + date_cutoff.strftime('%m/%d/%Y') + "'")

    emails = []
    for i, message in enumerate(messages):
        if i >= max_emails:
            break
        
        # Extrahiere den Hauptteil der E-Mail, ohne Zitate früherer Nachrichten
        body = message.Body.split("From:")[0].strip()
        
        emails.append({
            "subject": message.Subject,
            "recipient": message.To,
            "sent_on": message.SentOn.strftime("%Y-%m-%d %H:%M:%S"),
            "body": body[:1000]  # Begrenzen Sie den body auf 1000 Zeichen
        })
    
    filename = 'sent_emails.json'
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(emails, f, ensure_ascii=False, indent=4)

    print(f"{len(emails)} gesendete E-Mails wurden extrahiert und in '{filename}' gespeichert.")

if __name__ == "__main__":
    max_emails = int(input("Wie viele E-Mails möchten Sie maximal extrahieren? "))
    days_back = int(input("Aus wie vielen Tagen in die Vergangenheit möchten Sie E-Mails extrahieren? "))
    extract_sent_emails(max_emails, days_back)

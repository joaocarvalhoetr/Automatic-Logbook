import os
import base64
import logging
import sys
from email import message_from_bytes
from email.header import decode_header
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Load environment variables
load_dotenv()

# Define the scope for Gmail modification
SCOPES = ['https://mail.google.com/']

def connect_to_gmail():
    creds = None
    # Determine base path
    if getattr(sys, 'frozen', False):
        # Running in a bundle
        base_path = sys._MEIPASS
    else:
        # Running in normal Python environment
        base_path = os.path.abspath(".")
    
    credential_path = os.path.join(base_path, 'credentials.json')
    token_path = os.path.join(base_path, 'token.json')

    logging.info(f"Credential path: {credential_path}")
    
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credential_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w') as token:
            token.write(creds.to_json())

    return creds

def fetch_emails(service, folder="INBOX"):
    # Use the Gmail API to fetch emails from the specified folder
    results = service.users().messages().list(userId='me', labelIds=[folder]).execute()
    return results.get('messages', [])

def get_email_subject(service, email_id):
    # Fetch the email by ID
    msg = service.users().messages().get(userId='me', id=email_id['id'], format='raw').execute()
    msg_str = base64.urlsafe_b64decode(msg['raw'].encode('ASCII'))
    mime_msg = message_from_bytes(msg_str)
    subject, encoding = decode_header(mime_msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding if encoding else "utf-8")
    return subject

def get_email_body(service, email_id):
    """Fetch the body of an email by ID."""
    msg = service.users().messages().get(userId='me', id=email_id, format='full').execute()
    payload = msg['payload']
    parts = payload.get('parts')
    body = ""

    if parts:
        for part in parts:
            if part['mimeType'] == 'text/plain':
                data = part['body'].get('data')
                if data:
                    body += base64.urlsafe_b64decode(data).decode('utf-8')
                    break  # Stop if plain text part is found
        else:
            for part in parts:
                if part['mimeType'] == 'text/html':
                    data = part['body'].get('data')
                    if data:
                        body += base64.urlsafe_b64decode(data).decode('utf-8')
                        break  # Stop if HTML part is found
    else:
        data = payload['body'].get('data')
        if data:
            body += base64.urlsafe_b64decode(data).decode('utf-8')

    return body

def print_email_subjects(service, email_ids):
    for email_id in email_ids:
        subject = get_email_subject(service, email_id)
        print(f"Subject: {subject}")

def filter_logbook_emails(service, email_ids):
    logbook_emails = []
    for email_id in email_ids:
        subject = get_email_subject(service, email_id)
        if subject.lower() == "logbook":
            logbook_emails.append(email_id)
    return logbook_emails

def delete_non_logbook_emails(service, email_ids):
    for email_id in email_ids:
        subject = get_email_subject(service, email_id)
        if subject.lower() != "logbook":
            # Move the email to the trash
            try:
                service.users().messages().trash(userId='me', id=email_id['id']).execute()
                print(f"Moved to trash email with subject: {subject}")
            except Exception as e:
                print(f"Failed to move to trash email with subject: {subject}, error: {e}")

def delete_specific_email(service, email_id):
    """Delete a specific email by ID."""
    try:
        service.users().messages().trash(userId='me', id=email_id).execute()
        print(f"Moved to trash email with ID: {email_id}")
    except Exception as e:
        print(f"Failed to move to trash email with ID: {email_id}, error: {e}")

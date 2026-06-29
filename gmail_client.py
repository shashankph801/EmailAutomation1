import base64
import os

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.send']


def gmail_api_configured():
    return all(os.environ.get(key) for key in (
        'GOOGLE_CLIENT_ID',
        'GOOGLE_CLIENT_SECRET',
    ))


def credentials_ready(sender_email=None, refresh_token=None):
    email = sender_email or (os.environ.get('SENDER_EMAIL') or '').strip()
    if not email:
        return False, 'Not logged in'
    token = refresh_token or os.environ.get('GOOGLE_REFRESH_TOKEN')
    if not token:
        return False, 'No refresh token found for this account'
    if not gmail_api_configured():
        return False, (
            'Gmail API credentials missing. Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET.'
        )
    return True, None


def get_gmail_service(refresh_token=None):
    token = refresh_token or os.environ['GOOGLE_REFRESH_TOKEN']
    creds = Credentials(
        None,
        refresh_token=token,
        token_uri='https://oauth2.googleapis.com/token',
        client_id=os.environ['GOOGLE_CLIENT_ID'],
        client_secret=os.environ['GOOGLE_CLIENT_SECRET'],
        scopes=GMAIL_SCOPES,
    )
    return build('gmail', 'v1', credentials=creds, cache_discovery=False)


def send_via_gmail_api(msg, refresh_token=None):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('ascii')
    service = get_gmail_service(refresh_token)
    result = service.users().messages().send(userId='me', body={'raw': raw}).execute()
    return result

#!/usr/bin/env python3
"""
One-time script to obtain a Gmail API refresh token.

Prerequisites:
  1. Create a Google Cloud project and enable the Gmail API.
  2. Configure the OAuth consent screen.
  3. Create OAuth 2.0 credentials (Desktop app).
  4. Download the client JSON or set GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET.

Usage:
  python get_gmail_refresh_token.py
  python get_gmail_refresh_token.py --credentials /path/to/client_secret.json
"""
import argparse
import json
import os

from google_auth_oauthlib.flow import InstalledAppFlow

from gmail_client import GMAIL_SCOPES


def load_client_config(path):
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if 'installed' in data:
        return data['installed']
    if 'web' in data:
        return data['web']
    return data


def main():
    parser = argparse.ArgumentParser(description='Obtain a Gmail API refresh token')
    parser.add_argument(
        '--credentials',
        help='Path to Google OAuth client JSON downloaded from Cloud Console',
    )
    args = parser.parse_args()

    if args.credentials:
        client_config = load_client_config(args.credentials)
        flow = InstalledAppFlow.from_client_config(
            {'installed': client_config},
            GMAIL_SCOPES,
        )
    else:
        client_id = os.environ.get('GOOGLE_CLIENT_ID')
        client_secret = os.environ.get('GOOGLE_CLIENT_SECRET')
        if not client_id or not client_secret:
            raise SystemExit(
                'Provide --credentials <client_secret.json> or set '
                'GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in the environment.'
            )
        flow = InstalledAppFlow.from_client_config(
            {
                'installed': {
                    'client_id': client_id,
                    'client_secret': client_secret,
                    'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
                    'token_uri': 'https://oauth2.googleapis.com/token',
                }
            },
            GMAIL_SCOPES,
        )

    creds = flow.run_local_server(port=0, prompt='consent')
    if not creds.refresh_token:
        raise SystemExit(
            'No refresh token returned. Revoke app access at '
            'https://myaccount.google.com/permissions and run again with consent.'
        )

    print('\nAdd these to your .env (local) or Render environment variables:\n')
    print(f'GOOGLE_CLIENT_ID={creds.client_id}')
    print(f'GOOGLE_CLIENT_SECRET={creds.client_secret}')
    print(f'GOOGLE_REFRESH_TOKEN={creds.refresh_token}')
    print(f'SENDER_EMAIL=<the Gmail account you signed in with>')


if __name__ == '__main__':
    main()

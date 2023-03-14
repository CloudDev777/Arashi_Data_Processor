# Author: Maika Rabenitas, maikacoleen1205@gmail.com

from __future__ import print_function

from googleapiclient.discovery import build
from google.oauth2 import service_account

SCOPE = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'

info = {
  "type": "service_account",
  "project_id": "ramenarashi",
  "private_key_id": "", # SENSITIVE DATA
  "private_key": "", # SENSITIVE DATA
  "client_email": "", # SENSITIVE DATA
  "client_id": "", # SENSITIVE DATA
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "" # SENSITIVE DATA
}

creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPE)
# creds = service_account.Credentials.from_service_account_file(
#     SERVICE_ACCOUNT_FILE, scopes=SCOPE)

# The ID and range of a spreadsheet.
SPREADSHEET_ID = "" # SENSITIVE DATA

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

from __future__ import print_function
import xlsxwriter
from time import sleep
from bs4 import BeautifulSoup

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from base64 import urlsafe_b64decode, urlsafe_b64encode
import email
import base64
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


email = input("Please enter an email: ")
sleep(1)
first_name = input("Please enter their first name: ")
sleep(1)
last_name = input("Please enter their last name: ")
sleep(1)
companyoption = input("Please enter a company name: ")
sleep(1)
filename = input("Please enter a filename: ")

stringed = filename + '.xlsm'


def createnew(emailed, firsted, lasted, companyed):
    storage_workbook = xlsxwriter.Workbook(stringed)
    worksheet = storage_workbook.add_worksheet('Info')
    worksheet.write('A1', 'Email')
    worksheet.write('B1', 'First Name')
    worksheet.write('C1', 'Last Name')
    worksheet.write('D1', 'Company')
    worksheet.write('A2', emailed)
    worksheet.write('B2', firsted)
    worksheet.write('C2', lasted)
    worksheet.write('D2', companyed)
    storage_workbook.close()


def initialize():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Call the Gmail API
    return build('gmail', 'v1', credentials=creds)


service = initialize()


def findingmessage(serviced):
    global results
    results = serviced.users().messages().list(userId='me', labelIds=[
        "INBOX"], q='from: umerfiaz251@gmail.com', maxResults=1).execute()
    global messageee
    messages = results.get('messages')
    for msg in messages:
        # Get the message from its id
        txt = serviced.users().messages().get(
            userId='me', id=msg['id']).execute()
        payload = txt['payload']
        headers = payload['headers']
        for d in headers:
            if d['name'] == 'Subject':
                subject = d['value']
            if d['name'] == 'From':
                sender = d['value']

        parts = payload.get('parts')[0]
        data = parts['body']
        data = data.replace("-", "+").replace("_", "/")
        decoded_data = base64.b64decode(data)
        soup = BeautifulSoup(decoded_data, "lxml")
        body = soup.body()
        print("Message: ", body)


def decodeemail():
    txt = service.users().messages().get(
        userId='me', id=messageee['id']).execute()
    payload = txt['payload']
    parts = payload.get('parts')[0]
    data = parts['body']['data']
    data = data.replace("-", "+").replace("_", "/")
    decoded_data = base64.b64decode(data)
    soup = BeautifulSoup(decoded_data, "lxml")
    body = soup.body()
    print("Message: ", body)


if __name__ == '__main__':
    createnew(email, first_name, last_name, companyoption)
    findingmessage(service)

# See https://chozinthet20602.medium.com/sending-email-with-python-using-gmail-api-33628e36306a

import os
import json
import base64
import mimetypes

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

with open('_config.json', 'r') as f:
    config = json.loads(f.read())
    
    CLIENT_SECRET_PATH = config.get('clientSecretPath')
    CREDENTIALS_PATH = config.get('credentialsPath')
    SENDER_EMAIL = config.get('senderEmail')
    RECEIVER_EMAIL = SENDER_EMAIL
    
    SCOPES = ['https://mail.google.com/']
    
    del config

credentials = None
if os.path.exists(CREDENTIALS_PATH):
    credentials = Credentials.from_authorized_user_file(CREDENTIALS_PATH, SCOPES)

if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
        except Exception as e:
            print('Credential error: {}'.format(e))
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_PATH, SCOPES)
            credentials = flow.run_local_server(port=0)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_PATH, SCOPES)
        credentials = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(CREDENTIALS_PATH, 'w') as f:
        f.write(credentials.to_json())

email_service = build('gmail', 'v1', credentials=credentials)


def send_update(file_list: list) -> None:
    
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message['From'] = SENDER_EMAIL
    message['To'] = RECEIVER_EMAIL
    message['Subject'] = 'Mackolik update'
    
    # Body is just a list of files
    message.attach(MIMEText('\n'.join(file_list), 'plain'))
    # See https://stackoverflow.com/a/11077828/5957296
    for filename in file_list:
        ctype, encoding = mimetypes.guess_type(filename)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        # Open file / read | binary
        with open(filename, 'rb') as file:
            # Add file as application/octet-stream
            part = MIMEBase(maintype, subtype)
            part.set_payload(file.read())
            # Encode file b64
            encoders.encode_base64(part)
            # Add header to attachment
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            # Add attachment to message
            message.attach(part)
            
    # Create a multipart message and set headers
    message = MIMEMultipart()
    message['From'] = FROM_LINE
    message['To'] = TO_LINE
    message['Subject'] = SUBJECT_LINE

    # body is just a list of files
    message.attach(MIMEText('\n'.join(file_list), 'plain'))
    # see https://stackoverflow.com/a/11077828/5957296
    for filename in file_list:
        ctype, encoding = mimetypes.guess_type(filename)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        # open file / read | binary
        with open(filename, 'rb') as file:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase(maintype, subtype)
            part.set_payload(file.read())
            # Encode file in ASCII characters to send by email    
            encoders.encode_base64(part)
            # Add header as key/value pair to attachment part
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            # Add attachment to message
            message.attach(part)
    
    # Attempt to send the message
    message_obj = {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}
    try:
        message_result = (
            email_service.
            users().
            messages().
            send(userId=SENDER_EMAIL, body=message_obj).
            execute()
        )
        print('Message out: {}'.format(message_result))
    except Exception as e:
        print('Gmail error: {}'.format(e))

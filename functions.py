import os
import pickle
import pandas as pd
from datetime import datetime
import numpy as np
from apiclient.discovery import build, MediaFileUpload
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/documents']

# The ID and range of the spreadsheet.
JOURNAL_SPREADSHEET_ID = '1g6C9nSwY_3I-ADzEUl87AKKyx3EKndM5PB2AzSXSvs4'
SAMPLE_RANGE_NAME = 'Form Responses 1!A1:H'

# Document and folder IDs
DOCUMENT_ID = '1-caO4bDTP-VH8CUdOH-JteR4fSRtqw6VArsDLjqf4Ow'
PARENT_FOLDER = '1B296OchXs5cA5fflhyzkopiggd9hQavB'

DATE_FORMAT = '%A, %B %d, %Y'
TIMESTAMP_FORMAT = '%m/%d/%Y %H:%M:%S'

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
token_file = os.path.join('Data', 'token.pickle')
if os.path.exists(token_file):
    with open(token_file, 'rb') as token:
        creds = pickle.load(token)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('Data\credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)

    # Save the credentials for the next run
    with open(token_file, 'wb') as token:
        pickle.dump(creds, token)


def read_journal_responses():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API to read entries
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=JOURNAL_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()

    # Get data and format it into a Pandas dataframe
    df = format_journal_responses(result.get('values', []))

    return df


def format_journal_responses(values):

    # Get columns names
    col_names = values[0]

    # Create a dictionary of entries
    df = dict((e, []) for e in col_names)
    for entry in values[1:]:

        for i, key in enumerate(col_names):

            df[key].append(entry[i])

    return pd.DataFrame.from_dict(df)


def generate_msg(df):

    st_ind = 100
    text = 'hi\n\n'
    requests = []
    body_end_ind = 1
    for i in range(3):#len(df)):

        # Create header
        date_text = df['Timestamp'][i]
        date_text = datetime.strptime(date_text, TIMESTAMP_FORMAT)
        date_text = datetime.strftime(date_text, DATE_FORMAT) + '\n'

        header_st_ind = body_end_ind + 1
        header_end_ind = header_st_ind + len(date_text)

        # Create body
        body_text = df['Journal entry'][i] + '\n\n'

        body_st_ind = header_end_ind
        body_end_ind = body_st_ind + len(body_text)

        requests.append(
            {  # Header
                'insertText': {
                    'location': {
                        'index': header_st_ind,
                    },
                    'text': date_text
                }
            })
        requests.append(
            {  # Journal entry
                'insertText': {
                    'location': {
                        'index': body_st_ind,
                    },
                    'text': body_text
                }
            })
        requests.append(
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': header_st_ind,
                        'endIndex': header_end_ind
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'HEADING_1',
                        'spaceAbove': {
                            'magnitude': 10.0,
                            'unit': 'PT'
                        },
                        'spaceBelow': {
                            'magnitude': 10.0,
                            'unit': 'PT'
                        }
                    },
                    'fields': 'namedStyleType,spaceAbove,spaceBelow'
                }
            }
        )



def get_start_index(document):

    a = 1


def make_journal_document(df):
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """

    service = build('docs', 'v1', credentials=creds)

    body = {'name': [], 'parents': ["11y0SiVNpROcQKCTVGWs-6xAgTpVLpZEE"]}

    # Retrieve the documents contents from the Docs service.
    document = service.documents().get(documentId=DOCUMENT_ID).execute()

    st_ind = 10
    text = 'hi\n\n'

    requests = [
        {
            'insertText': {
                'location': {
                    'index': st_ind,
                },
                'text': text
            }
        }
    ]

    result = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

    print('The title of the document is: {}'.format(document.get('title')))



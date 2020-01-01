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
        flow = InstalledAppFlow.from_client_secrets_file('Data/credentials.json', SCOPES)
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


def generate_msg(df, doc_inds):

    st_ind = 1
    text = 'hi\n\n'
    requests = []
    body_end_ind = 1
    for i in range(len(df)):

        # Create header
        date_text = df['Timestamp'][i]
        date_text = datetime.strptime(date_text, TIMESTAMP_FORMAT)
        date_text = datetime.strftime(date_text, DATE_FORMAT) + '\n'

        header_st_ind = int(doc_inds[i] + body_end_ind)
        header_end_ind = int(header_st_ind + len(date_text))

        # Create body
        body_text = 'Mood:\t\t\t\t{}\n'.format(df['Mood'][i]) + \
                    'Read scriptures:\t\t{}\n'.format(df['Read scriptures'][i]) + \
                    'Exercise:\t\t\t{}\n'.format(df['Exercise'][i]) + \
                    'Study:\t\t\t\t{}\n'.format(df['Study (language, programming, music)'][i]) + \
                    'Gratitude:\t\t\t{}\n'.format(df['Gratitude'][i]) + \
                    'New goals and progress:\t{}\n'.format(df['New goals and progress'][i]) + \
                    '\n{}\n\n\n'.format(df['Journal entry'][i])

        body_st_ind = int(header_end_ind)
        body_end_ind = int(body_st_ind + len(body_text))

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
                        'endIndex': body_end_ind
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'NORMAL_TEXT',
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

    return requests


def get_doc_dates(document):
    """
    Parses out the dates in journal entries and returns the dates and their indices (location) as datetimes.
    Args:
        document (document): result of the Google doc api for the journal document

    Returns:
        dates (list of datetime): dates in the document
        date_inds (index of date)
    """

    date_inds = list()
    dates = list()

    for i, elem in enumerate(document['body']['content']):

        if 'paragraph' in elem.keys():

            # Get paragraph text
            text = elem['paragraph']['elements'][0]['textRun']['content']

            # Strip formatting
            text = text.strip()

            # Attempt to read date format
            try:
                dt = datetime.strptime(text, DATE_FORMAT)

            except ValueError:
                dt = False

            if dt:
                dates.append(dt)
                date_inds.append(i)

    return dates, date_inds


def get_doc_indices(dates, date_inds, df):

    # Get the dates from the form
    form_dates = [datetime.strptime(i, TIMESTAMP_FORMAT) for i in df['Timestamp']]

    # Concatenate dates
    dates = form_dates + dates

    # Sort dates
    inds = np.argsort(dates)

    # Get doc list indices
    doc_list_inds = [np.argwhere(inds == i)[0][0] for i in range(len(form_dates))]

    # Determine the correct document index
    doc_inds = []
    for ind in doc_list_inds:

        if len(date_inds) == 0:
            doc_inds.append(0)
        elif ind >= len(date_inds):
            doc_inds.append(np.max(date_inds))

        else:
            doc_inds.append(date_inds[ind])

    return doc_inds


def filter_dates(dates, df):
    """
    Remove journal entries which already appear in the Google document. This is done by comparinng dates; each day is only allowed one entry.
    Args:
        dates (list of datetime): a list of datetimes derrived from the Google document.
        df (pandas dataframe): a Pandas dataframe containing all journal entries.

    Returns:

    """

    # Convert dataframe timestamps from strings to datetime
    df_dt = [datetime.strptime(i, TIMESTAMP_FORMAT) for i in df['Timestamp']]

    # Remove hours, minutes, seconds to allow matching with journal entry
    df_dt = [i.replace(hour=0, minute=0, second=0) for i in df_dt]

    # Find the days which appear in df, but not in dates
    ind_unique = list()
    for i, dt in enumerate(df_dt):
        if dt not in dates:
            ind_unique.append(i)

    # Filter dataframe
    df_filt = df.iloc[ind_unique]

    # Reset indices
    df_filt = df_filt.reset_index()

    return df_filt


def update_journal(df):

    # Set up service
    service = build('docs', 'v1', credentials=creds)

    # Retrieve the documents contents from the Docs service.
    document = service.documents().get(documentId=DOCUMENT_ID).execute()

    # Get list of dates and corresponding indicies which have journal entries
    dates, date_inds = get_doc_dates(document)

    # Determine which new entries are needed
    df = filter_dates(dates, df)

    # Find where the new entries should be inserted
    doc_inds = get_doc_indices(dates, date_inds, df)

    # Generate message
    requests = generate_msg(df, doc_inds)

    result = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

    print('The title of the document is: {}'.format(document.get('title')))

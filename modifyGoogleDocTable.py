#!/usr/bin/env python3.6
from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents']

# Configuration
DOCUMENT_ID = 'google_document_id'
# Searchs cell with TARGETCELLTEXT
TARGETCELLTEXT = 'some_text'
# Inserts text in the next column cell
TEXTTOINSERT = 'text_to_insert_in_next_cell'

def main():
    # Get credentials
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)

    # Retrieve the documents contents from the Docs service.
    document = service.documents().get(documentId=DOCUMENT_ID).execute()
    document_content = document.get('body').get('content')
    print('The title of the document is: {}'.format(document.get('title')))

    # Find location to insert text
    getNextCell = False
    fillLocation = []
    for value in document_content:
      if 'table' in value:
        #print(value)
        table = value.get('table')
        #print(table)
        for row in table.get('tableRows'):
          cells = row.get('tableCells')
          for cell in cells:
            startIndex = cell.get('content')[0].get('paragraph').get('elements')[0].get('startIndex')
            endIndex = cell.get('content')[0].get('paragraph').get('elements')[0].get('endIndex')
            text = cell.get('content')[0].get('paragraph').get('elements')[0].get('textRun').get('content')
            #print(startIndex, endIndex, text)
            if getNextCell:
              if text.lstrip().rstrip() == '': fillLocation = [startIndex, endIndex]
              getNextCell = False
            if TARGETCELLTEXT in text:
              getNextCell = True

    # Insert NTR
    if len(fillLocation) != 0:
      requests = [
        #{'deleteContentRange': {
        #  'range': {
        #    'startIndex': fillLocation[0],
        #    'endIndex':fillLocation[1]
        #    }
        #  }
        #},
        {'insertText': {
           'location': {
             'index': fillLocation[0],
            },
           'text': TEXTTOINSERT
          }
        },
      ]
      result = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()
      print('Inserted '+TEXTTOINSERT)
    
if __name__ == '__main__':
    main()

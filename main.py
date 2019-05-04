import re
import html
import pickle
import os.path
from tqdm import tqdm
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

TAG_RE = re.compile('#[A-Za-z0-9-_]+')

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/drive'
]


class Drive:
    def __init__(self):
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
                creds = flow.run_local_server()

            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build('drive', 'v3', credentials=creds)
        self.sheets = build('sheets', 'v4', credentials=creds).spreadsheets()

    def list_folder(self, folder_id):
        """List documents in a folder"""
        params = {
            'q': '"{}" in parents'.format(folder_id),
            'fields': '*'
        }
        resp = self.service.files().list(**params).execute()
        files = resp['files']
        while 'nextPageToken' in resp:
            page = resp['nextPageToken']
            resp = self.service.comments().list(pageToken=page, **params).execute()
            files += resp['files']

        # Limit to docs
        return [f for f in files if f['mimeType'] == 'application/vnd.google-apps.document']

    def get_tags(self, document_id):
        """Get tagged text and tags for a document"""
        # file = self.service.files().get(fileId=document_id).execute()
        # rev = self.service.revisions().get(fileId=document_id, revisionId='head').execute()

        # Fetch all comments
        params = {
            'fileId': document_id,
            'includeDeleted': False,
            'fields': '*'
        }
        resp = self.service.comments().list(**params).execute()
        comments = resp['comments']
        while 'nextPageToken' in resp:
            page = resp['nextPageToken']
            resp = self.service.comments().list(pageToken=page, **params).execute()
            comments += resp['comments']

        # Don't include resolved comments
        comments = [c for c in comments if not c['resolved']]

        # Extract tags
        tagged = []
        for c in comments:
            tags = [t.strip('#') for t in TAG_RE.findall(c['content'])]
            if not tags: continue
            text = c['quotedFileContent']['value']
            text = html.unescape(text)
            tagged.append((text, tags))
        return tagged

    def update_spreadsheet(self, sheet_id, tagged):
        # Reset spreadsheet
        self.sheets.values().clear(spreadsheetId=sheet_id, range='A:Z').execute()

        headers = ['Document ID', 'Text', 'Tags']
        values = [[doc_id, text, ', '.join(tags)] for doc_id, (text, tags) in tagged]
        body = {
            'values': [headers] + values
        }
        range = '1:{}'.format(len(body['values']))
        self.sheets.values().update(
            spreadsheetId=sheet_id,
            body=body,
            range=range,
            valueInputOption='RAW').execute()

if __name__ == '__main__':
    import sys
    try:
        FOLDER_ID = sys.argv[1]
        SHEET_ID = sys.argv[2]
    except IndexError:
        print('Please specify the folder and spreadsheet ID')
        sys.exit(1)

    tags = []
    drive = Drive()

    print('Reading comments...')
    files = drive.list_folder(FOLDER_ID)
    for f in tqdm(files):
        doc_id = f['id']
        tagged = drive.get_tags(doc_id)
        tags += [(doc_id, t) for t in tagged]

    print('Updating spreadsheet...')
    drive.update_spreadsheet(SHEET_ID, tags)
    print('Done')
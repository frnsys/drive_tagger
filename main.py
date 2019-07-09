import re
import html
import click
import pickle
import os.path
from tqdm import tqdm
from collections import defaultdict, namedtuple
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

TAG_RE = re.compile('#[A-Za-z0-9-_]+')

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/drive'
]

Comment = namedtuple('Comment', ['highlighted', 'text', 'tags', 'user', 'url'])


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

    def get_folder_tags(self, folder_id):
        """Get tagged text and tags for
        all documents in a folder"""
        tags = []
        files = self.list_folder(folder_id)
        doc_meta = {}
        for f in tqdm(files):
            doc_id = f['id']
            file = self.service.files().get(fileId=doc_id).execute()
            doc_meta[doc_id] = {
                'title': file['name']
            }
            tagged = self.get_doc_tags(doc_id)
            tags += [(doc_id, com) for com in tagged]
        return tags, doc_meta

    def get_doc_tags(self, document_id):
        """Get tagged text and tags for a document"""
        # file = self.service.files().get(fileId=document_id).execute()
        # rev = self.service.revisions().get(fileId=document_id, revisionId='head').execute()
        doc_url = 'https://docs.google.com/document/d/{}/edit'.format(document_id)

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
        comments = [c for c in comments if not c.get('resolved', False)]

        # Extract tags
        tagged = []
        for comment in comments:
            user = comment['author']['displayName']
            text = comment['content']
            highlighted = comment['quotedFileContent']['value']
            highlighted = html.unescape(highlighted)
            comment_url = '{}?disco={}'.format(doc_url, comment['id'])
            tags = []
            # Get tags from main comment and replies
            for c in [comment] + comment['replies']:
                # Standardize tags to lowercase
                tags += [t.strip('#').lower() for t in TAG_RE.findall(c['content'])]
            if not tags: continue
            tagged.append(Comment(highlighted, text, tags, user, comment_url))
        return tagged

    def create_sheet(self, sheet_id, title):
        requests = [{
            'addSheet': {
                'properties': {
                    'title': title
                }
            }
        }]
        body = {'requests': requests}
        resp = self.sheets.batchUpdate(spreadsheetId=sheet_id, body=body).execute()
        return resp['replies'][0]['addSheet']['properties']

    def update_spreadsheet(self, sheet_id, tagged, doc_meta):
        # Get sub-sheets
        resp = self.sheets.get(spreadsheetId=sheet_id).execute()
        sheets = [s['properties'] for s in resp['sheets']]

        # Delete sheets for missing tags
        unique_tags = [com.tags for _, com in tagged]
        unique_tags = set([t for ts in unique_tags for t in ts])
        to_delete = [s for s in sheets if s['index'] not in [0, 1] and s['title'] not in unique_tags]
        if to_delete:
            requests = [{
                'deleteSheet': {
                    'sheetId': s['sheetId']
                }
            } for s in to_delete]
            body = {'requests': requests}
            resp = self.sheets.batchUpdate(spreadsheetId=sheet_id, body=body).execute()

        # Reset sheets
        requests = [{
            'updateCells': {
                'range': {
                    'sheetId': s['sheetId']
                },
                'fields': 'userEnteredValue'
            }
        } for s in sheets]
        body = {'requests': requests}
        self.sheets.batchUpdate(spreadsheetId=sheet_id, body=body).execute()

        # Count documents tags appear in
        tag_counts = defaultdict(set)
        for doc_id, com in tagged:
            for tag in com.tags:
                tag_counts[tag].add(doc_id)
        tag_counts = {tag: len(docs) for tag, docs in tag_counts.items()}

        # Update first sheet (tag list)
        headers = ['Tag', '# Documents']
        values = [[tag, n_docs] for tag, n_docs in tag_counts.items()]
        body = {
            'values': [headers] + values
        }
        range = '1:{}'.format(len(body['values']))
        self.sheets.values().update(
            spreadsheetId=sheet_id,
            body=body,
            range=range,
            valueInputOption='RAW').execute()

        # Update second sheet (all tags)
        headers = ['Document ID', 'Title', 'Highlighted', 'Tags', 'Comment', 'User', 'Url']
        values = [[doc_id, doc_meta[doc_id]['title'], com.highlighted, ', '.join(com.tags), com.text, com.user, com.url]
                  for doc_id, com in tagged]
        try:
            all_tags_sheet = next(s for s in sheets if s['index'] == 1)
        except StopIteration:
            # Create if necessary
            all_tags_sheet = self.create_sheet(sheet_id, 'All Tags')

        requests = [{
            'updateCells': {
                'rows': [{
                    'values': [{
                        'userEnteredValue': {
                            'stringValue': c
                        }
                    } for c in m]
                } for m in [headers] + values],
                'range': {
                    'sheetId': all_tags_sheet['sheetId']
                },
                'fields': 'userEnteredValue'
            }
        }, {
            'updateSheetProperties': {
                'properties': {
                    'sheetId': all_tags_sheet['sheetId'],
                    'title': 'All Tags'
                },
                'fields': 'title'
            }

        }]
        body = {'requests': requests}
        resp = self.sheets.batchUpdate(spreadsheetId=sheet_id, body=body).execute()

        # Create per-tag sheets
        tag_groups = defaultdict(list)
        for doc_id, com in tagged:
            for tag in com.tags:
                tag_groups[tag].append((doc_id, com.highlighted))

        sheet_requests = []
        for tag, mentions in tqdm(tag_groups.items()):
            # Check if sheet exists for this tag
            for s in sheets:
                if s['title'] == tag:
                    sheet = s
                    break
            else:
                # Create new sheet
                sheet = self.create_sheet(sheet_id, tag)

            # This is heinous
            sheet_requests.append({
                'updateCells': {
                    'rows': [{
                        'values': [{
                            'userEnteredValue': {
                                'stringValue': c
                            }
                        } for c in m]
                    } for m in mentions],
                    'range': {
                        'sheetId': sheet['sheetId']
                    },
                    'fields': 'userEnteredValue'
                }
            })
        body = {'requests': sheet_requests}
        self.sheets.batchUpdate(spreadsheetId=sheet_id, body=body).execute()

@click.group()
def main():
    pass

@main.command()
@click.argument('folder_id')
@click.argument('sheet_id')
def sync(folder_id, sheet_id):
    drive = Drive()

    print('Reading comments...')
    tags, doc_meta = drive.get_folder_tags(folder_id)

    print('Updating spreadsheet...')
    drive.update_spreadsheet(sheet_id, tags, doc_meta)
    print('Done')


if __name__ == '__main__':
    main()
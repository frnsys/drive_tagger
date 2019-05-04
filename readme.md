# Drive Tagger

Will crawl documents in a Google Drive folder and extract tags (anything starting with `#` in comments), then create a spreadsheet of the tagged text.

## Setup

- Install dependencies: `pip install tqdm google-api-python-client google-auth-httplib2 google-auth-oauthlib`
- Go to <https://developers.google.com/drive/api/v3/quickstart/python> and click "Enable the Drive API" to download a `credentials.json` file. Place that in this folder.
- Go to <https://console.developers.google.com/apis/api/sheets.googleapis.com> and enable the Sheets API.
- Get the ID of the folder you want to crawl. Open the folder in Drive and look at the URL - the ID is the last part, i.e. `https://drive.google.com/drive/u/0/folders/<FOLDER_ID>`
- Create a spreadsheet that you want to populate with the tagged data. Get that ID, it's also in the URL: `https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit#gid=0`

## Usage

    python main.py FOLDER_ID SHEET_ID
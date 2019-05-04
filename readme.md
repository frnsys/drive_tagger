# Drive Tagger

Will crawl documents in a Google Drive folder and extract tags (anything starting with `#` in comments), then create a spreadsheet of the tagged text.

## Setup

- Go to <https://developers.google.com/drive/api/v3/quickstart/python> and click "Enable the Drive API" to download a `credentials.json` file. Place that in this folder.
- Get the ID of the folder you want to crawl. Open the folder in Drive and look at the URL - the ID is the last part, i.e. `https://drive.google.com/drive/u/0/folders/FOLDER_ID`

## Usage

    python main.py FOLDER_ID
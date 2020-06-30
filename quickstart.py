from __future__ import print_function
import pickle
import os.path
import csv

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


def createLink(id, mimeType):
    switcher ={
        'application/vnd.google-apps.spreadsheet':'https://docs.google.com/spreadsheets/d/',
        'application/vnd.google-apps.document':'https://docs.google.com/document/d/',
        'application/vnd.google-apps.presentation':'https://docs.google.com/presentation/d/',
        'application/vnd.google-apps.folder':'https://drive.google.com/drive/folders/',
        'application/vnd.google-apps.file':'https://drive.google.com/file/d/'
    }
    linkStart = switcher.get(mimeType, 'https://drive.google.com/file/d/')
    link = linkStart + id
    return link

def fileType(mimeType):
    switcher ={
        'application/vnd.google-apps.spreadsheet':'sheet',
        'application/vnd.google-apps.document':'doc',
        'application/vnd.google-apps.folder':'folder',
        'application/vnd.google-apps.presentation': 'slides',
        'application/vnd.google-apps.file':'file'
    }
    file = switcher.get(mimeType, mimeType)
    return file

def writeCSV(csv_writer, items,folderName):
    for item in items:
        print(u'{0} ({1})'.format(item['name'], item['id']))
        pageTitle = '"' + item['name'] + '"'
        fileId = item['id']
        # id = item['id']
        mimeType = fileType(item['mimeType'])
        link = createLink(fileId, item['mimeType'])
        folderNameUpdate = '"' + folderName +'"'
        if(item['mimeType']!='application/vnd.google-apps.folder'):
            csv_writer.writerow(
                {'page_title': pageTitle, 'type': mimeType, 'link': link, 'folder_name': folderNameUpdate})
        if(item['mimeType']=='application/vnd.google-apps.folder'):
            newfolderName = folderName + "/" + item['name']
            formattedFolderId = "'{}'".format(item['id'])
            query = formattedFolderId + ' in parents'
            formattedQuery = '"{}"'.format(query)
            # Call the Drive v3 API
            newresults = service.files().list(q=query,
                                           pageSize=1000,
                                           fields="nextPageToken, files(kind,id,name,mimeType,parents)").execute()
            newitems = newresults.get('files', [])
            writeCSV(csv_writer,newitems,newfolderName)

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']
creds = None
csvLocation = 'links.csv'

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
service = build('drive', 'v3', credentials=creds)
folderAddress = input('Enter Folder Address: ')
folderName = input('Enter Folder Name: ')
folderId = folderAddress.split('/folders/')[1]
formattedFolderId = "'{}'".format(folderId)
query = formattedFolderId + ' in parents'
formattedQuery = '"{}"'.format(query)

# Call the Drive v3 API
results = service.files().list(q=query,
    pageSize=1000, fields="nextPageToken, files(kind,id,name,mimeType,parents)").execute()
items = results.get('files', [])

if not items:
    print('No files found.')
else:
    print('Files:')

    with open(csvLocation, mode='w') as csvFile:
        fieldnames = ['folder_name','page_title', 'type', 'link','id']
        csv_writer = csv.DictWriter(csvFile, fieldnames=fieldnames)
        csv_writer.writeheader()
        writeCSV(csv_writer,items,folderName)

        # Announce that we've finished
        print("Finished! Written to CSV.")
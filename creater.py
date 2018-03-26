from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from apiclient import errors


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/script-python-quickstart.json
SCOPES = ['https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Apps Script API Python'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'script-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


"""Shows basic usage of the Apps Script API.

Call the Apps Script API to create a new script project, upload a file to the
project, and log the script's URL to the user.
"""

def create_document_with_paragraph(string_to_write):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('script', 'v1', http=http)
    drive = discovery.build('drive', 'v3', http=http)
    try:
        request = {'function': 'createDocument'}
        response = service.scripts().run(scriptId='1MArmq6LjsPkAjQ5krm1vCWuMslLrw9kLZyWvlCRoUP_QyOPOVFiTgqoa', body=request).execute()
        docId = response['response']['result']
        print(response)
        request = {'function': 'appendParagraphToDocumentById', 'parameters': [docId, string_to_write]}
        response = service.scripts().run(scriptId='1MArmq6LjsPkAjQ5krm1vCWuMslLrw9kLZyWvlCRoUP_QyOPOVFiTgqoa', body=request).execute()
        print(response)
    except errors.HttpError as e:
        print(e.content)


def main():
    # Authorize and create a service object.
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('script', 'v1', http=http)
    drive = discovery.build('drive', 'v3', http=http)

# 1Rv62tEB2C_98iqry2ROM_BtrdGzybDa00_97p1bWW1E

    try:
        request = {'function': 'createDocument'}
        response = service.scripts().run(scriptId='1MArmq6LjsPkAjQ5krm1vCWuMslLrw9kLZyWvlCRoUP_QyOPOVFiTgqoa', body=request).execute()
        print(response)
        # docId = response['response']['result']
        # print(docId)
        testJson = {"heading": "Awesome heading", "text": "This is cool paragraph text"}
        request = {'function': 'printJson', 'parameters': [testJson]}
        response = service.scripts().run(scriptId='1MArmq6LjsPkAjQ5krm1vCWuMslLrw9kLZyWvlCRoUP_QyOPOVFiTgqoa', body=request).execute()
        print(response)
    except errors.HttpError as e:
        # The API encountered a problem.
        print(e.content)


if __name__ == '__main__':
    main()

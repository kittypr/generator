from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from apiclient import errors

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/script-python-quickstart.json
SCOPES = ['https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Apps Script API Python. Google docs generator'


class GDoc:
    def __init__(self, writer, id):
        self.writer = writer
        self.id = id

    def write(self, string):
        self.writer.append_paragraph(self.id, string)


class GDocsWriter:
    def __init__(self):
        credentials = self.get_credentials()
        http = credentials.authorize(httplib2.Http())
        self.app_script_service = discovery.build('script', 'v1', http=http)

    def get_credentials(self):
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
                                       'generator.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            credentials = tools.run_flow(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def create_new_document(self, document_name):
        try:
            request = {'function': 'createDocument', 'parameters': [document_name]}
            response = self.app_script_service.scripts().run(scriptId='1MArmq6LjsPkAjQ5krm1vCWuMslLrw9kLZyWvlCRoUP_QyOPOVFiTgqoa', body=request).execute()
            doc_id = response['response']['result']
            return GDoc(self, doc_id)
        except errors.HttpError as e:
            print(e.content)

    def append_paragraph(self, doc_id, string_to_write):
        try:
            request = {'function': 'appendParagraphToDocumentById', 'parameters': [doc_id, string_to_write]}
            response = self.app_script_service.scripts().run(scriptId='1MArmq6LjsPkAjQ5krm1vCWuMslLrw9kLZyWvlCRoUP_QyOPOVFiTgqoa', body=request).execute()
            print(response)
        except errors.HttpError as e:
            print(e.content)

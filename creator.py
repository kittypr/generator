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

SCRIPT_ID = '1MArmq6LjsPkAjQ5krm1vCWuMslLrw9kLZyWvlCRoUP_QyOPOVFiTgqoa'


class GDoc:
    def __init__(self, writer, id):
        self.writer = writer
        self.id = id

    def write_para(self, string):
        self.writer.append_paragraph(self.id, string)

    def write_heading(self, string, level):
        self.writer.append_heading_paragraph(self.id, string, level)

    def write_table(self, table):
        self.writer.append_table(self.id, table)


class GDocsWriter:  # TODO proper exception handling
    def __init__(self, flags):
        credentials = self.get_credentials(flags)
        http = credentials.authorize(httplib2.Http())
        self.app_script_service = discovery.build('script', 'v1', http=http)

    def create_new_document(self, document_name):
        try:
            request = {'function': 'createDocument', 'parameters': [document_name]}
            response = self.app_script_service.scripts().run(scriptId=SCRIPT_ID, body=request).execute()
            doc_id = response['response']['result']
            return GDoc(self, doc_id)
        except errors.HttpError as e:
            print(e.content)

    def append_paragraph(self, doc_id, string_to_write):
        request = {'function': 'appendParagraphToDocumentById', 'parameters': [doc_id, string_to_write]}
        try:
            response = self.app_script_service.scripts().run(scriptId=SCRIPT_ID, body=request).execute()
            print(response)
        except errors.HttpError as e:
            print(e.content)

    def append_heading_paragraph(self, doc_id, string_to_write, heading_level):
        request = {'function': 'appendHeadingParagraphToDocumentById', 'parameters': [doc_id, string_to_write, heading_level]}
        try:
            response = self.app_script_service.scripts().run(scriptId=SCRIPT_ID, body=request).execute()
            print(response)
        except errors.HttpError as e:
            print(e.content)

    def append_table(self, doc_id, table):
        request = {'function': 'appendTableToDocumentById', 'parameters': [doc_id, table]}
        try:
            response = self.app_script_service.scripts().run(scriptId=SCRIPT_ID, body=request).execute()
            print(response)
        except errors.HttpError as e:
            print(e.content)

    def get_credentials(self, flags):
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
        credential_path = os.path.join(credential_dir, 'generator.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            credentials = tools.run_flow(flow, store, flags)
            print('Storing credentials to ' + credential_path)
        return credentials
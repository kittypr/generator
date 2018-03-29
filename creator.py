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

    def write(self, string):
        self.writer.append_paragraph(self.id, string)

    def flush(self):
        self.writer.execute_batch()


def request_callback(id, response, exception):
    print(response)


class GDocsWriter:  # TODO proper exception handling
    def __init__(self, flags, batch_capacity=100):
        credentials = self.get_credentials(flags)
        http = credentials.authorize(httplib2.Http())
        self.app_script_service = discovery.build('script', 'v1', http=http)
        self.batch = None
        self.batch_capacity = batch_capacity
        self.batch_size = 0

    def create_new_document(self, document_name):
        try:
            request = {'function': 'createDocument', 'parameters': [document_name]}
            response = self.app_script_service.scripts().run(scriptId=SCRIPT_ID, body=request).execute()
            doc_id = response['response']['result']
            return GDoc(self, doc_id)
        except errors.HttpError as e:
            print(e.content)

    def append_paragraph(self, doc_id, string_to_write):
        if self.batch is None:
            self.batch = self.app_script_service.new_batch_http_request(callback=request_callback)
        request = {'function': 'appendParagraphToDocumentById', 'parameters': [doc_id, string_to_write]}
        self.batch.add(self.app_script_service.scripts().run(scriptId=SCRIPT_ID, body=request))
        self.batch_size += 1
        if self.batch_size == self.batch_capacity:
            try:
                self.batch.execute()
                print('batch request executed on append_paragraph')
                self.batch_size = 0
                self.batch = None
            except errors.HttpError as e:
                print(e.content)

    def append_heading_paragraph(self, doc_id, string_to_write, heading_level):
        if self.batch is None:
            self.batch = self.app_script_service.new_batch_http_request(callback=request_callback)
        request = {'function': 'appendHeadingParagraphById', 'parameters': [doc_id, string_to_write, heading_level]}
        self.batch.add(self.app_script_service.scripts().run(scriptId=SCRIPT_ID, body=request))
        self.batch_size += 1
        if self.batch_size == self.batch_capacity:
            try:
                self.batch.execute()
                print('batch request executed on append_heading_paragraph')
                self.batch_size = 0
                self.batch = None
            except errors.HttpError as e:
                print(e.content)

    def execute_batch(self):
        if self.batch is None:
            return
        try:
            self.batch.execute()
            print('batch request executed on execute_batch')
            self.batch_size = 0
            self.batch = None
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
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from apiclient import http
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from apiclient import errors

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/script-python-quickstart.json
SCOPES = ['https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Apps Script API Python. Google docs generator'


class Uploader:
    def __init__(self, flags):
        credentials = self.get_credentials(flags)
        http = credentials.authorize(httplib2.Http())
        self.drive_service = discovery.build('drive', 'v3', http=http)

    def upload(self):
        file_metadata = {'name': 'Test_Report2.docx', 'mimeType': 'application/vnd.google-apps.document'}
        media = http.MediaFileUpload('Test_Report2.docx',
                                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        file = self.drive_service.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()

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
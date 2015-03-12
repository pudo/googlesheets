import os
from gdata.spreadsheet.service import SpreadsheetsService
from gdata.docs.client import DocsClient

from googlesheets.util import SOURCE_NAME


class Connection(object):
    """ The connection controls user credentials and the connections
    to various google services. """

    def __init__(self, google_user=None, google_password=None):
        self._google_user = google_user
        self._google_password = google_password

    @property
    def google_user(self):
        """ The Google user name set via function arguments or
        the environment. """
        return self._google_user or os.environ.get('GOOGLE_USER')

    @property
    def google_password(self):
        """ The Google account password set via function arguments or
        the environment. """
        return self._google_password or os.environ.get('GOOGLE_PASSWORD')

    @property
    def docs_client(self):
        """ A DocsClient singleton, used to look up spreadsheets
        by name. """
        if not hasattr(self, '_docs_client'):
            client = DocsClient()
            client.ClientLogin(self.google_user, self.google_password,
                               SOURCE_NAME)
            self._docs_client = client
        return self._docs_client

    @property
    def sheets_service(self):
        """ A SpreadsheetsService singleton, used to perform
        operations on the actual spreadsheet. """
        if not hasattr(self, '_sheets_service'):
            service = SpreadsheetsService()
            service.email = self.google_user
            service.password = self.google_password
            service.source = SOURCE_NAME
            service.ProgrammaticLogin()
            self._sheets_service = service
        return self._sheets_service

    @classmethod
    def connect(cls, conn=None, google_user=None,
                google_password=None):
        if conn is None:
            conn = cls(google_user=google_user,
                       google_password=google_password)
        return conn

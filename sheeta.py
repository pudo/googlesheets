import os
from gdata.docs.data import Resource
from gdata.docs.service import DocsService, DocumentQuery
from gdata.docs.client import DocsClient
from gdata.spreadsheet.service import SpreadsheetsService

SOURCE_NAME = 'Sheeta/Python'


class Connection(object):
    """ The connection controls user credentials and the connections
    to various google services. """

    def __init__(self, google_user=None, google_password=None):
        self._google_user = google_user
        self._google_password = google_password

    @property
    def google_user(self):
        return self._google_user or os.environ.get('GOOGLE_USER')

    @property
    def google_password(self):
        return self._google_password or os.environ.get('GOOGLE_PASSWORD')

    @property
    def docs_service(self):
        if not hasattr(self, '_docs_service'):
            service = DocsService(source=SOURCE_NAME)
            service.ClientLogin(self.google_user, self.google_password)
            self._docs_service = service
        return self._docs_service

    @property
    def docs_client(self):
        if not hasattr(self, '_docs_client'):
            client = DocsClient()
            client.ClientLogin(self.google_user, self.google_password,
                               SOURCE_NAME)
            self._docs_client = client
        return self._docs_client

    @property
    def sheets_service(self):
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


class Spreadsheet(object):
    """ A simple wrapper for google docs spreadsheets. Spreadsheets
    only have a ``title`` and an ``id``, all actual data is stored
    in a set of associated ``sheets``. """

    def __init__(self, id, conn):
        self.id = id
        self.conn = conn

    @property
    def meta(self):
        if not hasattr(self, '_wsf') or self._wsf is None:
            self._wsf = self.conn.sheets_service.GetWorksheetsFeed(self.id)
        return self._wsf

    @property
    def title(self):
        return self.meta.title.text

    def __repr__(self):
        return '<Spreadsheet(%r, %r)>' % (self.id, self.title)

    def __unicode__(self):
        return self.title

    @classmethod
    def open(cls, title, conn=None, google_user=None,
             google_password=None):
        """ Open the spreadsheet named ``title``. If no spreadsheet with
        that name exists, a new one will be created. """
        spreadsheet = cls.by_title(title, conn=conn, google_user=google_user,
                                   google_password=google_password)
        if spreadsheet is None:
            spreadsheet = cls.create(title, conn=conn, google_user=google_user,
                                     google_password=google_password)
        return spreadsheet

    @classmethod
    def create(cls, title, conn=None, google_user=None,
               google_password=None):
        """ Create a new spreadsheet with the given name. """
        conn = Connection.connect(conn=conn, google_user=google_user,
                                  google_password=google_password)
        doc = Resource(type='spreadsheet', title=title)
        doc = conn.docs_client.CreateResource(doc)
        id = doc.id.text.rsplit('%3A', 1)[-1]
        return cls.by_id(id, conn=conn)

    @classmethod
    def by_id(cls, id, conn=None, google_user=None,
              google_password=None):
        """ Open a spreadsheet via its resource ID. This is more precise
        than opening a document by title, and should be used with
        preference. """
        conn = Connection.connect(conn=conn, google_user=google_user,
                                  google_password=google_password)
        return cls(id=id, conn=conn)

    @classmethod
    def by_title(cls, title, conn=None, google_user=None,
                 google_password=None):
        conn = Connection.connect(conn=conn, google_user=google_user,
                                  google_password=google_password)
        q = DocumentQuery(categories=['spreadsheet'],
                          text_query=title)
        feed = conn.docs_service.Query(q.ToUri())
        for entry in feed.entry:
            if title == entry.title.text:
                id = entry.feedLink.href.rsplit('%3A', 1)[-1]
                return cls.by_id(id, conn=conn)

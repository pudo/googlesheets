import os
from gdata.docs.data import Resource
from gdata.docs.client import DocsClient, DocsQuery
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


class Sheet(object):
    """ A sheet is a particular table of data within the context
    of a ``Spreadsheet``. """

    def __init__(self, spreadsheet, ws):
        self._ss = spreadsheet
        self._ws = ws
        self.id = ws.id.text.rsplit('/', 1)[-1]

    @property
    def title(self):
        return self._ws.title.text

    def delete(self):
        self._ss.conn.sheets_service.DeleteWorksheet(self._ws)

    def find(self, **kwargs):
        feed = self._ss.conn.sheets_service.GetListFeed(self._ss.id,
                                                        wksht_id=self.id)
        for entry in feed.entry:
            row = {}
            for k, v in entry.custom.items():
                row[k] = v.text
            yield row

    def find_one(self, **kwargs):
        for row in self.find(**kwargs):
            return row

    def __iter__(self):
        return self.find()

    def __repr__(self):
        return '<Sheet(%r, %r, %r)>' % (self._ss, self.id, self.title)

    def __unicode__(self):
        return self.title


class Spreadsheet(object):
    """ A simple wrapper for google docs spreadsheets. Spreadsheets
    only have a ``title`` and an ``id``, all actual data is stored
    in a set of associated ``sheets``. """

    def __init__(self, id, conn, resource=None):
        self.id = id
        self.conn = conn
        self._res = resource

    @property
    def _worksheet_feed(self):
        if not hasattr(self, '_wsf') or self._wsf is None:
            self._wsf = self.conn.sheets_service.GetWorksheetsFeed(self.id)
        return self._wsf

    @property
    def resource(self):
        if not hasattr(self, '_res') or self._res is None:
            self._res = self.conn.docs_client.GetResourceById(self.id)
        return self._res

    @property
    def sheets(self):
        return [Sheet(self, ws) for ws in self._worksheet_feed.entry]

    @property
    def default_sheet(self):
        return self.get('od6')

    def get(self, key, create_missing=True):
        for sheet in self.sheets:
            if key == sheet or sheet.id == key or sheet.title == key:
                return sheet
        if create_missing:
            self.create_sheet(key)

    def __getitem__(self, key):
        return self.get(key)

    def create_sheet(self, title):
        """ Create a sheet with the given title. This does not check if
        another sheet by the same name already exists. """
        ws = self.conn.sheets_service.AddWorksheet(title, 10, 10, self.id)
        self._wsf = None
        return Sheet(self, ws)

    def __iter__(self):
        return self.sheets

    @property
    def title(self):
        return self._worksheet_feed.title.text

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
        """ Create a new spreadsheet with the given ``title``. """
        conn = Connection.connect(conn=conn, google_user=google_user,
                                  google_password=google_password)
        res = Resource(type='spreadsheet', title=title)
        res = conn.docs_client.CreateResource(res)
        id = res.id.text.rsplit('%3A', 1)[-1]
        return cls(id, conn, resource=res)

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
        """ Open the first document with the given ``title`` that is
        returned by document search. """
        conn = Connection.connect(conn=conn, google_user=google_user,
                                  google_password=google_password)
        q = DocsQuery(categories=['spreadsheet'], title=title)
        feed = conn.docs_client.GetResources(q=q)
        for entry in feed.entry:
            if entry.title.text == title:
                id = entry.id.text.rsplit('%3A', 1)[-1]
                return cls.by_id(id, conn=conn)

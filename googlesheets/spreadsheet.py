from gdata.docs.data import Resource
from gdata.docs.client import DocsQuery

from googlesheets.connection import Connection
from googlesheets.sheet import Sheet


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
            return self.create_sheet(key)

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

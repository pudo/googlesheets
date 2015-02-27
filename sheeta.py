import os
import re
import json
from gdata.docs.data import Resource
from gdata.docs.client import DocsClient, DocsQuery
# from gdata.spreadsheets.client import SpreadsheetsClient
from gdata.spreadsheet.service import SpreadsheetsService
from gdata.spreadsheet.service import CellQuery, ListQuery

SOURCE_NAME = 'Sheeta/Python'


def normalize_header(name, existing=[]):
    name = re.sub('\W+', '', name, flags=re.UNICODE).lower()
    # TODO handle multiple columns with the same name.
    return name


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
        self._headers = None

    @property
    def _service(self):
        return self._ss.conn.sheets_service

    def _add_column(self, label, field):
        # Don't call this directly.
        assert self.headers is not None
        cols = max([int(c.cell.col) for c in self._headers])
        new_col = cols + 1
        if int(self._ws.col_count.text) < new_col:
            self._ws.col_count.text = str(new_col)
            self._update_metadata()

        cell = self._service.UpdateCell(1, new_col, label,
                                        self._ss.id, self.id)
        self._headers.append(cell)

    def _create_columns(self, columns):
        columns = {normalize_header(c): c for c in columns}
        # existing = set(self.headers)
        for column in set(columns.keys()).difference(self.headers):
            self._add_column(columns[column], column)
        self._headers = None

    def _convert_value(self, row):
        assert isinstance(row, dict) or hasattr(row, 'items')
        self._create_columns(row.keys())
        data = {}
        for k, v in row.items():
            data[normalize_header(k)] = unicode(v)
        return data

    @property
    def headers(self):
        if self._headers is None:
            query = CellQuery()
            query.max_row = '1'
            feed = self._service.GetCellsFeed(self._ss.id, self.id,
                                              query=query)
            self._headers = feed.entry
        return [normalize_header(h.cell.text) for h in self._headers]

    def _update_metadata(self):
        self._ws = self._service.UpdateWorksheet(self._ws)

    def insert(self, row):
        data = self._convert_value(row)
        self._service.InsertRow(data, self._ss.id, self.id)

    def update(self, row, keys=[]):
        changed = 0
        data = self._convert_value(row)
        keys = [normalize_header(k) for k in keys]
        filters = {k: data.get(k) for k in keys}
        for entry in self._find_entries(**filters):
            row = self._entry_data(entry)
            row.update(data)
            self._service.UpdateRow(entry, row)
            changed += 1
        return changed

    def upsert(self, row, keys=[]):
        if self.update(row, keys=keys) == 0:
            self.insert(row)

    def remove(self, _query=None, **kwargs):
        """ Remove all rows matching the current query. If no query
        is given, this will truncate the entire table. """
        for entry in self._find_entries(_query=_query, **kwargs):
            self._service.DeleteRow(entry)

    @property
    def title(self):
        """ The title of the sheet. """
        return self._ws.title.text

    @title.setter
    def title(self, title):
        self._ws.title.text = unicode(title)
        self._update_metadata()

    def delete(self):
        """ Delete the entire sheet. """
        self._ss.conn.sheets_service.DeleteWorksheet(self._ws)

    def _find_entries(self, _query=None, **kwargs):
        query = None
        if _query is not None:
            query = ListQuery()
            query.sq = _query
        elif len(kwargs.keys()):
            text = []
            for k, v in kwargs.items():
                k = normalize_header(k)
                v = json.dumps(unicode(v))
                text.append("%s = %s" % (k, v))
            query = ListQuery()
            query.sq = ' and '.join(text)
        feed = self._service.GetListFeed(self._ss.id, wksht_id=self.id,
                                         query=query)
        return feed.entry

    def _entry_data(self, entry):
        row = {}
        for k, v in entry.custom.items():
            row[k] = v.text
        return row

    def find(self, _query=None, **kwargs):
        for entry in self._find_entries(**kwargs):
            yield self._entry_data(entry)

    def find_one(self, _query=None, **kwargs):
        for row in self.find(**kwargs):
            return row

    def __iter__(self):
        return self.find()

    def __len__(self):
        # This is not precise.
        return int(self._ws.row_count.text)

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

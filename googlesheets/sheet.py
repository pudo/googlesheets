import json
from gdata.spreadsheet.service import CellQuery, ListQuery

from googlesheets.util import normalize_header


class Sheet(object):
    """ A sheet is a particular table of data within the context
    of a ``Spreadsheet``. Each sheet is comprised of a set of rows
    which can be iterated over, filtered, added, updated and
    removed. The title of the sheet is also mutable. """

    def __init__(self, spreadsheet, ws):
        self._ss = spreadsheet
        self._ws = ws
        self.id = ws.id.text.rsplit('/', 1)[-1]
        self._headers = None

    @property
    def _service(self):
        # Alias for brevity
        return self._ss.conn.sheets_service

    def _add_column(self, label, field):
        """ Add a new column to the table. It will have the header
        text ``label``, but for data inserts and queries, the
        ``field`` name must be used. If necessary, this will expand
        the size of the sheet itself to allow for the new column. """
        # Don't call this directly.
        assert self.headers is not None
        cols = 0
        if len(self._headers) > 0:
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
            if v is None:
                v = ''
            data[normalize_header(k)] = unicode(v)
        return data

    @property
    def headers(self):
        """ Return the name of all headers currently defined for the
        table. """
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
        """ Insert a new row. The row will be added to the end of the
        spreadsheet. Before inserting, the field names in the given 
        row will be normalized and values with empty field names 
        removed. """
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

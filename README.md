# GoogleSheets

``googlesheets`` provides a simple wrapper to read and write data in a Google Spreadsheet in a way that resembles using the [dataset](http://dataset.rtfd.org) library. Each sheet within a spreadsheet can be read and written to using a row-based approach. Headers are stored in the top row and normalized according to Google's API convention.

Please note that [gspread](https://github.com/burnash/gspread) is a more mature library with a more procedural API. If you want to access cells within your spreadsheets directly (ie. by their address), then use that library instead.

## Installation

The easiest way of using ``googlesheets`` is via PyPI:

```bash
$ pip install googlesheets
```

Alternatively, check out the repository from GitHub and install it locally:

```bash
$ git clone https://github.com/pudo/googlesheets.git
$ cd googlesheets
$ python setup.py develop
```

The library depends on Google's own Python service bindings, which will be installed as a dependency. 

## Usage

To use ``googlesheets``, you need to provide valid Google account credentials which have permission to access the spreadsheet you intend to work on. These credentials can be given either through the ``GOOGLE_USER`` and ``GOOGLE_PASSWORD`` environment variables, or by passing ``google_user`` and ``google_password`` directly to the ``Spreadsheet`` classmethods, like ``open``, ``create``, ``by_id`` and ``by_title``.

Here's an example of using the library to access, read and update rows in a Google Spreadsheet:

```python
from googlesheets import Spreadsheet

# This will either create or open a spreadsheet:
spreadsheet = Spreadsheet.open('Things I know')

# You can also choose what to do explicitly:
spreadsheet = Spreadsheet.by_title('Things I know')
spreadsheet = Spreadsheet.create('Things I know')

# Or open a document by it's ID:
spreadsheet = Spreadsheet.by_id('07ej6bb...')

# List all the sheets (tabs) in the given spreadsheet:
for sheet in spreadsheet:
	print sheet, sheet.title
	
# You can access sheets by name (this will create one implicitly
# if the given name does not exist):
sheet = spreadsheet['Sheet 1']

# And you can re-name them:
sheet.title = 'Cheatsheet'

# Next iterate over the data in this sheet:
for row in sheet:
	print row
	
# Or retrieve a filtered version, based on column names and their
# values:
for row in sheet.find(foo='bar'):
	assert row['foo'] == 'bar', row['foo']

# Insert some data:
sheet.insert({'foo': "I'm a banana!", 'bar': "And how."})

# Update and upsert it, based on the fields named in the second
# argument:
sheet.update({'foo': "I'm a banana!", 'bar': "No you're not."}, ['foo'])
sheet.upsert({'foo': "I'm a bean!", 'bar': "Very tasty."}, ['foo'])

# Finally, you can remove rows:
sheet.remove(foo="I'm a bean!")

# Or the entire sheet:
sheet.delete()
```

## License

``googlesheets`` is open source, licensed under a standard MIT license (included in this repository as ``LICENSE``).

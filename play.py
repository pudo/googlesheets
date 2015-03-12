
from googlesheets import Spreadsheet

ss = Spreadsheet.by_title('angaben_a14')
print ss.sheets

s = ss['Sheet 1']
x = ss['Foo']

# x.insert({'foo': 'xxx', 'wahlfisch': 'bla'})

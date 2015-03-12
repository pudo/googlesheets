import re

SOURCE_NAME = 'Python googlesheets'


def normalize_header(name, existing=[]):
    """ Try to emulate the way in which Google does normalization
    on the column names to transform them into headers. """
    name = re.sub('\W+', '', name, flags=re.UNICODE).lower()
    # TODO handle multiple columns with the same name.
    return name

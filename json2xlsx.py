"""Read in list of key-value pairs in JSON format, and output an Excel file with keys as columns (in order
found and corresponding values for a list entry as rows."""

# inspired by https://stackoverflow.com/a/57346253/53897
import datetime
import itertools
import re
import sys
import json
from glob import glob
from json import JSONDecodeError

import xlsxwriter

xlsx_file = sys.argv[1]  # "test.xlsx"

first = False
column_for = dict()

raw_filenames = sys.argv[2:]
filenames = list(itertools.chain.from_iterable([glob(f) for f in raw_filenames]))  # flatten(..)
# filenames = filenames.sort(reverse=True)  # Returns None??

with xlsxwriter.Workbook(xlsx_file, {'default_date_format': 'yyyy-mm-dd hh:mm:ss'}) as workbook:
    worksheet = workbook.add_worksheet()
    row_no = 0
    for filename in filenames:
        print(filename)
        with open(filename) as f:
            try:
                data = json.load(f)
            except JSONDecodeError as e:
                raise RuntimeError(f"For file {filename}") from e

            # Argument to put in the first column, if timestamp convert to date.
            first_field_value = filename
            try:
                if re.search('\\\\(\\d+)[^\\\\]*', first_field_value):
                    first_field_value = datetime.datetime.fromtimestamp(int(re.group(1)))
            except ValueError:
                pass

            if first:
                # Identify a column for every key.
                column_for["_"] = len(column_for)
                for row in data:
                    for key in row:
                        if not key in column_for:
                            column_width = []
                            column_for[key] = len(column_for)

                column_width = [0] * len(column_for)
                #

                if True:  # Is row 0 column names?
                    for title, field_no in column_for.items():
                        worksheet.write(row_no, field_no, title)
                    row_no = row_no + 1
                first = False

            for original_row in data:
                row = dict(original_row)
                row["_"] = first_field_value

                for k, v in row.items():
                    col = column_for[k]
                    if v or len(v):
                        column_width[col] = max(column_width[col], len(str(v)))
                        try:
                            v = int(v)
                        except (ValueError, TypeError) as e:
                            pass
                        # if isinstance(v, datetime):
                        #     worksheet.write_datetime(row_no, col, v)
                        # else:
                        worksheet.write(row_no, col, v)

                row_no = row_no + 1

    # set width of each column to fit the widest value in that column
    for i, width in enumerate(column_width):
        worksheet.set_column(i, i, width)

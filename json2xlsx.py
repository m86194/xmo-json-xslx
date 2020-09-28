"""Read in list of key-value pairs in JSON format, and output an Excel file with keys as columns (in order
found and corresponding values for a list entry as rows."""

# inspired by https://stackoverflow.com/a/57346253/53897
import datetime
import sys
import json
import xlsxwriter

xlsx_file = "test.xlsx"

# Argument to put in the first column, if timestamp convert to date.
first_field_value = sys.argv[2]
try:
    first_field_value = datetime.datetime.fromtimestamp(int(first_field_value))
except ValueError:
    pass

with open(sys.argv[1]) as f:
    data = json.load(f)

# Identify a column for every key.
column_for = dict()
column_for["_"] = len(column_for)
for row in data:
    for key in row:
        if not key in column_for:
            column_width = []
            column_for[key] = len(column_for)

column_width = [0] * len(column_for)
#

# define the excel file
with xlsxwriter.Workbook(xlsx_file, {'default_date_format':
                                         'yyyy-mm-dd hh:mm:ss'}
                         ) as workbook:
    # create a sheet for our work
    worksheet = workbook.add_worksheet()

    row_no = 0
    if True:  # Is row 0 column names?
        for title, field_no in column_for.items():
            worksheet.write(row_no, field_no, title)
        row_no = row_no + 1

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

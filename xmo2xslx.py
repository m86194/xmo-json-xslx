"""Read in list of key-value pairs in JSON format, and output an Excel file with keys as columns (in order
found and corresponding values for a list entry as rows."""

# inspired by https://stackoverflow.com/a/57346253/53897
import datetime
import gzip
import itertools
import os
import re
import sys
from glob import glob

import pytz
import xlsxwriter
from tqdm import tqdm

xlsx_file = sys.argv[1]  # "test.xlsx"
first_row_is_header = True

first_file = True
column_for = dict()

raw_filenames = sys.argv[2:]
filenames = list(itertools.chain.from_iterable([glob(f) for f in raw_filenames]))  # flatten(..)
filenames.sort(reverse=True)  # Datestamped, so newest first

xmo_split_re = re.compile("^ {2}\w+$")
xmo_collect_re = re.compile("^ {4}( *\w+) : (.*)$")


def best_datatype_for(value):
    """The datatype of the value helps xlsxwriter to determine the best Excel representation (making
    timestamps into dates etc.), so convert the strings into whatever is better."""
    if not value:
        return value
    if value == "Thu Jan  1 01:00:00 1970":  # Quick hack for null dates
        return None
    try:
        return int(value)
    except (ValueError, TypeError):
        pass
    try:  # expensive op ahead.  If "Mon Sep 28 08:39:33 2020" convert to tande's clock time
        # First char is first char of Weekday, and last char is last char of year.
        if value[0] in {"M", "T", "W", "F", "S"} and value[-1] in {'0', '1', '2', '3', '4', '5', '6', '7', '8', '9'}:
            # https://docs.python.org/3/library/datetime.html#strftime-and-strptime-format-codes
            utc_datetime = pytz.utc.localize(datetime.datetime.strptime(value, '%a %b %d %H:%M:%S %Y'))
            return utc_datetime.astimezone(pytz.timezone("Etc/GMT-1")).replace(tzinfo=None)
    except (ValueError, TypeError):
        pass
    return value


# https://stackoverflow.com/a/6826099/53897
with open(xlsx_file, "r+") as file_not_locked_by_excel:
    pass

with xlsxwriter.Workbook(xlsx_file, {'default_date_format': 'yyyy-mm-dd hh:mm:ss'}) as workbook:
    worksheet = workbook.add_worksheet()
    row_no = 0

    for filename in tqdm(filenames):
        if os.path.getsize(filename) == 0:  # File system ran full?
            continue

        if filename.endswith(".gz"):
            f_open = gzip.open
        else:
            f_open = open

        # Read output from "xmo-client -p" as list[dict] - currently only single values, not nested
        with f_open(filename, mode="rt") as f:
            entries = dict()
            entries_list = []
            for line_with_lineending in f:
                line = line_with_lineending.rsplit("\n")[0]
                if xmo_split_re.match(line):
                    if entries:
                        entries_list.append(entries)
                        entries = dict()
                else:
                    match = xmo_collect_re.match(line)
                    if match:
                        if entries.get(match.group(1)):
                            raise RuntimeError(">> already in host : " + line)
                        entries[match.group(1)] = match.group(2).lstrip("'").rstrip("'")

            if entries:
                entries_list.append(entries)

            # Argument to put in the first column, if "...\timestamp..." convert to date (Windows only for now).
            first_field_value = filename
            try:
                match = re.search('\\\\(\\d+)[^\\\\]*', first_field_value)
                if match:
                    first_field_value = datetime.datetime.fromtimestamp(int(match.group(1)))
            except ValueError:
                pass

            if first_file:
                if len(entries_list) == 0:
                    raise RuntimeError(f"broken data, no columns found in {filename}")  # bad transfer?
                # Identify a column for every key.
                column_for["_"] = len(column_for)
                for row in entries_list:
                    for key in row:
                        if not key in column_for:
                            column_width = []
                            column_for[key] = len(column_for)

                column_width = [0] * len(column_for)

                if first_row_is_header:  # Is row 0 column names?
                    for title, field_no in column_for.items():
                        worksheet.write(row_no, field_no, title.replace(" ", "_"))
                    row_no = row_no + 1
                first_file = False

            for original_row in entries_list:
                row = dict(original_row)
                row["_"] = first_field_value

                for k, v in row.items():
                    col = column_for.get(k)
                    if col is None:
                        raise KeyError(f"'{k}' in {column_for} for {filename}")
                    if v or len(v):
                        column_width[col] = max(column_width[col], len(str(v)))
                        bv = best_datatype_for(v)
                        worksheet.write(row_no, col, bv)

                row_no = row_no + 1

    # set width of each column to fit the widest value in that column
    for i, width in enumerate(column_width):
        worksheet.set_column(i, i, width)

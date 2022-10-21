#!/usr/bin/env python3

# Copyright since 2022 The spreadsheet-converter developers.
# Use of this source code is governed by a BSD-style license
# that can be found in the LICENSE file at the root of the
# source tree or at https://opensource.org/licenses/BSD-3-Clause.

from typing import List

import pyexcel

import argparse
import os
import sys


def report_error(msg: str):
    """Print the given error message to stderr (with the default prefix).
    Newlines are automatically indented to make individual messages more visible."""
    print("[ERROR]: ", msg.replace("\n", "\n  "), file=sys.stderr)


def filter_empty(book):
    def filter_empty_row_or_col(index, row_or_col):
        del index  # unused
        for entry in row_or_col:
            if type(entry) != str or entry.strip() != "":
                return False
        return True

    for current_sheet_name in book.sheet_names():
        del book[current_sheet_name].row[filter_empty_row_or_col]
        del book[current_sheet_name].column[filter_empty_row_or_col]


def export(book, output_dir: str, output_format: str) -> List[str]:
    if len(book.sheet_names()) == 0:
        report_error("Data source doesn't contain a single sheet")
        sys.exit(1)

    if os.path.exists(output_dir):
        if not os.path.isdir(output_dir):
            report_error(
                "Given output directory \"%s\" is not actually a directory" % output_dir)
            sys.exit(1)
    else:
        os.makedirs(output_dir)

    exported_files = []

    for current_sheet_name in book.sheet_names():
        current_sheet = book[current_sheet_name]

        output_path = os.path.realpath(os.path.join(
            output_dir, current_sheet_name + "." + output_format))

        current_sheet.save_as(filename=output_path)

        exported_files.append(output_path)

    return exported_files


def main():
    parser = argparse.ArgumentParser(
        description="Converter used to convert between different spreadsheet formats")

    parser.add_argument(
        "--input", "-i", help="Path to the input file", metavar="PATH", required=True)
    parser.add_argument(
        "--out_dir", "-o", help="Path to the output directory", metavar="PATH", required=True)
    parser.add_argument("--no-filter-empty",
                        action="store_true", default=False)
    parser.add_argument("--print-exported-files",
                        action="store_true", default=False)
    parser.add_argument("--output-format",
                        help="The desired output format", default="csv")

    args = parser.parse_args()

    book = pyexcel.get_book(file_name=args.input)

    if not args.no_filter_empty:
        filter_empty(book)

    exported_files = export(
        book=book, output_dir=args.out_dir, output_format=args.output_format)

    if args.print_exported_files:
        for current_file in exported_files:
            print(current_file)


if __name__ == "__main__":
    main()

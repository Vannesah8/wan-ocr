"""Build the Booklists reporting sheet from the audited wide extraction."""
import csv
from itertools import chain

import xlsxwriter

from booklists import parse_quantity


def product_value(text):
    quantity, status = parse_quantity(text)
    if status == "blank":
        return None
    return text if quantity is None else quantity


def write_workbook(source, destination):
    options = {"constant_memory": True, "strings_to_formulas": False, "strings_to_urls": False}
    with source.open(encoding="utf-8-sig", newline="") as stream, xlsxwriter.Workbook(destination, options) as book:
        rows = csv.reader(stream)
        headers = next(rows)
        first = next(rows, None)
        sheet = book.add_worksheet("For reporting")
        sheet.freeze_panes(2, 3)
        band = book.add_format({"align": "center", "bold": True})
        heading = book.add_format({"text_wrap": True, "valign": "bottom", "border": 1})
        sheet.merge_range("E1:AK1", "VEDA", band)
        sheet.merge_range("AL1:AV1", "OFFICEPOINT", band)
        sheet.set_column("A:A", 15)
        sheet.set_column("B:B", 16)
        sheet.set_column("C:C", 32)
        sheet.set_column("D:D", 12)
        sheet.set_column("E:AX", 13)
        sheet.set_row(1, 66)
        sheet.write_row(1, 0, headers[:49], heading)
        first_total = sum(value for text in (first or [])[4:49]
                          if isinstance(value := product_value(text), int))
        sheet.write_formula("AX2", "=SUM(E3:AW3)", heading, first_total)
        for index, row in enumerate(chain([first], rows) if first else (), 2):
            if index >= 1048576:
                raise ValueError("Reporting sheet exceeds Excel's row limit.")
            sheet.write_row(index, 0, row[:4])
            sheet.write_row(index, 4, [product_value(text) for text in row[4:49]])

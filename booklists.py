"""Extract Word booklists. Python 3.9+, python-docx>=1.2, pandas>=2."""
import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from contextlib import ExitStack
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table, _Cell


def clean(text):
    return " ".join(text.split())


def key(text):
    return clean(text).casefold()


def parse_quantity(text):
    """Return (individual units, status). None means unresolved."""
    value = key(text)
    if not value:
        return 0, "blank"
    for pattern, multiplier in [
        (r"([0-9]+)(?:\s+(?:pieces?|pcs?|units?))?", 1),
        (r"([0-9]+)\s+dozens?", 12),
    ]:
        match = re.fullmatch(pattern, value)
        if match:
            return int(match[1]) * multiplier, "parsed"
    match = re.fullmatch(r"([0-9]+)\s+packs?\s+(?:of\s+|[x×]\s*)([1-9][0-9]*)(?:\s+(?:pieces?|pcs?|units?))?", value)
    if match:
        return int(match[1]) * int(match[2]), "parsed"
    if re.fullmatch(r"[0-9]+\s+packs?", value):
        return None, "unknown pack size"
    return None, "unresolved quantity"


def load_mapping(path):
    if not path:
        return {}
    data = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    pairs = data[["original_product", "standard_product"]]
    if pairs.original_product.duplicated().any() or any(not clean(v) for v in pairs.standard_product):
        raise ValueError("Mapping needs one nonempty target per unique original header.")
    return dict(pairs.itertuples(index=False, name=None))


def standard_name(original, reviewed, spelling):
    if original in reviewed:
        return reviewed[original]
    if key(original) in {"veda colour pallette", "veda colour palette"}:
        return "Veda Colour Palette"
    if key(original) in {"vedatriangular pencils", "veda triangular pencils"}:
        return "Veda Triangular Pencils"
    return spelling.setdefault(key(original), clean(original))


LOC = ["region_folder", "source_file", "table", "row"]
IDS = ["REGION", "BA NAME", "SCHOOL NAME", "Grade"]
PRODUCTS = ['Veda Extraa Dark Pencils',
 'VedaTriangular Pencils',
 'Veda Sharpeners',
 'Veda Black Erasers',
 'Veda Erasers (ER 40)',
 'Veda Crayons (CR 8A)',
 'Veda Crayons (CR 12A)',
 'Veda Jumbo Crayons',
 'Veda Colour Pencils - Half Size',
 'Veda Colour Pencils - Full size',
 'Veda dual colour pencils',
 'Veda Exercise Books - Ruled',
 'Veda Exercise Books - Squared',
 'Veda Drawing Book (BCR - A4)',
 'Veda Scrap Book ',
 'Veda Modelling Clay MC 1',
 'Veda Water colours (WCC 12)',
 'Veda water colour cake (WCC-12P)',
 'Veda Water colours paint (WCT 01)',
 'Veda colour pallette',
 'Veda artistic brush',
 'Veda Powder Paints',
 'Veda Fluorescent Pad',
 'Veda Clear Glue (GS 10)',
 'Veda Glue Stick (36 g)',
 'Veda Ruler - 30 CM',
 'Veda flexible ruler',
 'Veda Mathematical Set - Mat 1',
 'Veda Mathematical Set - Mat 5',
 'Veda Scientific Calculator - SX 82 MS',
 'Veda Activity Books',
 'Veda Gell Pens',
 'Veda pencil pouch',
 'Office point Pencils (2800)',
 'Office Point ball Pens',
 'Office Point Marker Pens - Red',
 'Office Point Marker Pens - Blue',
 'Office Point Marker Pens - Black ',
 'Office Point Marker Pens - Green ',
 'Office Point display book',
 'Office Point Spring File',
 'Office Point Scissors - SC5/5A',
 'Office Point Scissors',
 'Office Point Masking Tape ',
 'Graph Books']
REQUIRED = ["REGION ", "BA NAME", "SCHOOL NAME ", "Grade", *PRODUCTS, "51"]
CELL = LOC + IDS + ["column", "original_product", "standard_product", "original_text",
                    "quantity_units", "parse_status", "included_in_total"]
NOTE = """Raw rows preserve physical Word cell texts in order as JSON arrays, including headers.
Row, table and physical column numbers start at 1. Tables are in main-body XML order,
including nested tables, whose rows are audited separately. Word formatting is outside CSV.
Merged or offset rows, unexpected headers and widths are flagged for layout review.
No identifier is filled from another row. The folder supplies the reporting region.
A missing school or grade, or missing/conflicting REGION, excludes the row from totals.
Named schools with a matching region count once per folder, even if grade is missing.
School matching ignores case and repeated whitespace only. Duplicate candidates are
listed across files and regions. All eligible submissions contribute to provisional
totals until a reviewer resolves overlaps. No automatic deletion or fuzzy matching.
Whitespace-only product cells count as zero and are logged. Whole numbers and explicit
units count as individual items. Dozens multiply by 12. Packs require an explicit
positive size in the cell. Rates, ranges, unknown pack sizes and other unresolved text
are flagged and excluded. The wide CSV follows the supplied 50-column order. Header mismatches are held for review.
The sample shows a first-row total in AX2. Extra source values in AX are logged separately. Graph Books is counted without a brand.
Empty totals mean no eligible resolved cells for that group.
Product names retain model, colour and brand. Only case, spacing, the specified palette
typo and exact reviewed overrides are combined. Repeated standard columns within a row
are held for review to prevent double counting. Absent columns do not imply zero.
"""


def process(root, output, region=None, limit=None, mapping=None, progress_callback=None):
    reviewed = load_mapping(mapping)
    files = sorted(p for p in root.glob("*/*") if p.is_file() and p.suffix.lower() == ".docx"
                   and not p.name.startswith("~$") and (region is None or p.parent.name == region))
    if limit is not None:
        if limit < 1:
            raise ValueError("Limit must be positive.")
        files = files[:limit]
    if not files:
        raise ValueError("No Word files matched ROOT/REGION/*.docx.")
    output.mkdir(parents=True, exist_ok=False)
    totals, included, excluded = Counter(), Counter(), Counter()
    schools, appearances = defaultdict(set), defaultdict(set)
    observed, spelling = {}, {}
    progress = defaultdict(Counter)
    with ExitStack() as stack:
        def writer(name, columns):
            handle = stack.enter_context((output / name).open("w", newline="", encoding="utf-8-sig"))
            out = csv.writer(handle)
            out.writerow(columns)
            return out.writerow

        raw = writer("raw_rows.csv", LOC + ["cells_json", "grid_before", "grid_after", "spans_json"])
        wide = writer("booklists_wide.csv", REQUIRED)
        long = writer("long.csv", CELL)
        blank = writer("blanks.csv", CELL)
        flag = writer("flagged.csv", LOC + ["column", "original_text", "reason"])
        for index, path in enumerate(files):
            if progress_callback:
                progress_callback(index, len(files), path.relative_to(root).as_posix())
            folder = path.parent.name
            rel = path.relative_to(root).as_posix()
            stat = progress[folder]
            try:
                doc = Document(path)
            except Exception as error:
                flag([folder, rel, "", "", "", "", "read error: " + str(error)])
                stat["read_errors"] += 1
                continue
            stat["files_read"] += 1
            tables = doc.element.body.xpath(".//w:tbl")
            if not tables:
                flag([folder, rel, "", "", "", "", "no body tables"])
                stat["layout_issues"] += 1
            for t, xml in enumerate(tables, 1):
                table = Table(xml, doc)
                headers = None
                names = []
                for r, row in enumerate(table.rows, 1):
                    loc = [folder, rel, t, r]
                    cells = row._tr.tc_lst
                    texts = [_Cell(tc, table).text for tc in cells]
                    spans = [(tc.grid_span, tc.vMerge) for tc in cells]
                    original = json.dumps(texts, ensure_ascii=False)
                    raw(loc + [original, row.grid_cols_before, row.grid_cols_after, json.dumps(spans)])
                    stat["raw_rows"] += 1
                    regular = not (row.grid_cols_before or row.grid_cols_after or
                                   any(span != 1 or merge for span, merge in spans))
                    header_like = [key(v) for v in texts[:4]] == [key(v) for v in IDS]
                    if header_like:
                        headers = None
                        product_end = len(IDS) + len(PRODUCTS)
                        extra_total = len(texts) == product_end + 1 and (
                            not key(texts[-1]) or key(texts[-1]).isdigit())
                        if regular and (len(texts) == product_end or extra_total):
                            originals = texts[4:product_end]
                            names = [standard_name(v, reviewed, spelling) for v in originals]
                            expected = [standard_name(v, reviewed, spelling) for v in PRODUCTS]
                            if list(map(key, names)) == list(map(key, expected)):
                                headers = texts
                                observed.update(zip(originals, names))
                                repeated = {v for v, n in Counter(names).items() if n > 1}
                                continue
                    if not regular or headers is None or len(texts) != len(headers):
                        flag(loc + ["", original, "unrecognised header, merged cells or row width"])
                        stat["rows_flagged"] += 1
                        stat["layout_issues"] += 1
                        continue
                    wide(texts if len(texts) == len(REQUIRED) else [*texts, ""])
                    stat["data_rows"] += 1
                    reg, ba, school, grade = texts[:4]
                    school_key = key(school)
                    if school_key:
                        appearances[school_key].add((folder, rel, school))
                        if key(reg) == key(folder):
                            schools[folder].add(school_key)
                    issues = []
                    if key(reg) != key(folder):
                        issues.append("missing or conflicting REGION")
                    if not school_key or not key(grade):
                        issues.append("missing school or grade")
                    if issues:
                        flag(loc + ["", original, ", ".join(issues)])
                    row_flagged = bool(issues)
                    if len(texts) == len(REQUIRED) and clean(texts[-1]):
                        flag(loc + [len(headers), texts[-1], "extra value outside product columns; check source"])
                        row_flagged = True
                    for col, (header, name, value) in enumerate(zip(headers[4:], names, texts[4:]), 5):
                        qty, status = parse_quantity(value)
                        eligible = not issues and qty is not None and name not in repeated
                        record = loc + texts[:4] + [col, header, name, value, qty, status, eligible]
                        long(record)
                        if status == "blank":
                            blank(record)
                        reasons = []
                        if qty is None:
                            reasons.append(status)
                        if name in repeated:
                            reasons.append("repeated standard product column")
                        if reasons:
                            flag(loc + [col, value, ", ".join(reasons)])
                            row_flagged = True
                        group = folder, name
                        if eligible:
                            totals[group] += qty
                            included[group] += 1
                        else:
                            excluded[group] += 1
                    stat["rows_flagged"] += int(row_flagged)

        products = writer("product_mapping.csv", ["original_product", "standard_product"])
        for pair in sorted(observed.items()):
            products(pair)
        regional = writer("regional_totals.csv", ["region", "standard_product", "known_units",
                                                   "included_cells", "excluded_cells"])
        for group in sorted(included.keys() | excluded.keys()):
            regional([*group, totals[group] if included[group] else None, included[group], excluded[group]])
        counts = writer("school_counts.csv", ["region", "unique_schools"])
        duplicates = writer("duplicate_schools.csv", ["school_key", "region_folder", "source_file", "school_as_written"])
        for school_key, entries in sorted(appearances.items()):
            if len({rel for _, rel, _ in entries}) > 1:
                for entry in sorted(entries):
                    duplicates([school_key, *entry])
        for folder, stat in sorted(progress.items()):
            counts([folder, len(schools[folder])])
            print(f"{folder}: {stat['files_read']} files read, {stat['raw_rows']} raw rows, "
                  f"{stat['data_rows']} data rows extracted, {stat['rows_flagged']} rows flagged, "
                  f"{stat['read_errors']} file read errors, {stat['layout_issues']} layout issues")
    (output / "assumptions.txt").write_text(NOTE, encoding="utf-8")


def self_test():
    samples = {"": (0, "blank"), "  ": (0, "blank"), "0": (0, "parsed"),
               "1 dozen": (12, "parsed"), "2 packs of 6": (12, "parsed"),
               "2 packs": (None, "unknown pack size"), "2 per pupil": (None, "unresolved quantity")}
    for value, expected in samples.items():
        assert parse_quantity(value) == expected
    with TemporaryDirectory() as temp:
        root = Path(temp) / "Booklists"
        out = Path(temp) / "results"
        for i, (region, school, grade, value) in enumerate([
            ("Nakuru", "Lake School", "1", "1 dozen"),
            ("Nakuru", "Lake School", "2", "2 packs of 6"),
            ("Mombasa", "Lake School", "1", "2 per pupil"),
            ("Mombasa", "Port School", "3", "2 packs"),
        ]):
            path = root / region / f"{i}.docx"
            path.parent.mkdir(parents=True, exist_ok=True)
            doc = Document()
            header = "Veda colour pallette" if i % 2 else "Veda Colour Palette"
            products = PRODUCTS.copy()
            products[products.index("Veda colour pallette")] = header
            values_by_name = {header: value, "Office Point Scissors": "  ",
                              "Office Point Scissors - SC5/5A": "3"}
            headings = ["REGION ", "BA NAME", "SCHOOL NAME ", "Grade"] + products + ["51"]
            assert len(headings) == 50
            table = doc.add_table(rows=0, cols=len(headings))
            row_values = [values_by_name.get(name, "1") for name in products]
            for values in [headings, [region, "BA", school, grade, *row_values, "99"]]:
                for cell, text in zip(table.add_row().cells, values):
                    cell.text = text
            doc.save(path)
        process(root, out)

        def read(name):
            return pd.read_csv(out / (name + ".csv"), dtype=str, keep_default_na=False)

        totals = read("regional_totals").set_index(["region", "standard_product"])
        assert totals.loc[("Nakuru", "Veda Colour Palette"), "known_units"] == "24"
        assert totals.loc[("Mombasa", "Veda Colour Palette"), "known_units"] == ""
        assert totals.loc[("Nakuru", "Office Point Scissors"), "known_units"] == "0"
        assert totals.loc[("Nakuru", "Office Point Scissors - SC5/5A"), "known_units"] == "6"
        assert len(read("long")) == 180 and len(read("flagged")) == 6
        assert totals.loc[("Nakuru", "Graph Books"), "known_units"] == "2"
        assert totals.loc[("Mombasa", "Veda Triangular Pencils"), "known_units"] == "2"
        assert "51" not in read("regional_totals").standard_product.values
        assert len(read("blanks")) == 4 and len(read("duplicate_schools")) == 3
        assert dict(read("school_counts").values) == {"Nakuru": "1", "Mombasa": "2"}
        assert list(read("booklists_wide").columns) == REQUIRED
        audited = read("raw_rows")
        assert len(audited) == 8 and json.loads(audited.iloc[1].cells_json)[-1] == "99"
    print("Self-test passed (four synthetic Word files).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("root", type=Path, nargs="?")
    parser.add_argument("output", type=Path, nargs="?", help="New output directory")
    parser.add_argument("--region", help="Exact region folder name")
    parser.add_argument("--limit", type=int, help="Maximum files in this run")
    parser.add_argument("--mapping", type=Path, help="Reviewed product_mapping.csv")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        self_test()
    elif args.root and args.output:
        process(args.root, args.output, args.region, args.limit, args.mapping)
    else:
        parser.error("Supply ROOT OUTPUT or --self-test")

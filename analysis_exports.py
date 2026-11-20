"""Build a shareable summary or the full audit bundle from a completed run."""
import json
import os
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document

WORD = "analysis-summary.docx"
ZIP = "analysis-bucket.zip"


def make_summary(summary, destination):
    document = Document()
    document.add_heading(summary.get("project", "Booklists"), 0)
    document.add_heading("Booklist analysis", level=1)
    document.add_paragraph(f"Completed {summary['date']} · {summary['files']} source files")
    document.add_paragraph(f"{summary['flags']} review flags · {summary['duplicates']} potential duplicate school names")
    document.add_paragraph("Quantities are provisional. Unresolved cells are excluded. "
                           "Duplicate submissions may contribute to totals until reviewed.")
    totals = summary["totals"]
    for school in summary["schools"]:
        region = school["region"]
        document.add_heading(region, level=2)
        document.add_paragraph(f"Unique schools: {school['unique_schools']}")
        table = document.add_table(rows=0, cols=3)
        table.style = "Light Shading Accent 1"
        rows = [("Product", "Known units", "Excluded cells")]
        rows.extend((r["standard_product"], r["known_units"] or "Unresolved", r["excluded_cells"])
                    for r in totals if r["region"] == region)
        for values in rows:
            for cell, value in zip(table.add_row().cells, values):
                cell.text = value
    document.save(destination)


def export(job, name, reports):
    """Write on disk so large ZIP downloads need no in-memory copy."""
    output = job / "reports"
    target = output / name
    if target.exists():
        return target
    with tempfile.NamedTemporaryFile(dir=output, delete=False) as temp:
        staged = Path(temp.name)
    try:
        if name == WORD:
            make_summary(json.loads((job / "summary.json").read_text(encoding="utf-8")), staged)
        elif name == ZIP:
            word = export(job, WORD, reports)
            with ZipFile(staged, "w", ZIP_DEFLATED) as archive:
                for path in [word, *(output / item for item in reports)]:
                    archive.write(path, path.name)
        else:
            raise ValueError("Unknown export.")
        os.replace(staged, target)
        return target
    finally:
        staged.unlink(missing_ok=True)

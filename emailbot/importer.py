"""Reads recipient lists from CSV / Excel files or text pasted from a spreadsheet."""

import csv
import io
import re
from datetime import date, datetime

def _long_date(d):
    # e.g. "March 5, 2026". (strftime's "%-d" for an unpadded day doesn't work on Windows.)
    return f"{d:%B} {d.day}, {d.year}"


def _cell(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, datetime):
        return _long_date(value) if value.time() == datetime.min.time() else str(value)
    if isinstance(value, date):
        return _long_date(value)
    return str(value).strip()


def _table_from_text(text, delimiter=None):
    text = text.lstrip("\ufeff")
    if delimiter is None:
        sample = text[:5000]
        if "\t" in sample:
            delimiter = "\t"
        else:
            try:
                delimiter = csv.Sniffer().sniff(sample, delimiters=",;").delimiter
            except csv.Error:
                delimiter = ","
    return list(csv.reader(io.StringIO(text), delimiter=delimiter))


def _table_from_xlsx(data):
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise ValueError(
            "Opening Excel files isn't available because part of Email Bot's setup is missing. "
            "Use “Paste from Excel” instead, or save the file as CSV and open that."
        ) from None

    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _kind(header):
    h = header.lower().replace("-", "").replace("_", " ").strip()
    if re.search(r"\bcc\b", h) or h in ("copy to", "carbon copy"):
        return "cc"
    if "email" in h or h in ("to", "address"):
        return "email"
    return None


def _decode(data):
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace")


def parse(filename=None, data=None, text=None):
    """Read a spreadsheet into {"headers", "rows", "email_col", "cc_col"}.

    Rows are lists of strings lined up with headers. email_col / cc_col are the
    best guesses at those columns (or None); the person importing confirms them.
    """
    name = (filename or "").lower()
    if text is not None:
        table = _table_from_text(text)
    elif name.endswith((".xlsx", ".xlsm")):
        table = _table_from_xlsx(data)
    elif name.endswith((".tsv", ".tab")):
        table = _table_from_text(_decode(data or b""), delimiter="\t")
    elif name.endswith(".csv"):
        decoded = _decode(data or b"")
        try:
            delimiter = csv.Sniffer().sniff(decoded[:5000], delimiters=",;").delimiter
        except csv.Error:
            delimiter = ","
        table = _table_from_text(decoded, delimiter=delimiter)
    else:
        table = _table_from_text(_decode(data or b""))

    table = [[_cell(c) for c in row] for row in table]
    table = [row for row in table if any(row)]
    if not table:
        raise ValueError("The file or pasted text didn't contain any rows.")
    if len(table) < 2:
        raise ValueError("Only one row was found. Make sure you include the title row and at least one person.")

    width = max(len(r) for r in table)
    headers = table[0] + [""] * (width - len(table[0]))
    headers = [h or f"Column {i + 1}" for i, h in enumerate(headers)]
    # Duplicate titles would overwrite each other, so number the repeats.
    seen = {}
    for i, h in enumerate(headers):
        key = h.lower()
        if key in seen:
            seen[key] += 1
            headers[i] = f"{h} ({seen[key]})"
        else:
            seen[key] = 1
    rows = [r + [""] * (width - len(r)) for r in table[1:]]

    email_col = next((i for i, h in enumerate(headers) if _kind(h) == "email"), None)
    cc_col = next((i for i, h in enumerate(headers) if _kind(h) == "cc" and i != email_col), None)
    if email_col is None:
        # No title with "email" in it: pick the column with the most addresses.
        counts = [sum("@" in r[i] for r in rows) for i in range(width)]
        if max(counts) > 0:
            email_col = counts.index(max(counts))

    return {"headers": headers, "rows": rows, "email_col": email_col, "cc_col": cc_col}

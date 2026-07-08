"""Load and validate the distribution list, the table that says who gets a
report, what they may see, and how it reaches them.

Source format: CSV or .xlsx. Columns (header row required, order-independent):

    RecipientName   Used for the output file name and the {name} placeholder.
    AccessKey       Comma-separated EmployeeIDs this person may see. EMPTY = all
                    (no restriction). One value = a single person. Several = a
                    manager who sees those people.
    Email           Outlook address. Blank to skip emailing this recipient.
    OneDriveFolder  Destination folder for a copied file. Blank to skip.
    Enabled         TRUE/FALSE (default TRUE). FALSE keeps the row but skips it.

The access key is parsed the same way the Power Query AllowedIds query parses it,
so what the preview shows is what the workbook will actually filter to:
    ""      -> ALL          (unrestricted)
    "3"     -> {3}
    "3, 4"  -> {3, 4}       (whitespace tolerated)
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

REQUIRED_COLUMNS = {"RecipientName", "AccessKey"}
KNOWN_COLUMNS = REQUIRED_COLUMNS | {"Email", "OneDriveFolder", "Enabled"}


class RecipientsError(Exception):
    """Raised for a malformed or invalid distribution list."""


@dataclass(frozen=True)
class Recipient:
    name: str
    access_key: str                 # raw string exactly as written (tool writes this to the cell)
    email: str
    onedrive_folder: str
    enabled: bool

    @property
    def unrestricted(self) -> bool:
        return self.access_key.strip() == ""

    @property
    def allowed_ids(self) -> list[int] | None:
        """The parsed id set, or None when unrestricted. Mirrors AllowedIds.m."""
        if self.unrestricted:
            return None
        ids: list[int] = []
        for part in self.access_key.split(","):
            part = part.strip()
            if part == "":
                continue
            try:
                ids.append(int(part))
            except ValueError as exc:
                raise RecipientsError(
                    f"Recipient '{self.name}': access key '{self.access_key}' contains a "
                    f"non-numeric id ('{part}'). Use comma-separated whole numbers, e.g. 3,4."
                ) from exc
        return ids

    def scope_label(self) -> str:
        """Human-readable scope for the preview (so you can spot a missing id)."""
        return "ALL (unrestricted)" if self.unrestricted else ",".join(str(i) for i in self.allowed_ids)


def _to_bool(value: str) -> bool:
    return str(value).strip().lower() in ("true", "1", "yes", "y")


def _rows_from_csv(path: Path) -> list[dict]:
    with path.open("r", newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            raise RecipientsError(f"{path} has no header row.")
        _check_columns(set(reader.fieldnames), path)
        return [dict(r) for r in reader]


def _rows_from_xlsx(path: Path, sheet: str) -> list[dict]:
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(str(path), update_links=False, read_only=True)
        try:
            try:
                ws = wb.sheets[sheet]
            except Exception as exc:
                raise RecipientsError(f"Sheet '{sheet}' not found in {path}.") from exc
            grid = ws.used_range.value
        finally:
            wb.close()
    finally:
        app.quit()

    if not grid or not isinstance(grid, list):
        raise RecipientsError(f"Sheet '{sheet}' in {path} is empty.")
    if not isinstance(grid[0], list):        # a single row comes back flat
        grid = [grid]
    header = [("" if h is None else str(h).strip()) for h in grid[0]]
    _check_columns(set(header), path)
    rows: list[dict] = []
    for raw in grid[1:]:
        cells = raw if isinstance(raw, list) else [raw]
        if all(c is None or str(c).strip() == "" for c in cells):
            continue                          # skip fully blank rows
        rows.append({header[i]: cells[i] if i < len(cells) else None for i in range(len(header))})
    return rows


def _check_columns(found: set[str], path: Path) -> None:
    missing = REQUIRED_COLUMNS - found
    if missing:
        raise RecipientsError(
            f"{path} is missing required column(s): {', '.join(sorted(missing))}. "
            f"Expected at least {', '.join(sorted(REQUIRED_COLUMNS))}."
        )


def _cell(row: dict, key: str) -> str:
    value = row.get(key)
    if value is None:
        return ""
    # An Excel numeric id like 3.0 -> "3"; leave text keys ("3,4") untouched.
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def load_recipients(path: Path, sheet: str) -> list[Recipient]:
    if not path.is_file():
        raise RecipientsError(f"Recipients list not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        rows = _rows_from_csv(path)
    elif suffix in (".xlsx", ".xlsm"):
        rows = _rows_from_xlsx(path, sheet)
    else:
        raise RecipientsError(f"Unsupported recipients format '{suffix}'. Use .csv or .xlsx.")

    recipients: list[Recipient] = []
    seen_names: set[str] = set()
    for i, row in enumerate(rows, start=2):        # start=2: row 1 is the header
        name = _cell(row, "RecipientName")
        if name == "":
            raise RecipientsError(f"Row {i}: RecipientName is blank.")
        if name.lower() in seen_names:
            raise RecipientsError(f"Row {i}: duplicate RecipientName '{name}'. Names must be unique (they name the output file).")
        seen_names.add(name.lower())

        enabled_raw = row.get("Enabled")
        enabled = True if enabled_raw is None or str(enabled_raw).strip() == "" else _to_bool(enabled_raw)

        recipient = Recipient(
            name=name,
            access_key=_cell(row, "AccessKey"),
            email=_cell(row, "Email"),
            onedrive_folder=_cell(row, "OneDriveFolder"),
            enabled=enabled,
        )
        _ = recipient.allowed_ids                   # validate the key parses now, not mid-run
        recipients.append(recipient)

    if not recipients:
        raise RecipientsError(f"{path} has a header but no recipient rows.")
    return recipients

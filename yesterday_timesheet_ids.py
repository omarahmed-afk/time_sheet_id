"""Export previous days' Deputy timesheets, including Friday-Sunday on Mondays."""

import argparse
import csv
import json
import os
import tempfile
from pathlib import Path

import gspread
import requests

from Calculate_time import query_resource, target_range, to_ny

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "google_sheets_config.json"


def write_csv(output, rows):
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        file = output.open("w", newline="", encoding="utf-8-sig")
    except PermissionError:
        file = tempfile.NamedTemporaryFile(
            mode="w", newline="", encoding="utf-8-sig", delete=False,
            dir=output.parent, prefix=f"{output.stem}_", suffix=output.suffix,
        )
        print(f"Cannot overwrite {output}; saving a separate CSV instead.")
    with file:
        csv.writer(file).writerows(rows)
    return Path(file.name)


def google_destination(start_date):
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig")) if CONFIG_PATH.is_file() else {}
    url = os.environ.get("GOOGLE_SPREADSHEET_URL", config.get("spreadsheet_url", "")).strip()
    credentials_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not url:
        raise RuntimeError(
            "Set GOOGLE_SPREADSHEET_URL or spreadsheet_url in google_sheets_config.json. "
            "Use --csv-only for CSV only."
        )
    if not url or "YOUR_SHEET_ID" in url:
        raise RuntimeError("Set spreadsheet_url in google_sheets_config.json.")
    if credentials_json:
        client = gspread.service_account_from_dict(json.loads(credentials_json))
    else:
        credentials = Path(config.get("credentials_file") or "service_account.json")
        if not credentials.is_absolute():
            credentials = BASE_DIR / credentials
        if not credentials.is_file():
            raise RuntimeError(f"Google service account JSON not found: {credentials}")
        client = gspread.service_account(filename=str(credentials))
    spreadsheet = client.open_by_url(url)
    title = "Timesheet ID"
    return spreadsheet, title


def write_google_sheet(spreadsheet, title, rows):
    try:
        worksheet = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=title, rows=max(100, len(rows)), cols=3)
    old_rows = len(worksheet.get("A:C"))
    values = rows + [["", "", ""] for _ in range(max(0, old_rows - len(rows)))]
    if len(values) > worksheet.row_count:
        worksheet.resize(rows=len(values))
    worksheet.update(values=values, range_name=f"A1:C{len(values)}", value_input_option="RAW")
    print(f"Google Sheet updated: {spreadsheet.url} (tab: {title})")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv-only", action="store_true", help="Save CSV without writing to Google Sheets")
    args = parser.parse_args()
    start_date, end_date, start, end = target_range(None)
    destination = None if args.csv_only else google_destination(start_date)
    timesheets = query_resource("Timesheet", {"search": {
        "start": {"field": "StartTime", "data": start, "type": "ge"},
        "end": {"field": "StartTime", "data": end, "type": "le"},
    }})

    def metadata_name(timesheet):
        return ((timesheet.get("_DPMetaData") or {}).get("EmployeeInfo") or {}).get("DisplayName", "")

    missing_ids = sorted({t["Employee"] for t in timesheets
                          if not metadata_name(t) and t.get("Employee")})
    names = {}
    for offset in range(0, len(missing_ids), 200):
        employees = query_resource("Employee", {"search": {
            "ids": {"field": "Id", "data": missing_ids[offset:offset + 200], "type": "in"},
        }})
        for employee in employees:
            names[employee["Id"]] = employee.get("DisplayName") or " ".join(
                str(employee.get(key) or "") for key in ("FirstName", "LastName")).strip()

    rows = [["timesheet_id", "name", "date"]]
    for timesheet in sorted(timesheets, key=lambda t: (t.get("StartTime", 0), t.get("Id", 0))):
        rows.append([
            timesheet.get("Id"),
            metadata_name(timesheet) or names.get(timesheet.get("Employee"), ""),
            to_ny(timesheet["StartTime"]).date().isoformat(),
        ])
    date_label = str(start_date) if start_date == end_date else f"{start_date}_to_{end_date}"
    output = BASE_DIR / "reports" / f"yesterday_{date_label}_timesheet_ids.csv"
    output = write_csv(output, rows)
    print(f"Saved {len(timesheets)} timesheets for {date_label} (New York time): {output}")
    if destination:
        write_google_sheet(*destination, rows)


if __name__ == "__main__":
    try:
        main()
    except (requests.RequestException, RuntimeError, ValueError, OSError, gspread.exceptions.GSpreadException) as exc:
        raise SystemExit(f"Export failed: {exc}") from None

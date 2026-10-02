# Daily Deputy timesheet export

GitHub Actions runs `yesterday_timesheet_ids.py` when dispatched manually or by your external cron job. There is no built-in schedule. Configure your cron job to trigger Monday through Friday using America/New_York time, skipping Saturday and Sunday. Monday exports the previous **Friday, Saturday and Sunday** together; other days export the previous New York calendar day. It writes timesheet IDs, names and dates to the **Timesheet ID** Google Sheets tab, replacing that tab's previous export. A CSV is also saved as a workflow artifact for seven days. This workflow does not approve timesheets.

## GitHub setup

1. Create a **private** GitHub repository and upload this project's source files, `requirements.txt` and `.github/workflows/daily-timesheets.yml`. Keep the workflow on the default branch. Do not upload credential JSON files, `google_sheets_config.json`, browser profiles or reports. The `.gitignore` protects these when uploading through Git; browser uploads require selecting files carefully.
2. Under **Settings → Secrets and variables → Actions → New repository secret**, add:

   | Secret | Value |
   | --- | --- |
   | `DEPUTY_TOKEN` | A valid Deputy access token for your installation |
   | `GOOGLE_SPREADSHEET_URL` | The spreadsheet URL from your local `google_sheets_config.json` |
   | `GOOGLE_SERVICE_ACCOUNT_JSON` | The entire contents of your Google service account JSON file, including the braces |

3. Share the spreadsheet with the service account's `client_email` as **Editor**. Enable the Google Sheets and Google Drive APIs in its Google Cloud project if needed.
4. Open **Actions → Export yesterday's timesheets → Run workflow** to test. Confirm that the **Timesheet ID** tab updates and download the CSV under the run's **Artifacts** section.

Your external cron job should dispatch `daily-timesheets.yml` on branch `main`. Choose the run time in your cron service; this repository does not set one. Manual runs remain available through **Run workflow**.

The Deputy token must remain valid for unattended runs. If it expires, replace the `DEPUTY_TOKEN` secret; this project does not refresh tokens. Rotate the token previously embedded in `Calculate_time.py` before publishing or enabling automation.

## Local use (PowerShell)

```powershell
python -m pip install -r requirements.txt
$env:DEPUTY_TOKEN = 'your-token'
python yesterday_timesheet_ids.py
```

Local Google credentials can still be configured using `google_sheets_config.json` with `spreadsheet_url` and `credentials_file`. Alternatively set the same Google environment variables used by the workflow. Use `--csv-only` to skip Google Sheets.

`Calculate_time.py` and `september_timesheet_ids.py` also use `DEPUTY_TOKEN` now. The daily workflow runs only the yesterday export. `Calculate_time.py` retains its separate preview and explicit `--approve` behavior.

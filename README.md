# LME Trading Calendar + Daily Cash / 3M Date Excel

Python toolkit that builds and daily-refreshes an Excel workbook of LME
Settlement Business Days, plus that day's **Cash Date** and **3M Date**.

No VBA. Safe to run from cron or Windows Task Scheduler.

## Workbook layout (stable for downstream readers)

| Sheet | Purpose |
|---|---|
| `Trading_Calendar` | Every calendar day from the snapshot date through +27 months, with UK/US holiday flags and `Is_Valid_Trading_Day` |
| `Cash_3M_Daily` | Snapshot of `Snapshot_Date`, `Cash_Date`, `3M_Date`, `Days_Cash_to_3M` |
| `Holiday_Reference` | UK / US / Both bank holidays used by the calculation |

Column names and sheet names are part of the contract for the LME quoting-board
Word report. Do not rename them.

Named ranges always point at the **current** snapshot row: `Snapshot_Date`,
`Cash_Date`, `ThreeM_Date`, `Days_Cash_to_3M`.

The updater **only** rewrites those three sheets. Other sheets in the same
workbook (for example Bloomberg `BDP`/`BDH` tabs) are left untouched.

## Prompt-date rules

- **Cash Date** = 2nd Settlement Business Day after the trade date (T+2).
- **3M Date** = trade date + 3 calendar months, then LME Rulebook 8.4.1 / 8.4.2:
  - default: next Settlement Business Day
  - Saturday (and Friday is a SBD) → previous SBD (usually Friday)
  - Good Friday → previous SBD
  - Christmas Day on Tue–Fri → previous SBD
  - if the result would fall in the 4th calendar month after the trade month →
    last SBD of the 3rd calendar month
- **Settlement Business Day** = weekday that is not a UK bank holiday and not a
  US bank holiday. Either region's holiday invalidates the prompt.

Reference check (also executed as a unit test and logged on every run):

| Trade Date | Cash Date | 3M Date |
|---|---|---|
| 2026-08-07 (Fri) | 2026-08-11 (Tue) | 2026-11-06 (Fri) |

Unadjusted 3M target 2026-11-07 is Saturday, so 8.4.1(a) rolls **back** to
2026-11-06, not forward to 2026-11-09.

## Setup

```bash
python3 -m pip install -r requirements.txt
```

## Run once (creates / refreshes the workbook)

```bash
python3 update_lme_calendar.py
```

Defaults:

- `--excel output/lme_prompt_calendar.xlsx`
- `--holidays data/holidays.csv`
- `--as-of` today
- `--months 27`
- `--snapshot-mode overwrite`

Useful flags:

```bash
# Accumulate daily Cash/3M history instead of keeping a single row
python3 update_lme_calendar.py --snapshot-mode append

# Keep rows that have rolled out of the 27-month window
python3 update_lme_calendar.py --retain-expired

# Pin the trade date (useful for backfills / tests)
python3 update_lme_calendar.py --as-of 2026-08-19 --snapshot-mode overwrite

# Write a log file as well as stdout
python3 update_lme_calendar.py --log-file logs/lme_calendar.log
```

On success the script prints today's Cash Date and 3M Date, and re-runs the
2026-08-07 validation anchor.

## Overwrite vs append

| `--snapshot-mode` | `Cash_3M_Daily` behaviour |
|---|---|
| `overwrite` (default) | Sheet always has header + **one** data row (today's snapshot). Named ranges point at row 2. |
| `append` | Each new snapshot date is appended. Re-running on the same date replaces that day's row instead of duplicating it. Named ranges point at the latest row. |

Filter `Trading_Calendar` on `Is_Valid_Trading_Day = TRUE` to get the valid
prompt-date list.

## Update the holiday list

The script **will not invent holidays**. If `--holidays` is missing or the CSV
is malformed it exits with a description of the required format.

Required columns:

```text
Date,Region,Holiday_Name
2026-08-31,UK,Summer bank holiday
2026-11-11,US,Veterans Day
2026-12-25,Both,Christmas Day
```

`Region` must be `UK`, `US`, or `Both`.

Preferred sources (in order):

1. Bloomberg Terminal `CDR <GO>` London calendar export, mapped to the three columns above.
2. Latest LME notice *Banking Holidays, Non-Valid Prompt Dates and LME Business Days*.
3. The shipped `data/holidays.csv` (gov.uk + US bank holidays, cross-checked against LME Notice 26/131). See `data/README.md`.

If Cash Date or 3M Date falls outside the min/max dates in the holiday file, a
`HolidayCoverageWarning` is emitted. Extend the CSV before trusting that run.

JPY-only / EUR-only banking holidays are **not** in this file: they do not
invalidate USD/GBP LME prompt dates.

## Scheduler

### cron (Linux / macOS)

```cron
# Every weekday at 06:15 local time
15 6 * * 1-5 cd /path/to/fantastic-octo-palm-tree && /usr/bin/python3 update_lme_calendar.py --snapshot-mode append --log-file logs/lme_calendar.log
```

### Windows Task Scheduler

1. Action: `python.exe`
2. Arguments: `update_lme_calendar.py --snapshot-mode append --log-file logs\lme_calendar.log`
3. Start in: the repository root
4. Trigger: daily, weekdays, after London midnight (so "today" is the new trade date)

Exit codes: `0` success, `2` holiday-file problem, `1` any other failure.

## Tests

```bash
python3 -m pytest -q
```

## Layout

```text
update_lme_calendar.py      scheduler entry point
lme_calendar/dates.py       is_business_day, calc_cash_date, calc_3m_date, generate_trading_calendar
lme_calendar/holidays.py    CSV loader (fail closed)
lme_calendar/excel.py       writes only the three managed sheets
data/holidays.csv           UK+US holidays 2026–2029
output/lme_prompt_calendar.xlsx
tests/
```

# Naming patterns observed in the Box end directories

Observed read-only on 2026-11-02. Nothing in Box or the project root was modified.
Box root on this PC: `C:\Users\AhmedElrefai\Box`. Other usernames seen in old notes and
scripts: `ZiadAhmed`, `ZiadMostafaAhmed`. Only the `C:\Users\<name>\` part differs.

## 1. INPUTS (raw daily exports, dropped by hand)

### Not Ready (daily XLSX, one per day)
`Box\Agent Not Ready 2025\2026\<month folder>\`
(the parent is called "2025" but holds a `2025` and a `2026` sub-folder)

| Month | Folder name seen |
|---|---|
| Jan, Feb | `Jan 2026`, `Feb 2026` (abbreviated) |
| Mar - Sep | `March 2026`, `April 2026`, `May 2026`, `June 2026`, `July 2026`, `August 2026`, `September 2026` (full) |
| Oct | `Oct 2026` (abbreviated again) |

2025 folders are even looser: `April 2025`, `FEB 2025`, `Jan 2025`, `NOV 2025`,
`December 2025 Not Ready`, `October 2025.old`.

Daily file names inside change over time:
- Jan-Mar: `<D> <Month>.xlsx` (`1 Jan.xlsx`, `1 March.xlsx`)
- April: `April <D>.xlsx`
- May onward: `<D>.xlsx` (`1.xlsx` ... `31.xlsx`)

The pipeline creates `<month folder>\cleaned\` with `<D>_clean.xlsx` files inside.

### Login / Logout (daily XLSX, one per day)
`Box\Agent Not Ready 2025\2026\<Mon> Login-LogOut 2026\`
Seen: `March`, `April`, `May`, `June`, `July`, `August`, and `Sept` (abbreviated).
Files: `1.xlsx` ... `31.xlsx`.

### Utilization raw (two monthly exports, same folder)
`Box\Utilization (Tracker)\`
- `<Mon> Ready time Report.xlsx` (seen: `Sep Ready time Report.xlsx`)
- `<Mon> Enage Time Report.xlsx` (seen: `Sep Enage Time Report.xlsx`; "Enage" is a typo that is part of the real name)

Only the latest month's raws are kept there. The scripts tell Ready from Engage by header,
not by filename.

## 2. OUTPUTS (the "end" directories)

| Series | Folder | File pattern | Example |
|---|---|---|---|
| Login | `Box\Not Ready tracker\Login Logout\` | `Master_<FullMonth>_Login.xlsx` | `Master_September_Login.xlsx` |
| Not Ready | `Box\Not Ready tracker\Not Ready\` | `MASTER_<FullMonth>.xlsx` (no year) | `MASTER_September.xlsx` |
| Ready | `Box\Utilization (Tracker)\Agent Summary Activity Report (Active)\` | `MASTER_<FullMonth>_<YYYY>_Ready.xlsx` | `MASTER_September_2026_Ready.xlsx` |
| Engage | `Box\Utilization (Tracker)\Agent Utilization Report (Total Enage Time)\` | `MASTER_<FullMonth>_<YYYY>_Engage.xlsx` | `MASTER_September_2026_Engage.xlsx` |

Notes
- Output month names are full names. The one exception is `MASTER_Feb.xlsx`.
- `Master_` (Login) vs `MASTER_` (the others): the case differs, keep it per series.
- Not Ready has no year in the real files. `report_pipeline.py` would default to
  `MASTER_<Month>_<Year>.xlsx`, so `--output` must be passed to match the real convention.
- Folder names must be reproduced exactly: `Enage` typo, parentheses, `Not Ready tracker`
  (lowercase t), `Login Logout`.

## 3. Cycle
Each month is processed in the first days of the next month (output files dated the 1st-5th).
Today is 2026-11-02, so the default target month = previous month = October 2026.
October is NOT ready to run yet:
- `Oct 2026` has 1 file (`26.xlsx`)
- there is no `Oct Login-LogOut 2026` folder
- the Utilization root still holds only the September raws

## 4. Script command shapes (from `How To Refai.txt`, unchanged)
```
python Login_time.py <in_dir> <out_dir>
python report_pipeline.py  --folder <month dir> --output <MASTER_<Month>.xlsx> --year <YYYY>
python clean_ready_time.py  --folder <Utilization root> --outdir <Ready out dir>   [--month YYYY-MM]
python clean_engage_time.py --folder <Utilization root> --outdir <Engage out dir>  --month YYYY-MM
```
Not in the workflow (legacy): `Grouping.py` (hard-coded `C:\Users\ZiadAhmed\Box\April`),
`utlization.py` (older cleaner, superseded by `report_pipeline.py`).

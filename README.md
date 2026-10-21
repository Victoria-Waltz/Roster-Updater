# TAIT Roster Updater

This is a local Python application for comparing Sven's monthly active-user file with the approved employee baseline. Employee data stays in the folder selected by the user. The application does not upload data to an external service.

## What it produces

The application compares employees by `Logon Name` and creates one Excel report containing:

- Summary counts
- Items requiring review
- Field-by-field employee changes
- The updated employee roster

It identifies new employees, employees missing from the new pull, title and reporting changes, missing capabilities, manager capability suggestions, incomplete manager information, and managers who are not present in the active pull.

The updated roster carries forward an existing capability when Sven's file leaves it blank. Approved capability corrections are saved separately and reused in later months. A manager's capability is shown only as a suggestion until a user approves it.

## First-time setup

1. Put the app on each coworker's Windows computer.
2. Create or choose one shared folder that every user can access. A company network drive or a company-approved OneDrive/SharePoint-synced folder works.
3. Open the app and click **Choose folder**. Every user must choose the same shared folder.
4. One person clicks **Initialize baseline** and selects the approved workbook containing `All_Users_Active_Info (1)`.

The shared folder will contain:

- `current_baseline.xlsx` — the official approved baseline
- `capability_corrections.json` — approved capability corrections
- `history` — dated baseline backups and copies of approved source pulls
- `reports` — reports saved automatically when a month is approved
- `approval_history.csv` — who approved each update and when

## Monthly process

1. Open the app.
2. Click **Choose file** and select Sven's new `.xlsx` or `.csv` file.
3. Click **Run update**.
4. Review the colored items. Select a row to edit a capability or apply the manager suggestion.
5. Click **Export report** to share the results with Amanda.
6. When the month is ready, click **Approve as new baseline**.

Approval automatically saves a dated backup. If another user changes the baseline after someone runs a comparison, the app prevents the older comparison from being approved until it is rerun.

## Run from Python

The easiest source-code setup on Windows is to double-click `run_app.bat`. It creates a local Python environment, installs `openpyxl`, and starts the application.

Python 3.11 or 3.12 is recommended. Standard Windows Python includes Tkinter, which the desktop window uses.

## Build a Windows application

To make a single Windows `.exe`, run `build_windows_app.bat` on a Windows computer. The finished file will be created in `dist`.

Windows or company security may show a warning for an unsigned internal application. If the app will be distributed broadly, have IT review and sign the executable.

## Expected columns

The monthly pull must contain these columns. Column order does not matter.

1. S.No.
2. Display Name
3. First Name
4. Last Name
5. Logon Name
6. Email Address
7. Title
8. Department
9. Capability
10. Manager
11. Manager Email Address
12. Office
13. City
14. Employee Type

The baseline may also include `ERP Category`. The app carries this value forward for existing employees.

## Important behavior

- `Logon Name` is the unique employee key. Duplicate logon names stop the update.
- The app never automatically applies a manager capability suggestion.
- A manager is flagged as potentially inactive when the manager cannot be found in the new active-user pull.
- The application can still approve a month with unresolved items, but it warns the user first.
- Do not place the shared data folder inside the application folder. Keep it in a company-approved shared location.


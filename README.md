# TAIT Roster Updater

## Code organization

New to the code? Read the [code walkthrough](docs/CODE_WALKTHROUGH.md) for terminology, a complete monthly-update example, and guidance on where to make changes. Python docstrings explain the functions and classes; inline comments explain the key processing decisions.

Application code lives in `roster_updater/`, with desktop interface code in `roster_updater/ui/`, development tools in `scripts/`, and regression tests in `tests/`. See [the project structure guide](docs/PROJECT_STRUCTURE.md) for a file-by-file map and development commands. The existing Windows launch and build scripts still work from the repository root.

This is a local Python application for comparing Sven's monthly active-user file with the approved employee baseline. Employee data stays in the folder selected by the user. The application does not upload data to an external service.

## What it produces

The export is **one workbook with four tabs**:

| Tab | Purpose |
| --- | --- |
| Summary | Employee, department, and capability counts; review-category color legend; workbook instructions |
| Review | Only employees with changes or issues, with full-row dashboard colors and filterable Category/Issue columns |
| Changes | Only changed fields, with the old baseline value and raw new-pull value side by side |
| Roster Draft | The complete proposed roster, including unchanged active employees, for reviewing/editing the next baseline |

Blue means information changed, red means missing from the new pull, green means new employee, orange/pink means a manager issue, and yellow means capability review. Review can contain several findings for one employee. Roster Draft's capability-cell colors describe value sources separately from the review-category legend.

For Excel edits, correct employee values on **Roster Draft**, save the complete workbook as a reviewed copy, and click **Import reviewed roster as baseline**. The app previews the employee/edit counts and confirmed username links before asking for approval. This imports the actual saved Excel edits, migrates confirmed correction keys, and reconciles capability corrections. **Initialize baseline** is for first-time setup. **Approve in-app changes** is the alternative for users who completed all edits in the app; it does not read Excel edits.

Summary counts, Review, Changes, and Capability Source are snapshots taken at export time. Editing the draft does not recalculate those values. Rerun comparison after loading the approved roster to generate a fresh report.

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

Report and baseline sheets use their structured Excel tables' filters; Summary uses a worksheet filter. After updating from an earlier version that produced Excel repair warnings, run the comparison and export a fresh report. Previously generated files are not changed by the update. Verify the new export opens normally in desktop Excel before sharing it.

1. Open the app.
2. Click **Choose file** and select Sven's new `.xlsx` or `.csv` file.
3. Click **Run update**.
4. Review the colored items. Select a row to edit a capability or apply the manager suggestion.
5. Resolve any verified username links in the app, then click **Export report**.
6. Review the workbook with Amanda. Make Excel corrections on **Roster Draft** and save the **complete workbook** as a clearly named reviewed copy, for example `November_Roster_Reviewed.xlsx`.
7. Close the workbook, click **Import reviewed roster as baseline**, and select that saved copy. Check the preview counts and approve only when the review is complete.
8. The app saves the new baseline, reconciles capability corrections, archives the exact reviewed workbook, and keeps backups/audit details. Next month's comparison starts from that baseline.

If all changes were made in the app and you did not edit Excel, **Approve in-app changes** remains available as an alternative. Do not use it to import workbook edits.

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

### Reviewing a changed Logon Name

The app checks otherwise unmatched employees for the same nonblank Display Name and at least two matching, nonblank fields among Title, Department, Manager, and Manager Email Address. This is evidence for review, not proof of identity. A surname/login/email change can therefore be found when the display name and supporting fields remain consistent; changes with insufficient evidence still appear as new/missing employees.

1. Select a purple **Possible Logon Name Change** finding. Current and Suggested show the old and new logon names; Notes contains the matching values.
2. Click **Confirm logon change**, verify the employee and evidence, and confirm. Ambiguous candidates involving multiple qualifying records cannot be linked; verify/correct the source files instead.
3. Review the recalculated proposal. The linked employee is no longer counted as new/departed. Changes includes the old/new Logon Name and other changed fields; Review records the confirmation.
4. Confirmed links let Roster Draft retain the prior ERP Category and use the existing capability precedence. Saved old-key corrections carry into the preview; an already saved new-key correction takes priority.
5. Use **Clear logon confirmations** to discard all preview links before approval. Links also reset when the input baseline/monthly bytes change or the app closes. Running the same comparison again retains links, including after capability edits.
6. **Approve in-app changes** persists the new-key roster and migrates old-key corrections, preserving their source/notes metadata. History includes a corrections backup and a confirmed-key mapping receipt. Confirmation alone does not write shared data. As with other approvals, this is a sequence of file writes rather than a multi-file transaction.

Possible matches remain counted as new/departed until confirmed and are included in unresolved-item warnings. Approval still allows unresolved items after a warning, so inspect these findings before approving. If the baseline, source pull, or saved corrections change after confirmation, rerun before approval.

If you edit Roster Draft in Excel, use **Import reviewed roster as baseline** to preserve exported confirmations and reconcile corrections. The workbook retains its comparison context across app restarts; no active comparison is required for import. Older reports without context, raw roster files, or reports exported from a different/older baseline must be replaced with a fresh app export.

### Reviewed workbook rules

- Keep all four visible tabs and save as `.xlsx` (or `.xlsm`). A very-hidden `_RosterContext` tab retains source/proposal snapshots and confirmed links. It is import context, not an editable approval form or a permanent identity rule.
- Edit employee values on Roster Draft, including ERP Category. Keep every employee, Logon Name, and standard roster column. If employees or usernames are wrong, correct the monthly source and run a new comparison instead of deleting rows or changing keys in the draft.
- A nonblank capability edited in Excel becomes an approved **Reviewed Excel** correction for later comparisons. Clearing capability removes its saved correction and leaves the imported baseline blank; future source pulls can provide a value under the usual precedence rules.
- Unchanged capability corrections retain their metadata. Confirmed old-key corrections move to the new key; an explicit Excel capability edit takes priority over the exported correction.
- Changed baseline/correction snapshots or a workbook modified after the preview cause import to stop. Export a fresh report when the baseline or corrections changed. Source-monthly files are not needed during reviewed import because the workbook retains its original source snapshot.
- Summary, Review, Changes, and Capability Source remain export-time snapshots. Import archives the exact edited workbook and records field edits in a history receipt; rerun comparison to generate current findings. Import is the explicit approval of all saved roster values, not an automatic acceptance of suggestions.
- The reviewed importer stages files, backs up baseline/corrections/audit data, and attempts rollback on write failures. A filesystem crash or unavailable storage may still require restoring history backups. No claim of a filesystem-wide transaction is made.

For a fictional rename demo, run `py -m scripts.make_sample_data` and use `sample_rename_baseline.xlsx` and `sample_rename_monthly_pull.xlsx` from `sample_data/` with a separate test shared folder.

- `Logon Name` is the unique employee key. Duplicate logon names stop the update.
- The app never automatically applies a manager capability suggestion.
- A manager is flagged as potentially inactive when the manager cannot be found in the new active-user pull.
- The application can still approve a month with unresolved items, but it warns the user first.
- Do not place the shared data folder inside the application folder. Keep it in a company-approved shared location.







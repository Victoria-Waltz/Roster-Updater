# Project structure

The code is grouped by responsibility. The monthly workflow and shared data file names are unchanged.

| Location | Responsibility |
| --- | --- |
| `roster_updater/constants.py` | Column names and accepted header aliases |
| `roster_updater/models.py` | Roster, comparison, change, and review data models |
| `roster_updater/utils.py` | Text normalization and file hashing |
| `roster_updater/loading.py` | CSV/Excel reading and input validation |
| `roster_updater/comparison.py` | Employee changes, capability precedence, and manager suggestions |
| `roster_updater/identity.py` | Conservative username-change candidates, confirmed-link validation, and correction-key mapping |
| `roster_updater/review_categories.py` | Shared review category labels and colors used by the desktop and export |
| `roster_updater/corrections.py` | Reading saved capability corrections |
| `roster_updater/excel.py` | Common workbook formatting and baseline output |
| `roster_updater/reports.py` | Monthly report sheets and file-to-report workflow |
| `roster_updater/review_context.py` | Workbook snapshot/context retained during Excel editing |
| `roster_updater/reviewed_import.py` | Reviewed import preview, correction reconciliation, backups, and promotion |
| `roster_updater/storage.py` | Shared files, locking, corrections, backups, and approvals |
| `roster_updater/ui/app.py` | Application startup and button actions |
| `roster_updater/ui/view.py` | Window layout, styles, and review list display |
| `roster_updater/ui/dialogs.py` | Capability correction dialog |
| `roster_updater/ui/config.py` | Per-user settings |
| `scripts/make_sample_data.py` | Synthetic sample workbook generator |
| `tests/test_workflow.py` | Monthly workflow regression tests |

## Entry points

Run these commands from the repository root after installing `requirements.txt`:

```sh
python -m roster_updater
python -m scripts.make_sample_data
python -m unittest discover -s tests -v
```

`roster_app.py` and `make_sample_data.py` remain small compatibility launchers. The existing `run_app.bat` and `build_windows_app.bat` still target `roster_app.py`, so Windows users can use the same commands. Keep the `roster_updater` package beside the launcher when running or building from source; a built executable contains the imported modules.

`roster_engine.py` re-exports the original public models, constants, and functions for existing scripts. New code should import directly from the module that owns the responsibility. Internal package modules do not import the compatibility file.

## Dependency direction

The UI calls storage and reporting. Storage calls loading, comparison, corrections, and Excel output. These modules share models, constants, and normalization utilities. Core processing modules do not import the UI or start a desktop window.

`RosterUpdaterApp` extends `RosterView`: the view contains presentation methods, while the application owns state and action handlers. The capability dialog and user preferences are independent modules.

## Data and testing

The generator writes synthetic workbooks to `sample_data/` at the repository root. Tests create and clean up their own temporary folders and do not touch the user's configured shared folder. Generated sample files are ignored by Git. Keep real employee files in the approved shared data location described in the README.

The automated tests cover comparison, capability corrections, CSV validation, report output, approval backups/audit history, stale approvals, locking, and launcher imports. Interactive Tkinter behavior and Windows executable packaging should also be checked on Windows before distribution.




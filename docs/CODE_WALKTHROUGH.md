# Reading and understanding the code

Start with this guide, then use [PROJECT_STRUCTURE.md](PROJECT_STRUCTURE.md) to find the file responsible for a particular feature. Functions and classes contain docstrings: the text inside triple quotes immediately below a definition explains what it does, its inputs/results, and relevant side effects. Inline `#` comments explain decisions within longer workflows.

## Terms used throughout the app

| Term | Meaning |
| --- | --- |
| Baseline | The employee roster most recently approved by a user |
| Monthly pull | The new active-user file being compared with that baseline |
| Canonical header | The standard spelling used internally, such as `Logon Name` |
| Normalized key | Trimmed, case-insensitive text used to match employees |
| Override / correction | A capability value explicitly saved by a user and reused in later comparisons |
| Review item | One finding for a reviewer; one person can have several findings |
| Field change | One difference between an old field and its raw incoming value |
| Proposal / result | The comparison held in memory before approval |
| Hash | A fingerprint of file bytes used to detect a changed baseline |

## Follow a monthly update

1. `roster_app.py` calls `ui/app.py:main()`. The app creates state variables, builds widgets through `RosterView`, and starts Tkinter's event loop, which waits for user interaction.
2. Choosing a folder creates a `SharedRosterStore`. Its constructor records paths; `prepare()` creates directories. `ConfigStore` remembers the folder in local preferences.
3. `run_update()` calls `store.compare()`, which loads the baseline, monthly file, and saved corrections. `load_roster()` normalizes headings and rows, validates required columns and employee keys, and records file metadata.
4. `compare_rosters()` builds fresh proposed rows. It matches people by normalized logon name, computes raw field changes, identifies missing/new people, and adds manager/capability findings.
5. `_display_result()` populates summary labels and the review table. A dictionary connects each table row ID to its full `ReviewItem` object.
6. Saving a capability correction writes the corrections JSON and runs comparison again. It does not approve the baseline. Recalculation also refreshes suggestions affected by a manager's corrected capability.
7. Export writes the current result to Excel. Approval additionally backs up data, replaces the baseline, and records an audit entry.

## Example: why an empty capability is preserved or suggested

Suppose Alex has `Enable` in the baseline but a blank capability in the new pull. The proposal keeps `Enable`. A new employee, Dana, has no baseline value and reports to Alex. Dana's proposed capability remains blank, with `Enable` offered as a review suggestion. Only an explicit correction applies it.

The priority order is: saved nonblank correction, nonblank monthly capability, nonblank baseline capability, then blank. A saved correction takes priority even if a later pull provides a different value; it must be changed or removed if no longer appropriate.

The Changes sheet compares baseline fields with the raw monthly file. The Roster Draft sheet shows the final selected capabilities after precedence rules. These sheets answer different questions, so their capability values can differ. Review contains findings only, with a Category column and row colors matching the desktop through `review_categories.py`. Unchanged employees remain available on Roster Draft, alongside the focused comparison tabs in the same workbook.

## How manager matching works

The comparison builds dictionaries for email and display name from the proposed roster. A manager's email is tried first; if not found, an unambiguous display name is tried. Duplicate display names are removed from the name lookup to avoid guessing. Email duplicates currently keep the last record. An unmatched manager is a possible inactive account, not a confirmed departure. Hierarchy lookup checks only the manager and the manager's manager.

## What writes to disk

| Action | Files affected |
| --- | --- |
| Choose shared folder | Local preferences and any missing shared directories |
| Initialize baseline | Baseline workbook; backup if replacing an existing baseline |
| Run comparison | None |
| Save/remove correction | Corrections JSON |
| Export | Report at the selected destination |
| Approve | Baseline backup, source-pull copy, report, new baseline, and audit CSV |

The shared lock uses exclusive file creation to serialize cooperating writers accessing the same filesystem. A stale comparison is rejected if the baseline's hash changed after loading. Workbook and JSON writes use temporary files before replacement, but the entire approval sequence is not one transaction. A failure midway can leave some steps completed. Independently synchronized folders do not make the lock a distributed coordination mechanism.

## How the desktop code fits together

`RosterUpdaterApp` inherits layout/display methods from `RosterView`. A callback is a function passed to a button or dialog to run later. For example, `CapabilityDialog` collects inputs and calls its `on_save` callback; the controller saves them through the store. `StringVar` values connect application state to labels and input fields. The app processes work on the GUI thread; its busy cursor does not create a background worker.

`RosterError` describes expected processing problems in user-facing language. UI actions catch errors and display them. Core processing remains usable without creating a desktop window, which lets tests exercise it in temporary folders.

## Where to make common changes

- Add a supported heading spelling: update `HEADER_ALIASES` in `constants.py`.
- Change employee comparison rules: use `comparison.py` and add a workflow regression case.
- Change report colors or formatting: use `excel.py` and `reports.py`.
- Change a button/layout: use `ui/view.py`; change what it does in `ui/app.py`.
- Change shared filenames or approval behavior: use `storage.py`.
- Change preferences: use `ui/config.py`.

When changing report columns, keep `ReviewItem.as_row()` and `FieldChange.as_row()` aligned with the matching report headers. The report appends a Category value to each model's exported row. When adding input columns, review loading, baseline output, and comparison rules together.

For manual Excel corrections, edit Roster Draft, save a reviewed copy, and use Initialize baseline once that roster is approved. The loader detects the full roster sheet among the other tabs. This backs up and replaces the baseline; it does not synchronize workbook edits into an existing in-memory comparison. Approve as new baseline still uses the app's current result. Saved capability corrections persist independently, so a correction may need to be updated or removed if a later Excel edit supersedes it. The Summary tab explains this distinction and that counts/findings do not recalculate inside Excel.

Run `python -m unittest discover -s tests -v` from the repository root. The tests use synthetic employees and temporary folders. They cover core workflows and imports; they do not replace an interactive Windows check or an executable build test.


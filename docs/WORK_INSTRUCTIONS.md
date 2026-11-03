# TAIT Roster Updater: Easy Instructions

Compare the approved roster with Sven's new monthly file, review changes, and save the corrected roster for next month.

## 1. First-time setup only

1. Open **TAIT Roster Updater**.
2. Click **Choose folder** and select the company-approved shared folder where you will keep the baseline and reports. Keep this separate from the app folder. For practice, use a separate test folder.
3. Click **Initialize baseline**.
4. Select your current approved roster, for example **October Roster.xlsx**.

When you see **Baseline ready**, setup is complete. Use this same shared folder each month. If a baseline already exists, skip this setup.

## 2. Compare the new monthly file

1. Under **Sven's new file**, click **Choose file**.
2. Select the new roster, for example **November Roster.xlsx**.
3. Click **Run update**.

The app lists employees with changes or issues. One employee can have several findings.

If it says **Comparison complete. 0 items are available for review**, the comparison finished and found nothing to flag.

## 3. Review the findings

| Color | Meaning | What to do |
| --- | --- | --- |
| Green | New employee | Check that they should be included. |
| Red | Missing from the new file | Check whether they left or their username changed. |
| Blue | Employee information changed | Check the old and new values. |
| Yellow | Capability suggestion or review | Check the capability before accepting a suggestion. |
| Orange/pink | Manager issue | Check the manager information. |
| Purple | Possible or confirmed username change | Verify whether the old and new usernames belong to the same person. |

For a verified username change, select the purple **Possible Logon Name Change** row and click **Confirm logon change** before exporting. Only confirm when you have verified it is the same person. A confirmation alone does not save a new baseline.

To accept a capability suggestion, select its row and click **Apply suggestion**. Use **Edit capability** to enter a different value. Suggestions are not applied automatically.

## 4. Export and make corrections in Excel

Click **Export report**, save the workbook, and open it in Excel.

| Tab | Use it for |
| --- | --- |
| **Summary** | Counts, color guide, and instructions |
| **Review** | Employees with changes or issues; filter Category or Issue to focus the list |
| **Changes** | Specific old baseline and new-file values |
| **Roster Draft** | The full proposed roster; make your corrections here |

**Edit the Roster Draft tab in this exported workbook.** You do not need to edit the original October baseline separately.

Keep the entire workbook, including all its tabs. Keep the column headings and employee Logon Names unchanged, and do not add or delete employee rows. If those need correcting, fix the source roster and run a new comparison.

Summary, Review, and Changes do not refresh when you edit Roster Draft. Roster Draft capability-cell colors show where values came from; use the Review tab's category colors to understand findings.

## 5. Save the reviewed roster as the new baseline

1. Finish reviewing the roster and making corrections on **Roster Draft**.
2. Save the **entire workbook** as an Excel **.xlsx** file, for example **November Roster Reviewed.xlsx**.
3. Close the workbook in Excel.
4. Return to the app and click **Import reviewed roster as baseline**.
5. Select the workbook you just saved.
6. Check the filename, employee count, edited-field count, and confirmed username links in the preview. Read any warning about unresolved findings.
7. Approve only when the review is complete.

When the app says the reviewed roster was imported, you are finished. The review list clears, and your corrected roster becomes the baseline for next month. The app also keeps backups and a copy of the reviewed workbook.

**For Excel edits, use Import reviewed roster as baseline.** The **Approve in-app changes** button saves only changes made inside the app; it does not read your Excel edits.

## Next month

Choose the same shared folder, select Sven's next monthly file, and click **Run update**. You do not need to initialize the baseline again.

## Quick help

- **Baseline already initialized:** Continue with the monthly update. To save a reviewed Excel workbook, use **Import reviewed roster as baseline**.
- **Zero review items:** The comparison ran and found nothing to flag.
- **Import asks for a fresh export:** Run a new comparison and export a new workbook. Older reports or reports based on a baseline/corrections that have since changed cannot be imported.
- **Need to find the saved baseline:** Click **Open shared folder**. The approved roster is **current_baseline.xlsx**.

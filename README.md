# Matching tool

This bound Google Apps Script adds **Home Depot → Apply approved candidates** to the reconciliation workbook. Set column A to Candidate 1, Candidate 2, or Candidate 3 on the rows you want to post, then run the menu command and confirm the preview. The command scans the whole audit, including filtered-out rows. It writes signed amounts to `JULY!H` and records them in the application log.

## Connect the existing Apps Script project

1. Copy the Script ID from **Extensions → Apps Script → Project Settings** in the target workbook.
2. In the repository root, create an ignored `.clasp.json` with `{"scriptId":"YOUR_SCRIPT_ID"}`.
3. Run `clasp login` if needed. Enable the [Apps Script API for your Google account](https://script.google.com/home/usersettings) if prompted.
4. Back up the current online project with `clasp clone-script YOUR_SCRIPT_ID` in a separate directory. A push replaces the Apps Script project's files. Preserve any other code and settings before syncing.
5. In the repository, check `clasp show-file-status` and run `clasp push`. Reload the spreadsheet to see the menu.

The manifest declares the advanced Sheets service. A project linked to a standard Google Cloud project also needs the Sheets API enabled in that Cloud project.

The script expects the July 2026 reconciliation audit layout. It skips conflicting identifiers, ambiguous jobs, and existing amounts it cannot reconcile to the application log. The command runs only when you choose it and confirm the preview.

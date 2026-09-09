# Unitflow

Static client unit tracker UI backed by the supplied Google Sheet through Google Apps Script.

## Run the interface

Open `index.html` in a browser. When `API_URL` is configured, the login page validates credentials against the `Accounts` sheet through Google Apps Script. If `API_URL` is blank, the UI runs in local demo mode.

Demo accounts all use password `demo123`:

- `technician`: full workspace access
- `office`: all branches and unit data
- `rosales.admin`: BNB Rosales branch only

When using the Apps Script backend before creating your first permanent account, the script automatically adds a temporary row to the `Accounts` sheet. Sign in with username `temp.technician` and password `UnitflowTemp2026!`. Use it to create a permanent technician account, then remove `TEMP_TECHNICIAN` and `ensureTemporaryAccount` from `gs/Code.gs` and redeploy the web app.

## Connect Google Sheets

1. In the supplied spreadsheet, create sheets named `Accounts`, `Units`, and `Branches`. The current Apps Script also creates these tabs automatically if they are missing.
2. Add these header rows:
   - `Accounts`: `name`, `username`, `password`, `role`, `branch`, `status`
   - `Units`: `Unit Code`, `Client Name`, `Model`, `Processor`, `RAM`, `Storage Size (HDD/SSD/SD)`, `Unit Price`, `Status`, `Current Location`, `Date Received`, `Released Date`
   - `Branches`: `name`, `created`
3. Open **Extensions > Apps Script**, paste in `gs/Code.gs`, and deploy it as a web app. Set access to the people who should use the tracker.
4. Put the deployed web app URL into `API_URL` at the top of `js/app.js`.
5. In Apps Script, select `initializeDefaultBranches` from the function menu and click **Run** once to create the six starting branch rows and tabs. Branch tabs are synchronized automatically when branches or units are created, edited, or deleted.
6. If old branch tabs remain after deleting their branch rows, select `cleanupOrphanBranchSheets` and click **Run** once. It deletes only tabs using the Unitflow unit-table headers that no longer have an active branch row.
7. To remove all six original branches and their tabs at once, select `deleteAllBranches` and click **Run**. This preserves `Accounts`, `Units`, and `Branches`.

For production, replace the plain-text password column with salted password hashes and move login to a proper identity provider. Google Apps Script is suitable for a small internal tool, but it should not be treated as a high-security identity system.

Unit statuses are restricted to: `For observation`, `Released`, `For release`, `To be transfered`, and `In warehouse`.

Unit prices are entered and displayed in Philippine Peso (`PHP`).

Technician accounts can view and manage the complete workspace, including accounts. Office accounts can view the complete workspace and accounts, but cannot edit accounts. Admin accounts see their permitted unit locations: their assigned branch, `BNB Rosales branch`, and `Warehouse`. Admins can edit units and change their location only among those three locations. Admins cannot access Accounts or Branches management.
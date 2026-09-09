# Unitflow

Static client unit tracker UI backed by the supplied Google Sheet through Google Apps Script.

## Run the interface

Open `index.html` in a browser. The UI starts in local demo mode so the workflow can be reviewed without credentials or an API URL.

Demo accounts all use password `demo123`:

- `technician`: full workspace access
- `office`: all branches and unit data
- `rosales.admin`: BNB Rosales branch only

When using the Apps Script backend before creating your first permanent account, temporarily sign in with username `temp.technician` and password `UnitflowTemp2026!`. Use it to create a permanent technician account, then remove `TEMP_TECHNICIAN` from `gs/Code.gs` and redeploy the web app.

## Connect Google Sheets

1. In the supplied spreadsheet, create sheets named `Accounts`, `Units`, and `Branches`.
2. Add these header rows:
   - `Accounts`: `name`, `username`, `password`, `role`, `branch`, `status`
   - `Units`: `Unit Code`, `Client Name`, `Model`, `Processor`, `RAM`, `Storage Size (HDD/SSD/SD)`, `Unit Price`, `Status`, `Current Location`, `Date Received`, `Released Date`
   - `Branches`: `name`, `created`
3. Open **Extensions > Apps Script**, paste in `gs/Code.gs`, and deploy it as a web app. Set access to the people who should use the tracker.
4. Put the deployed web app URL into `API_URL` at the top of `js/app.js`.

For production, replace the plain-text password column with salted password hashes and move login to a proper identity provider. Google Apps Script is suitable for a small internal tool, but it should not be treated as a high-security identity system.

Unit statuses are restricted to: `For observation`, `Released`, `For release`, `To be transfered`, and `In warehouse`.

Unit prices are entered and displayed in Philippine Peso (`PHP`).

Technician accounts can view and manage the complete workspace, including accounts. Office accounts can view the complete workspace and accounts, but cannot edit accounts. Admin accounts see their permitted unit locations: their assigned branch, `BNB Rosales branch`, and `Warehouse`. Admins can edit units and change their location only among those three locations. Admins cannot access Accounts or Branches management.
# Setup Guide

## 1. Create The Spreadsheet Tabs

Create these Google Sheets tabs:

- `Source_Inventory`
- `Inventory_Master`
- `Students`
- `Transaction_Log`
- `QR_Labels`
- `QR_Print`
- `Inventory_Audit_Log`

See `docs/sheet-schema.md` for required columns.

## 2. Add Apps Script Files

In Apps Script, create or replace:

- `Code.gs`
- `instrument.gs`
- `Index.html`

Copy the matching files from `src/`.

`VisualSample.html` is optional and exists only as a design prototype.

## 3. Configure Script Properties

Open Apps Script project settings and add:

```text
ADMIN_PASSWORD=replace-with-backup-password
ADMIN_EMAILS=admin@example.edu,assistant@example.edu
WEB_APP_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
```

Use real values only in the Apps Script editor, not in GitHub.

## 4. Deploy The Web App

Recommended:

```text
Execute as: Me
Who has access: Anyone with a Google account
```

If the deployment must be limited to staff and students in one organization, use your Workspace domain option.

The public web app does not need direct spreadsheet access for students. It should run through Apps Script server functions, which perform validation and logging.

## Public Data Boundary

The student-facing lookup returns redacted instrument data only. It does not return the full student roster, current borrower identity, checkout dates, operational notes, transaction logs, or audit logs.

Admin tools unlock through an admin session token created by approved Google account access or the backup password.

## 5. Install Optional Triggers

From the spreadsheet menu:

- `Inventory Tools > Install QR Automation Triggers`
- `Inventory Tools > Install Parent Consent Trigger`

## 6. Generate IDs And QR Labels

Use:

- `Inventory Tools > Generate Missing IDs`
- `Inventory Tools > Configure QR Web App URL`
- `Inventory Tools > Sync Valid QR Labels`

Do not regenerate IDs after QR labels are physically attached unless you intentionally plan a full relabeling.

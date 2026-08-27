# QR Instrument Checkout Web App

A Google Apps Script and Google Sheets inventory web app for managing school-owned instrument checkout, return, QR lookup, admin inventory tools, and end-of-year reconciliation.

This repository is a sanitized portfolio/demo version. It contains no production student data, parent data, real deployment URLs, real admin emails, passwords, school logos, or private spreadsheet IDs.

## What It Does

- Loads instruments from QR URLs using `?inst=InstrumentID`
- Validates student checkout eligibility server-side without returning the full student roster to the browser
- Requires typed signatures in `Last, First` format
- Logs every checkout and check-in transaction
- Blocks normal checkout when an instrument is already out
- Supports password or Google-account allowlisted admin access
- Allows admins to add instruments and students
- Warns admins about duplicate serial numbers
- Supports inventory audit/reconciliation logging
- Generates QR label data only from valid source inventory IDs
- Uses `LockService` to reduce race-condition risk during transactions

## Tech Stack

- Google Apps Script
- JavaScript
- HTML
- CSS
- Google Sheets
- Google Forms integration pattern

## Portfolio Security Highlights

- Server-side authorization checks for admin workflows
- Server-side validation for student ID, signature format, student lookup, repair holds, and override checkout
- Redacted public instrument lookup that does not expose current student assignment, checkout dates, notes, or logs
- Admin-only access to richer checkout/audit details through short-lived admin sessions
- Transaction logging with timestamps and active user email when available
- Separation between student-facing and admin-only workflows
- Script Properties used for configurable secrets and deployment-specific values
- Sheet-protection guidance for permanent IDs, generated QR labels, and logs
- Race-condition mitigation with `LockService`

## Repository Layout

```text
public-demo/
  README.md
  .gitignore
  appsscript.json.example
  src/
    Code.gs
    instrument.gs
    Index.html
    VisualSample.html
  docs/
    setup.md
    sheet-schema.md
    security-notes.md
  examples/
    sample-source-inventory.csv
    sample-inventory-master.csv
    sample-students.csv
```

## Required Script Properties

Configure these in Apps Script project settings:

```text
ADMIN_PASSWORD=replace-with-backup-password
ADMIN_EMAILS=admin@example.edu,assistant@example.edu
WEB_APP_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
```

Never commit real Script Property values.

## Deployment

Recommended web app deployment:

```text
Execute as: Me
Who has access: Anyone with a Google account
```

For stricter environments, restrict access to your Workspace domain.

See [docs/setup.md](docs/setup.md) for setup details.

## Public vs Admin Data Boundary

The public checkout page intentionally receives only the minimum data needed to complete a checkout workflow:

- instrument ID
- instrument name/type
- make/model
- condition
- availability/status
- repair-hold status

It does not receive:

- the student roster
- current borrower identity
- checkout dates
- operational notes
- transaction log rows
- audit log rows

Admin-only functions require an admin session token created by approved Google account access or the backup password.

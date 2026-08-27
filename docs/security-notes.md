# Security Notes

This app is designed for a lightweight school inventory workflow, not as a replacement for enterprise asset-management software.

## Do Not Publish Production Values

Never commit:

- real student or parent data
- real transaction logs
- real audit logs
- real school spreadsheet exports
- real admin emails
- real admin passwords
- real deployed Apps Script URLs
- real Google Form URLs
- school or district logos unless you have permission

## Admin Access

Admin access uses two mechanisms:

- Google account allowlist through `ADMIN_EMAILS`
- backup password through `ADMIN_PASSWORD`

Both are stored in Apps Script Script Properties.

Anyone who can edit the Apps Script project should be treated as a trusted admin, because script editors can alter code and potentially read Script Properties.

## Server-Side Validation

Do not rely on client-side HTML or JavaScript as a security boundary. The server-side Apps Script functions validate:

- admin authorization
- student ID format
- student existence
- signature/name match
- instrument checkout state
- override password
- repair hold status

## Least-Privilege Browser Data

The student-facing browser is intentionally not given the complete `Students` roster. Students type their ID, and the server validates the ID/signature pair internally during checkout or check-in.

Public instrument lookup returns only redacted instrument data:

- instrument ID
- instrument name/type
- make/model
- condition
- availability/status
- repair-hold status

Public lookup does not return current borrower identity, current borrower student ID, checkout dates, operational notes, transaction logs, audit logs, or staff emails.

Richer operational details are available only through admin-authorized functions using the admin session token. Examples include audit expected-student fields and admin dashboard/audit workflows.

## Sheet Access

Recommended:

- Do not give students direct spreadsheet access.
- Give sheet/script edit access only to trusted staff.
- Protect permanent ID columns and generated/formula areas.
- Protect logs or at least log headers.

## Race Conditions

Checkout and check-in operations use `LockService` so two simultaneous scans are less likely to create conflicting records.

## Privacy

Student names and IDs are sensitive operational data. Avoid publishing screenshots, CSV exports, or demo videos that show real records.

## Remaining Limitations

- Apps Script projects are not protected from trusted script editors. Anyone with script edit access can alter code and may be able to expose Script Properties or sheet data.
- The public checkout/check-in functions must still accept a student ID and signature to validate transactions. The demo uses generic validation errors to reduce roster enumeration, but any live deployment should still monitor for abuse.
- Instrument availability is public to anyone who can open a valid QR URL. The current borrower identity is not public.
- `Session.getActiveUser().getEmail()` behavior depends on Apps Script deployment settings and Google account context; it may be blank in some deployments.

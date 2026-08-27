# Sheet Schema

## Source_Inventory

Headers should be on row 2.

Required columns:

```text
Instrument_UID
Instrument_ID
Instrument Name
Instrument Make
Instrument Model
Instrument Serial Number
Condition
```

`Instrument_UID` and `Instrument_ID` are permanent identifiers. Do not edit them manually after labels are in use.

## Inventory_Master

Headers should be on row 1.

Columns A-G mirror source inventory data:

```text
Instrument_UID
InstrumentID
Instrument
Make
Model
Serial
Condition
```

Columns H-N are app-controlled operational fields:

```text
Status
CurrentStudentID
CurrentStudentName
DateOut
DateIn
RepairStatus
Notes
```

Do not overwrite H-N from sync formulas or source imports.

## Students

Student records are authoritative for checkout validation.

```text
Column B = Student ID
Column C = Student Name
```

Rules:

- Student ID must be numeric only.
- Student Name must use `Last, First` format.
- Students not found in this sheet cannot check out instruments.

## Transaction_Log

Expected columns:

```text
Timestamp
Action
InstrumentID
StudentID
StudentName
UserEmail
Note
```

## QR_Labels

Expected columns:

```text
InstrumentID
QR
CodeText
```

`CodeText` should contain the deployed web app URL with `?inst=InstrumentID`.

## Inventory_Audit_Log

Expected columns:

```text
Timestamp
AuditSession
InstrumentID
InstrumentUID
InstrumentName
Make
Model
Serial
ExpectedStatus
ExpectedStudentID
ExpectedStudentName
ExpectedCondition
Found
SerialMatches
CasePresent
AccessoriesComplete
Condition
RepairNeeded
Notes
StaffEmail
```

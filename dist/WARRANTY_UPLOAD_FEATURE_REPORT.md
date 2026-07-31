# Warranty Upload Feature Report

## Scope

Added a MIS Master option for periodically replacing `tbl_warranty` from an uploaded Excel workbook.

## Input contract

- Accepted files: `.csv`, `.xlsx` and `.xls`
- Maximum file size: 5 MB
- Required headers within the first 20 worksheet rows:
  - `Item Code`
  - `Description`
  - `Month`
- Blank rows are ignored.
- Leading/trailing spaces, tabs, and line breaks are trimmed from imported text.
- `Item Code` is required and limited to 50 characters.
- `Description` is optional and limited to 300 characters.
- `Month` is optional; when provided it must be a whole number from 0 to 2147483647.
- Duplicate item codes are rejected case-insensitively.
- CSV files are read directly with standard ADO streams and do not require Microsoft Access Database Engine. Excel-style quoted fields, embedded commas, escaped quotes, and embedded line breaks are supported. Save CSV input as UTF-8.

## Replacement behavior

The complete workbook is validated before the database is changed. The delete and all inserts run in one database transaction. Any insert or commit failure rolls the transaction back so the prior warranty records remain available.

## Job submission warranty calculation

Within `Case "submitJob"` in `core/action.asp`, the submitted `job_Model` is validated as non-empty before querying `tbl_warranty`. When a matching numeric warranty month and valid purchase date are available, the warranty expiry date is calculated with `DateAdd("m", month, purchase_date)`. The job is marked `Under` through the expiry date and `Over` afterward. A missing model, purchase date, matching warranty row, or valid warranty month defaults to `Over`.

## Security and file handling

- The page is restricted to users whose existing `GAPS/slevel` cookie is `mis`.
- Multipart upload parsing uses `Request.BinaryRead` and binary `ADODB.Stream` objects; it does not require the optional `aspSmartUpload` COM component.
- Uploaded files receive a server-generated name in the operating system's temporary folder, outside the application URL space.
- Temporary workbooks are deleted after processing.
- SQL inserts use typed ADO parameters.

## Files

- Modified `core/mis_master.asp`
- Added `core/mis_master_warranty_upload.asp`
- Added this report at `dist/WARRANTY_UPLOAD_FEATURE_REPORT.md`

No files under `src/` were modified.

# CRM dashboard transformation report

Date: 2026-07-15

## Scope

- Source landing page identified as `core/rm_home.asp`. There is no `cs_home.asp` in the current checkout.
- Added transformed output `dist/core/rm_home.asp`.
- Added transformed login-routing output `dist/validateuser.asp`.
- No existing application file was modified.

## Dashboard behavior

- Displays current-day totals for jobsheets, non-cancelled invoices, non-cancelled receipts, and newly created customer records.
- Displays all-record jobsheet counts for the valid workflow statuses Open, Submitted, Accepted, Done, Posted, Cancelled, and Total so unresolved items remain visible after their creation date. The separate `Jobsheets Today` KPI remains current-date based.
- Displays invoice value, outstanding value on today's invoices, and receipts collected today.
- Displays current pending-approval counts for Stock-In, Stock-Out, Stock-Transfer, and Stock-Adjustment records at the `Submitted` stage. These are live outstanding totals and are not limited to the current date. Each count links to the corresponding Submitted list with an extended date range so the list reconciles with the dashboard backlog.
- The four top daily summary cards are clickable. Jobsheets, invoices, and receipts open their corresponding lists filtered to today; New Customers opens the customer list because that page does not support a created-date filter. The all-record jobsheet panel remains available separately below.
- Formats operational counts without thousands separators; monetary amounts retain normal currency grouping and two decimal places.
- Excludes `technician` and `technician2` access levels from the CRM dashboard. Both technician levels are routed to `rmtech_jobsheet_view.asp?job_status=Accepted`.
- Routes every other Riegen access level to `core/rm_home.asp`, including MIS, Finance, CS, SC, and Sales.
- Treats username `admin` as an explicit exception: login routes it to `core/mis_home.asp`, and direct access to the CRM dashboard redirects it to `mis_home.asp` regardless of access level.

## Data mapping

| Figure | Table | Current-date column | Other rule |
|---|---|---|---|
| Jobsheets | `tbljob` | `job_date` | Status breakdown uses `job_status` |
| Invoices | `tblinvoice` | `inv_date` | Excludes `inv_status = 'Cancel'` |
| Receipts | `tblreceipt` | `receipt_date` | Excludes `receipt_status = 'Cancel'` |
| New customers | `tblcustomer` | `cust_createddate` | Counts records created today |
| Stock-In pending approval | `tblstockin` | Not date-limited | `st_status = 'Submitted'` |
| Stock-Out pending approval | `tblstockOut` | Not date-limited | `so_status = 'Submitted'` |
| Stock-Transfer pending approval | `tblstocktransfer` | Not date-limited | `sf_status = 'Submitted'` |
| Stock-Adjustment pending approval | `tblstockadj` | Not date-limited | `sj_status = 'Submitted'` |

Date filtering uses an inclusive start (`YYYY-MM-DD 00:00`) and exclusive next-day boundary. This safely includes datetime values throughout the current day.

## Protection and compatibility checks

- Classic ASP include directives remain present in both outputs.
- Existing table names, column names, form keys, query-string keys, cookie keys, Session keys, and COM ProgIDs were not renamed.
- Existing source files remain unchanged; all new output is under `dist/`.
- The output is intended to overlay the matching application paths during the protected deployment/build process so its existing include targets resolve normally.

## Verification

- ASP delimiter counts are balanced in both generated ASP files.
- Required overlay include targets exist: `core/header.asp`, `core/footer.asp`, and `core/database/dbconnect.asp`.
- The technician landing target `core/rmtech_jobsheet_view.asp` exists.
- `git diff` confirms that `core/rm_home.asp` and `validateuser.asp` remain unchanged.
- A read-only execution check of the four dashboard queries was attempted against the configured database on 2026-07-15. The connection was blocked before query execution by the legacy provider error `[DBNETLIB][ConnectionOpen (SECCreateCredentials()).] SSL Security error.` Runtime values therefore still require verification in the configured IIS/database environment.

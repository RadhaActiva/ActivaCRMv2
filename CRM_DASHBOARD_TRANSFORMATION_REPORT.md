# CRM dashboard transformation report

Date: 2026-07-15

## Scope

- Source landing page identified as `core/rm_home.asp`. There is no `cs_home.asp` in the current checkout.
- Added transformed output `dist/core/rm_home.asp`.
- Added transformed login-routing output `dist/validateuser.asp`.
- Added database deployment output `dist/database/deploy_crm_dashboard.sql`.
- No existing application file was modified.

## Dashboard behavior

- Displays current-day totals for jobsheets, non-cancelled invoices, non-cancelled receipts, and newly created customer records.
- Displays all-record jobsheet counts for the valid workflow statuses Open, Submitted, Accepted, Done, Posted, Cancelled, and Total so unresolved items remain visible after their creation date. The separate `Jobsheets Today` KPI remains current-date based.
- Displays invoice value, outstanding value on today's invoices, and receipts collected today.
- Displays current pending-approval counts for Stock-In, Stock-Out, Stock-Transfer, and Stock-Adjustment records at the `Submitted` stage. These are live outstanding totals and are not limited to the current date. Each count links to the corresponding Submitted list with an extended date range so the list reconciles with the dashboard backlog.
- The four top daily summary cards are clickable. Jobsheets, invoices, and receipts open their corresponding lists filtered to today; New Customers opens the customer list because that page does not support a created-date filter. The all-record jobsheet panel remains available separately below.
- Formats operational counts without thousands separators; monetary amounts retain normal currency grouping and two decimal places.
- Displays a rolling three-month Top 10 Cities ranking using non-cancelled `tbljob` records grouped by `job_cust_city`.
- Displays a rolling three-month Top 10 Parts Replaced ranking using summed `tbljob_parts.jobp_qty`, joined to `tbljob` by `job_code`, and limited to Posted jobs to match the existing parts-analysis definition of completed replacement activity.
- Displays a rolling three-month Top 10 Technicians ranking by the number of Posted jobs, showing technician code and name. Cities, Parts Replaced, and Technicians use three equal-width cards on the same desktop row.
- Uses compact, standardized typography in the upper dashboard: 11px labels and notes, 18px KPI/status figures, reduced card heights, and tighter spacing based on the smaller ranking-card visual scale.
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
| Top 10 Cities | `tbljob` | Rolling three months by `job_date` | Groups non-cancelled jobs by trimmed `job_cust_city` |
| Top 10 Parts Replaced | `tbljob_parts` joined to `tbljob` | Rolling three months by `tbljob.job_date` | Posted jobs, grouped by `jobp_partcode`, ranked by summed `jobp_qty` |
| Top 10 Technicians | `tbljob` joined to `tbltechnician` | Rolling three months by `tbljob.job_date` | Posted jobs, grouped by `job_tech_code`, ranked by job count |

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

## Performance deployment

- `dbo.usp_CrmDashboard` consolidates the dashboard into one ADO command and returns six ordered result sets: daily jobsheets, all-record jobsheet statuses, combined daily/approval KPIs, Top 10 Cities, Top 10 Parts Replaced, and Top 10 Technicians.
- The ASP page materializes the two ranking result sets and closes the database connection before rendering the dashboard.
- The SQL deployment adds 11 guarded nonclustered indexes for the dashboard's date, status, technician, join, and aggregation paths.
- Run `dist/database/deploy_crm_dashboard.sql` against the ActivaCRM database before deploying `dist/core/rm_home.asp`. The page depends on `dbo.usp_CrmDashboard` and will not load until the procedure exists.
- Index creation can briefly consume database resources. Apply the SQL script during a low-traffic maintenance window and review for equivalent pre-existing indexes in production.
- Deployment status on 2026-07-16: not applied. The configured connection returned an SSL/encryption error, and elevated execution was not authorized because the target may be shared or production. No stored procedure or index changes were made to the database.

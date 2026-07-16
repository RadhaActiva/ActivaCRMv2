/*
  CRM dashboard database deployment
  1. Creates supporting indexes when their dashboard-specific names do not exist.
  2. Creates or updates dbo.usp_CrmDashboard.

  Run this script against the ActivaCRM database before deploying dist/core/rm_home.asp.
*/

SET NOCOUNT ON;

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tbljob') AND name = N'IX_tbljob_Dashboard_DateStatus')
    CREATE NONCLUSTERED INDEX IX_tbljob_Dashboard_DateStatus
        ON dbo.tbljob (job_date, job_status)
        INCLUDE (job_id, job_cust_city, job_code);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tbljob') AND name = N'IX_tbljob_Dashboard_Status')
    CREATE NONCLUSTERED INDEX IX_tbljob_Dashboard_Status
        ON dbo.tbljob (job_status)
        INCLUDE (job_id);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tbljob') AND name = N'IX_tbljob_Dashboard_TechnicianDate')
    CREATE NONCLUSTERED INDEX IX_tbljob_Dashboard_TechnicianDate
        ON dbo.tbljob (job_status, job_date, job_tech_code)
        INCLUDE (job_id);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tbljob_parts') AND name = N'IX_tbljob_parts_Dashboard_JobPart')
    CREATE NONCLUSTERED INDEX IX_tbljob_parts_Dashboard_JobPart
        ON dbo.tbljob_parts (job_code, jobp_partcode)
        INCLUDE (jobp_qty);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tblinvoice') AND name = N'IX_tblinvoice_Dashboard_DateStatus')
    CREATE NONCLUSTERED INDEX IX_tblinvoice_Dashboard_DateStatus
        ON dbo.tblinvoice (inv_date, inv_status)
        INCLUDE (inv_id, inv_totalAmt, inv_balance);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tblreceipt') AND name = N'IX_tblreceipt_Dashboard_DateStatus')
    CREATE NONCLUSTERED INDEX IX_tblreceipt_Dashboard_DateStatus
        ON dbo.tblreceipt (receipt_date, receipt_status)
        INCLUDE (receipt_id, receipt_totalpayment);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tblcustomer') AND name = N'IX_tblcustomer_Dashboard_CreatedDate')
    CREATE NONCLUSTERED INDEX IX_tblcustomer_Dashboard_CreatedDate
        ON dbo.tblcustomer (cust_createddate)
        INCLUDE (cust_id);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tblstockin') AND name = N'IX_tblstockin_Dashboard_Status')
    CREATE NONCLUSTERED INDEX IX_tblstockin_Dashboard_Status ON dbo.tblstockin (st_status) INCLUDE (st_id);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tblstockOut') AND name = N'IX_tblstockOut_Dashboard_Status')
    CREATE NONCLUSTERED INDEX IX_tblstockOut_Dashboard_Status ON dbo.tblstockOut (so_status) INCLUDE (so_id);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tblstocktransfer') AND name = N'IX_tblstocktransfer_Dashboard_Status')
    CREATE NONCLUSTERED INDEX IX_tblstocktransfer_Dashboard_Status ON dbo.tblstocktransfer (sf_status) INCLUDE (sf_id);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.tblstockadj') AND name = N'IX_tblstockadj_Dashboard_Status')
    CREATE NONCLUSTERED INDEX IX_tblstockadj_Dashboard_Status ON dbo.tblstockadj (sj_status) INCLUDE (sj_id);
GO

IF OBJECT_ID(N'dbo.usp_CrmDashboard', N'P') IS NULL
    EXEC(N'CREATE PROCEDURE dbo.usp_CrmDashboard AS SET NOCOUNT ON;');
GO

ALTER PROCEDURE dbo.usp_CrmDashboard
    @AsOfDate datetime = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @DateFrom datetime;
    DECLARE @DateTo datetime;
    DECLARE @PeriodFrom datetime;

    SET @DateFrom = DATEADD(day, DATEDIFF(day, 0, ISNULL(@AsOfDate, GETDATE())), 0);
    SET @DateTo = DATEADD(day, 1, @DateFrom);
    SET @PeriodFrom = DATEADD(month, -3, @DateFrom);

    /* Result 1: today's jobsheets. */
    SELECT COUNT(job_id) AS jobs_total
    FROM dbo.tbljob
    WHERE job_date >= @DateFrom AND job_date < @DateTo;

    /* Result 2: all-record jobsheet workflow. */
    SELECT
        COUNT(job_id) AS jobs_total,
        SUM(CASE WHEN job_status = 'Open' THEN 1 ELSE 0 END) AS jobs_open,
        SUM(CASE WHEN job_status = 'Submitted' THEN 1 ELSE 0 END) AS jobs_submitted,
        SUM(CASE WHEN job_status = 'Accepted' THEN 1 ELSE 0 END) AS jobs_accepted,
        SUM(CASE WHEN job_status = 'Done' THEN 1 ELSE 0 END) AS jobs_done,
        SUM(CASE WHEN job_status = 'Posted' THEN 1 ELSE 0 END) AS jobs_posted,
        SUM(CASE WHEN job_status = 'Cancel' THEN 1 ELSE 0 END) AS jobs_cancelled
    FROM dbo.tbljob;

    /* Result 3: today's financial/customer KPIs and the current stock approval backlog. */
    ;WITH InvoiceKpi AS
    (
        SELECT
            COUNT(inv_id) AS invoices_total,
            ISNULL(SUM(inv_totalAmt), 0) AS invoices_value,
            ISNULL(SUM(inv_balance), 0) AS invoices_outstanding
        FROM dbo.tblinvoice
        WHERE inv_date >= @DateFrom AND inv_date < @DateTo
          AND (inv_status IS NULL OR inv_status <> 'Cancel')
    ),
    ReceiptKpi AS
    (
        SELECT
            COUNT(receipt_id) AS receipts_total,
            ISNULL(SUM(receipt_totalpayment), 0) AS receipts_value
        FROM dbo.tblreceipt
        WHERE receipt_date >= @DateFrom AND receipt_date < @DateTo
          AND (receipt_status IS NULL OR receipt_status <> 'Cancel')
    ),
    CustomerKpi AS
    (
        SELECT COUNT(cust_id) AS customers_total
        FROM dbo.tblcustomer
        WHERE cust_createddate >= @DateFrom AND cust_createddate < @DateTo
    )
    SELECT
        i.invoices_total,
        i.invoices_value,
        i.invoices_outstanding,
        r.receipts_total,
        r.receipts_value,
        c.customers_total,
        (SELECT COUNT(st_id) FROM dbo.tblstockin WHERE st_status = 'Submitted') AS stockin_pending,
        (SELECT COUNT(so_id) FROM dbo.tblstockOut WHERE so_status = 'Submitted') AS stockout_pending,
        (SELECT COUNT(sf_id) FROM dbo.tblstocktransfer WHERE sf_status = 'Submitted') AS stocktransfer_pending,
        (SELECT COUNT(sj_id) FROM dbo.tblstockadj WHERE sj_status = 'Submitted') AS stockadjustment_pending
    FROM InvoiceKpi i
    CROSS JOIN ReceiptKpi r
    CROSS JOIN CustomerKpi c;

    /* Result 4: rolling three-month Top 10 Cities. */
    SELECT TOP 10
        LTRIM(RTRIM(job_cust_city)) AS city_name,
        COUNT(job_id) AS job_count
    FROM dbo.tbljob
    WHERE job_date >= @PeriodFrom AND job_date < @DateTo
      AND job_cust_city IS NOT NULL
      AND LTRIM(RTRIM(job_cust_city)) <> ''
      AND (job_status IS NULL OR job_status <> 'Cancel')
    GROUP BY LTRIM(RTRIM(job_cust_city))
    ORDER BY job_count DESC, city_name ASC;

    /* Result 5: rolling three-month Top 10 Parts Replaced from Posted jobs. */
    SELECT TOP 10
        jp.jobp_partcode,
        MAX(CONVERT(varchar(200), jp.jobp_desc)) AS part_description,
        SUM(ISNULL(jp.jobp_qty, 0)) AS total_qty
    FROM dbo.tbljob_parts jp
    INNER JOIN dbo.tbljob j ON jp.job_code = j.job_code
    WHERE j.job_date >= @PeriodFrom AND j.job_date < @DateTo
      AND j.job_status = 'Posted'
      AND jp.jobp_partcode IS NOT NULL
      AND LTRIM(RTRIM(jp.jobp_partcode)) <> ''
    GROUP BY jp.jobp_partcode
    ORDER BY total_qty DESC, jp.jobp_partcode ASC;

    /* Result 6: rolling three-month Top 10 Technicians by Posted jobs. */
    SELECT TOP 10
        j.job_tech_code,
        MAX(CONVERT(varchar(200), t.tech_name)) AS technician_name,
        COUNT(j.job_id) AS posted_jobs
    FROM dbo.tbljob j
    LEFT JOIN dbo.tbltechnician t ON j.job_tech_code = t.tech_code
    WHERE j.job_date >= @PeriodFrom AND j.job_date < @DateTo
      AND j.job_status = 'Posted'
      AND j.job_tech_code IS NOT NULL
      AND LTRIM(RTRIM(j.job_tech_code)) <> ''
    GROUP BY j.job_tech_code
    ORDER BY posted_jobs DESC, j.job_tech_code ASC;
END;
GO

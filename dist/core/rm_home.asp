<!-- #include file="header.asp" -->
<%
Dim dashboardAccessLevel
dashboardAccessLevel = LCase(Trim(Request.Cookies("GAPS")("slevel") & ""))

' The admin account always uses the MIS administration home page.
If LCase(Trim(Request.Cookies("GAPS")("sloginid") & "")) = "admin" Then
    Response.Redirect "mis_home.asp"
End If

' Technicians keep their operational landing pages; this dashboard is for CRM users.
If dashboardAccessLevel = "technician" Then
    Response.Redirect "rmtech_jobsheet_view.asp?job_status=Accepted"
ElseIf dashboardAccessLevel = "technician2" Then
    Response.Redirect "rmtech_jobsheet_view.asp?job_status=Accepted"
End If

Function DashboardNumber(value)
    If IsNull(value) Then
        DashboardNumber = 0
    ElseIf Trim(CStr(value)) = "" Then
        DashboardNumber = 0
    ElseIf IsNumeric(value) Then
        DashboardNumber = CDbl(value)
    Else
        DashboardNumber = 0
    End If
End Function

Function DashboardCount(value)
    DashboardCount = FormatNumber(DashboardNumber(value), 0, -1, 0, 0)
End Function

Dim dashboardDateFrom, dashboardDateTo, dashboardDateLabel, dashboardDateLink
dashboardDateFrom = ChkDateYYYYMMDD(Date())
dashboardDateTo = ChkDateYYYYMMDD(DateAdd("d", 1, Date()))
dashboardDateLabel = ChkDate(Date())
dashboardDateLink = Server.URLEncode(dashboardDateLabel)

Dim jobsTotal, jobsStatusTotal, jobsOpen, jobsSubmitted, jobsAccepted, jobsDone, jobsPosted, jobsCancelled
Dim invoicesTotal, invoicesValue, invoicesOutstanding
Dim receiptsTotal, receiptsValue, customersTotal
Dim stockInPending, stockOutPending, stockTransferPending, stockAdjustmentPending, stockApprovalTotal
Dim rsDashboard, sqlDashboard

jobsTotal = 0
jobsStatusTotal = 0
jobsOpen = 0
jobsSubmitted = 0
jobsAccepted = 0
jobsDone = 0
jobsPosted = 0
jobsCancelled = 0

sqlDashboard = "SELECT COUNT(job_id) AS jobs_total, " & _
    "SUM(CASE WHEN job_status = 'Open' THEN 1 ELSE 0 END) AS jobs_open, " & _
    "SUM(CASE WHEN job_status = 'Submitted' THEN 1 ELSE 0 END) AS jobs_submitted, " & _
    "SUM(CASE WHEN job_status = 'Accepted' THEN 1 ELSE 0 END) AS jobs_accepted, " & _
    "SUM(CASE WHEN job_status = 'Done' THEN 1 ELSE 0 END) AS jobs_done, " & _
    "SUM(CASE WHEN job_status = 'Posted' THEN 1 ELSE 0 END) AS jobs_posted, " & _
    "SUM(CASE WHEN job_status = 'Cancel' THEN 1 ELSE 0 END) AS jobs_cancelled " & _
    "FROM tbljob WHERE job_date >= '" & dashboardDateFrom & "' AND job_date < '" & dashboardDateTo & "'"

Set rsDashboard = Server.CreateObject("ADODB.Recordset")
rsDashboard.Open sqlDashboard, strconnect, 0, 1
If Not rsDashboard.EOF Then
    jobsTotal = DashboardNumber(rsDashboard("jobs_total"))
    jobsOpen = DashboardNumber(rsDashboard("jobs_open"))
    jobsSubmitted = DashboardNumber(rsDashboard("jobs_submitted"))
    jobsAccepted = DashboardNumber(rsDashboard("jobs_accepted"))
    jobsDone = DashboardNumber(rsDashboard("jobs_done"))
    jobsPosted = DashboardNumber(rsDashboard("jobs_posted"))
    jobsCancelled = DashboardNumber(rsDashboard("jobs_cancelled"))
End If
rsDashboard.Close

' The status panel is an all-record workflow snapshot, so unresolved actions remain visible on future days.
sqlDashboard = "SELECT COUNT(job_id) AS jobs_total, " & _
    "SUM(CASE WHEN job_status = 'Open' THEN 1 ELSE 0 END) AS jobs_open, " & _
    "SUM(CASE WHEN job_status = 'Submitted' THEN 1 ELSE 0 END) AS jobs_submitted, " & _
    "SUM(CASE WHEN job_status = 'Accepted' THEN 1 ELSE 0 END) AS jobs_accepted, " & _
    "SUM(CASE WHEN job_status = 'Done' THEN 1 ELSE 0 END) AS jobs_done, " & _
    "SUM(CASE WHEN job_status = 'Posted' THEN 1 ELSE 0 END) AS jobs_posted, " & _
    "SUM(CASE WHEN job_status = 'Cancel' THEN 1 ELSE 0 END) AS jobs_cancelled " & _
    "FROM tbljob"
rsDashboard.Open sqlDashboard, strconnect, 0, 1
If Not rsDashboard.EOF Then
    jobsStatusTotal = DashboardNumber(rsDashboard("jobs_total"))
    jobsOpen = DashboardNumber(rsDashboard("jobs_open"))
    jobsSubmitted = DashboardNumber(rsDashboard("jobs_submitted"))
    jobsAccepted = DashboardNumber(rsDashboard("jobs_accepted"))
    jobsDone = DashboardNumber(rsDashboard("jobs_done"))
    jobsPosted = DashboardNumber(rsDashboard("jobs_posted"))
    jobsCancelled = DashboardNumber(rsDashboard("jobs_cancelled"))
End If
rsDashboard.Close

invoicesTotal = 0
invoicesValue = 0
invoicesOutstanding = 0
sqlDashboard = "SELECT COUNT(inv_id) AS invoices_total, " & _
    "ISNULL(SUM(inv_totalAmt), 0) AS invoices_value, " & _
    "ISNULL(SUM(inv_balance), 0) AS invoices_outstanding " & _
    "FROM tblinvoice WHERE inv_date >= '" & dashboardDateFrom & "' AND inv_date < '" & dashboardDateTo & "' " & _
    "AND (inv_status IS NULL OR inv_status <> 'Cancel')"
rsDashboard.Open sqlDashboard, strconnect, 0, 1
If Not rsDashboard.EOF Then
    invoicesTotal = DashboardNumber(rsDashboard("invoices_total"))
    invoicesValue = DashboardNumber(rsDashboard("invoices_value"))
    invoicesOutstanding = DashboardNumber(rsDashboard("invoices_outstanding"))
End If
rsDashboard.Close

receiptsTotal = 0
receiptsValue = 0
sqlDashboard = "SELECT COUNT(receipt_id) AS receipts_total, " & _
    "ISNULL(SUM(receipt_totalpayment), 0) AS receipts_value " & _
    "FROM tblreceipt WHERE receipt_date >= '" & dashboardDateFrom & "' AND receipt_date < '" & dashboardDateTo & "' " & _
    "AND (receipt_status IS NULL OR receipt_status <> 'Cancel')"
rsDashboard.Open sqlDashboard, strconnect, 0, 1
If Not rsDashboard.EOF Then
    receiptsTotal = DashboardNumber(rsDashboard("receipts_total"))
    receiptsValue = DashboardNumber(rsDashboard("receipts_value"))
End If
rsDashboard.Close

customersTotal = 0
sqlDashboard = "SELECT COUNT(cust_id) AS customers_total FROM tblcustomer " & _
    "WHERE cust_createddate >= '" & dashboardDateFrom & "' AND cust_createddate < '" & dashboardDateTo & "'"
rsDashboard.Open sqlDashboard, strconnect, 0, 1
If Not rsDashboard.EOF Then
    customersTotal = DashboardNumber(rsDashboard("customers_total"))
End If
rsDashboard.Close

stockInPending = 0
stockOutPending = 0
stockTransferPending = 0
stockAdjustmentPending = 0
stockApprovalTotal = 0
sqlDashboard = "SELECT " & _
    "(SELECT COUNT(st_id) FROM tblstockin WHERE st_status = 'Submitted') AS stockin_pending, " & _
    "(SELECT COUNT(so_id) FROM tblstockOut WHERE so_status = 'Submitted') AS stockout_pending, " & _
    "(SELECT COUNT(sf_id) FROM tblstocktransfer WHERE sf_status = 'Submitted') AS stocktransfer_pending, " & _
    "(SELECT COUNT(sj_id) FROM tblstockadj WHERE sj_status = 'Submitted') AS stockadjustment_pending"
rsDashboard.Open sqlDashboard, strconnect, 0, 1
If Not rsDashboard.EOF Then
    stockInPending = DashboardNumber(rsDashboard("stockin_pending"))
    stockOutPending = DashboardNumber(rsDashboard("stockout_pending"))
    stockTransferPending = DashboardNumber(rsDashboard("stocktransfer_pending"))
    stockAdjustmentPending = DashboardNumber(rsDashboard("stockadjustment_pending"))
End If
rsDashboard.Close
stockApprovalTotal = stockInPending + stockOutPending + stockTransferPending + stockAdjustmentPending
Set rsDashboard = Nothing
%>
<style type="text/css">
.crm-dashboard { padding:24px; background:#f4f7fb; font-family:Arial, Helvetica, sans-serif; color:#1f2937; }
.crm-dashboard * { box-sizing:border-box; }
.crm-dashboard-header { display:flex; justify-content:space-between; align-items:flex-end; gap:18px; margin-bottom:20px; }
.crm-dashboard-title { margin:0; color:#183b66; font-size:25px; line-height:1.2; }
.crm-dashboard-subtitle { margin:6px 0 0; color:#6b7280; font-size:13px; }
.crm-dashboard-date { padding:10px 14px; border:1px solid #d9e2ef; border-radius:8px; background:#fff; color:#334155; font-size:13px; font-weight:bold; white-space:nowrap; }
.crm-kpi-grid { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:14px; margin-bottom:18px; }
.crm-kpi-card { position:relative; min-height:132px; padding:18px; overflow:hidden; border:1px solid #e2e8f0; border-radius:10px; background:#fff; box-shadow:0 3px 12px rgba(15, 23, 42, .06); color:inherit; text-decoration:none; }
.crm-kpi-card:hover { border-color:#93b4d8; box-shadow:0 6px 18px rgba(15, 23, 42, .10); }
.crm-kpi-accent { position:absolute; top:0; right:0; width:6px; height:100%; }
.crm-kpi-label { color:#64748b; font-size:12px; font-weight:bold; letter-spacing:.5px; text-transform:uppercase; }
.crm-kpi-value { margin-top:8px; color:#0f2947; font-size:31px; font-weight:bold; line-height:1; }
.crm-kpi-note { margin-top:10px; color:#64748b; font-size:12px; line-height:1.45; }
.crm-money { font-size:24px; }
.crm-panel-grid { display:grid; grid-template-columns:1.45fr 1fr; gap:14px; }
.crm-approval-panel { margin-bottom:18px; border:1px solid #e2e8f0; border-radius:10px; background:#fff; box-shadow:0 3px 12px rgba(15, 23, 42, .05); }
.crm-approval-grid { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:10px; padding:16px; }
.crm-approval-card { display:flex; align-items:center; justify-content:space-between; gap:12px; padding:15px; border:1px solid #f0d8a8; border-radius:8px; background:#fffaf0; color:#334155; text-decoration:none; }
.crm-approval-card:hover { border-color:#d9a441; background:#fff5dd; }
.crm-approval-label { display:block; color:#7c5a16; font-size:12px; font-weight:bold; text-transform:uppercase; }
.crm-approval-note { display:block; margin-top:5px; color:#8a7652; font-size:11px; }
.crm-approval-number { min-width:38px; color:#b45309; font-size:25px; font-weight:bold; text-align:right; }
.crm-approval-total { color:#b45309; font-size:12px; font-weight:bold; }
.crm-panel { border:1px solid #e2e8f0; border-radius:10px; background:#fff; box-shadow:0 3px 12px rgba(15, 23, 42, .05); }
.crm-panel-header { display:flex; justify-content:space-between; align-items:center; padding:16px 18px; border-bottom:1px solid #edf2f7; }
.crm-panel-title { margin:0; color:#183b66; font-size:16px; }
.crm-panel-link { color:#2563a6; font-size:12px; font-weight:bold; text-decoration:none; }
.crm-status-grid { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:10px; padding:16px; }
.crm-status { display:block; padding:13px 10px; border:1px solid #e5eaf1; border-radius:8px; background:#f8fafc; color:#334155; text-align:center; text-decoration:none; }
.crm-status:hover { background:#eef5fc; border-color:#aac5e1; }
.crm-status-number { display:block; margin-bottom:5px; color:#183b66; font-size:22px; font-weight:bold; }
.crm-status-label { display:block; font-size:11px; font-weight:bold; text-transform:uppercase; }
.crm-finance-list { padding:8px 18px 14px; }
.crm-finance-row { display:flex; justify-content:space-between; gap:16px; padding:13px 0; border-bottom:1px solid #edf2f7; }
.crm-finance-row:last-child { border-bottom:0; }
.crm-finance-label { color:#64748b; font-size:13px; }
.crm-finance-value { color:#183b66; font-size:15px; font-weight:bold; text-align:right; }
.crm-quick-links { display:flex; flex-wrap:wrap; gap:8px; padding:0 18px 18px; }
.crm-quick-link { padding:8px 11px; border-radius:6px; background:#eaf2fb; color:#215b91; font-size:12px; font-weight:bold; text-decoration:none; }
.crm-quick-link:hover { background:#d8e8f8; }
@media (max-width:1050px) {
    .crm-kpi-grid { grid-template-columns:repeat(2, minmax(0, 1fr)); }
    .crm-approval-grid { grid-template-columns:repeat(2, minmax(0, 1fr)); }
    .crm-panel-grid { grid-template-columns:1fr; }
}
@media (max-width:680px) {
    .crm-dashboard { padding:14px; }
    .crm-dashboard-header { display:block; }
    .crm-dashboard-date { display:inline-block; margin-top:12px; }
    .crm-kpi-grid { grid-template-columns:1fr; }
    .crm-approval-grid { grid-template-columns:1fr; }
    .crm-status-grid { grid-template-columns:repeat(2, minmax(0, 1fr)); }
}
</style>
<tr>
  <td>
    <div class="crm-dashboard">
      <div class="crm-dashboard-header">
        <div>
          <h1 class="crm-dashboard-title">CRM Daily Dashboard</h1>
          <p class="crm-dashboard-subtitle">A live view of today's customer service and financial activity.</p>
        </div>
        <div class="crm-dashboard-date">Today: <%=dashboardDateLabel%></div>
      </div>

      <div class="crm-kpi-grid">
        <a class="crm-kpi-card" href="rm_jobsheet_view.asp?job_status=All&amp;job_date_from=<%=dashboardDateLink%>&amp;job_date_to=<%=dashboardDateLink%>">
          <span class="crm-kpi-accent" style="background:#2563a6;"></span>
          <div class="crm-kpi-label">Jobsheets Today</div>
          <div class="crm-kpi-value"><%=DashboardCount(jobsTotal)%></div>
          <div class="crm-kpi-note"><%=DashboardCount(jobsOpen + jobsSubmitted + jobsAccepted)%> open actions across all records</div>
        </a>
        <a class="crm-kpi-card" href="rm_invoice_view.asp?inv_status=All&amp;inv_createddate_from=<%=dashboardDateLink%>&amp;inv_createddate_to=<%=dashboardDateLink%>">
          <span class="crm-kpi-accent" style="background:#7c3aed;"></span>
          <div class="crm-kpi-label">Invoices Issued</div>
          <div class="crm-kpi-value"><%=DashboardCount(invoicesTotal)%></div>
          <div class="crm-kpi-note">RM <%=FormatNumber(invoicesValue, 2)%> total invoice value</div>
        </a>
        <a class="crm-kpi-card" href="rm_receipt_view.asp?receipt_status=All&amp;receipt_createddate_from=<%=dashboardDateLink%>&amp;receipt_createddate_to=<%=dashboardDateLink%>">
          <span class="crm-kpi-accent" style="background:#059669;"></span>
          <div class="crm-kpi-label">Receipts Issued</div>
          <div class="crm-kpi-value"><%=DashboardCount(receiptsTotal)%></div>
          <div class="crm-kpi-note">RM <%=FormatNumber(receiptsValue, 2)%> collected today</div>
        </a>
        <a class="crm-kpi-card" href="rm_customer_view.asp">
          <span class="crm-kpi-accent" style="background:#d97706;"></span>
          <div class="crm-kpi-label">New Customers</div>
          <div class="crm-kpi-value"><%=DashboardCount(customersTotal)%></div>
          <div class="crm-kpi-note">Customer records created today</div>
        </a>
      </div>

      <div class="crm-approval-panel">
        <div class="crm-panel-header">
          <h2 class="crm-panel-title">Stock Pending Approval</h2>
          <span class="crm-approval-total"><%=DashboardCount(stockApprovalTotal)%> submitted transaction<% If stockApprovalTotal <> 1 Then Response.Write "s" End If %></span>
        </div>
        <div class="crm-approval-grid">
          <a class="crm-approval-card" href="rm_stockin_view.asp?st_status=Submitted&amp;st_createddate_from=01-Jan-2000&amp;st_createddate_to=<%=dashboardDateLink%>">
            <span><span class="crm-approval-label">Stock-In</span><span class="crm-approval-note">Awaiting approval</span></span>
            <span class="crm-approval-number"><%=DashboardCount(stockInPending)%></span>
          </a>
          <a class="crm-approval-card" href="rm_stockOut_view.asp?so_status=Submitted&amp;so_createddate_from=01-Jan-2000&amp;so_createddate_to=<%=dashboardDateLink%>">
            <span><span class="crm-approval-label">Stock-Out</span><span class="crm-approval-note">Awaiting approval</span></span>
            <span class="crm-approval-number"><%=DashboardCount(stockOutPending)%></span>
          </a>
          <a class="crm-approval-card" href="rm_stocktfr_view.asp?sf_status=Submitted&amp;sf_createddate_from=01-Jan-2000&amp;sf_createddate_to=<%=dashboardDateLink%>">
            <span><span class="crm-approval-label">Stock-Transfer</span><span class="crm-approval-note">Awaiting approval</span></span>
            <span class="crm-approval-number"><%=DashboardCount(stockTransferPending)%></span>
          </a>
          <a class="crm-approval-card" href="rm_stockAdj_view.asp?sj_status=Submitted&amp;sj_createddate_from=01-Jan-2000&amp;sj_createddate_to=<%=dashboardDateLink%>">
            <span><span class="crm-approval-label">Stock-Adjustment</span><span class="crm-approval-note">Awaiting approval</span></span>
            <span class="crm-approval-number"><%=DashboardCount(stockAdjustmentPending)%></span>
          </a>
        </div>
      </div>

      <div class="crm-panel-grid">
        <div class="crm-panel">
          <div class="crm-panel-header">
            <h2 class="crm-panel-title">Jobsheet Status - All Records</h2>
            <a class="crm-panel-link" href="rm_jobsheet_view.asp?job_status=All&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>">View all jobs</a>
          </div>
          <div class="crm-status-grid">
            <a class="crm-status" href="rm_jobsheet_view.asp?job_status=Open&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>"><span class="crm-status-number"><%=DashboardCount(jobsOpen)%></span><span class="crm-status-label">Open</span></a>
            <a class="crm-status" href="rm_jobsheet_view.asp?job_status=Submitted&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>"><span class="crm-status-number"><%=DashboardCount(jobsSubmitted)%></span><span class="crm-status-label">Submitted</span></a>
            <a class="crm-status" href="rm_jobsheet_view.asp?job_status=Accepted&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>"><span class="crm-status-number"><%=DashboardCount(jobsAccepted)%></span><span class="crm-status-label">Accepted</span></a>
            <a class="crm-status" href="rm_jobsheet_view.asp?job_status=Done&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>"><span class="crm-status-number"><%=DashboardCount(jobsDone)%></span><span class="crm-status-label">Done</span></a>
            <a class="crm-status" href="rm_jobsheet_view.asp?job_status=Posted&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>"><span class="crm-status-number"><%=DashboardCount(jobsPosted)%></span><span class="crm-status-label">Posted</span></a>
            <a class="crm-status" href="rm_jobsheet_view.asp?job_status=Cancel&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>"><span class="crm-status-number"><%=DashboardCount(jobsCancelled)%></span><span class="crm-status-label">Cancelled</span></a>
            <a class="crm-status" href="rm_jobsheet_view.asp?job_status=All&amp;job_date_from=01-Jan-2000&amp;job_date_to=<%=dashboardDateLink%>"><span class="crm-status-number"><%=DashboardCount(jobsStatusTotal)%></span><span class="crm-status-label">Total</span></a>
          </div>
        </div>

        <div class="crm-panel">
          <div class="crm-panel-header">
            <h2 class="crm-panel-title">Today's Financial Snapshot</h2>
          </div>
          <div class="crm-finance-list">
            <div class="crm-finance-row"><span class="crm-finance-label">Invoice value issued</span><span class="crm-finance-value">RM <%=FormatNumber(invoicesValue, 2)%></span></div>
            <div class="crm-finance-row"><span class="crm-finance-label">Outstanding on today's invoices</span><span class="crm-finance-value">RM <%=FormatNumber(invoicesOutstanding, 2)%></span></div>
            <div class="crm-finance-row"><span class="crm-finance-label">Receipts collected</span><span class="crm-finance-value">RM <%=FormatNumber(receiptsValue, 2)%></span></div>
          </div>
          <div class="crm-quick-links">
            <a class="crm-quick-link" href="rm_jobsheet.asp">New Jobsheet</a>
            <a class="crm-quick-link" href="rm_invoice_new.asp">New Invoice</a>
            <a class="crm-quick-link" href="rm_customer_new.asp">New Customer</a>
          </div>
        </div>
      </div>
    </div>
  </td>
</tr>
<!-- #include file="footer.asp" -->

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

Function DashboardText(value)
    If IsNull(value) Then
        DashboardText = ""
    Else
        DashboardText = Server.HTMLEncode(CStr(value))
    End If
End Function

Dim dashboardDateLabel, dashboardDateLink
dashboardDateLabel = ChkDate(Date())
dashboardDateLink = Server.URLEncode(dashboardDateLabel)

Dim jobsTotal, jobsStatusTotal, jobsOpen, jobsSubmitted, jobsAccepted, jobsDone, jobsPosted, jobsCancelled
Dim invoicesTotal, invoicesValue, invoicesOutstanding
Dim receiptsTotal, receiptsValue, customersTotal
Dim stockInPending, stockOutPending, stockTransferPending, stockAdjustmentPending, stockApprovalTotal
Dim conDashboard, cmdDashboard, rsDashboard, topCitiesData, topPartsData, topTechniciansData, cityRow, partRow, technicianRow

jobsTotal = 0
jobsStatusTotal = 0
jobsOpen = 0
jobsSubmitted = 0
jobsAccepted = 0
jobsDone = 0
jobsPosted = 0
jobsCancelled = 0
invoicesTotal = 0
invoicesValue = 0
invoicesOutstanding = 0
receiptsTotal = 0
receiptsValue = 0
customersTotal = 0
stockInPending = 0
stockOutPending = 0
stockTransferPending = 0
stockAdjustmentPending = 0
stockApprovalTotal = 0
topCitiesData = Empty
topPartsData = Empty
topTechniciansData = Empty

' One stored-procedure call returns all dashboard result sets.
Set conDashboard = Server.CreateObject("ADODB.Connection")
conDashboard.Open strconnect
Set cmdDashboard = Server.CreateObject("ADODB.Command")
Set cmdDashboard.ActiveConnection = conDashboard
cmdDashboard.CommandText = "dbo.usp_CrmDashboard"
cmdDashboard.CommandType = 4
cmdDashboard.CommandTimeout = 30
cmdDashboard.Parameters.Append cmdDashboard.CreateParameter("@AsOfDate", 135, 1, , Now())
Set rsDashboard = cmdDashboard.Execute

' Result 1: today's jobsheets.
If Not rsDashboard.EOF Then
    jobsTotal = DashboardNumber(rsDashboard("jobs_total"))
End If
Set rsDashboard = rsDashboard.NextRecordset

' Result 2: all-record jobsheet workflow.
If Not rsDashboard.EOF Then
    jobsStatusTotal = DashboardNumber(rsDashboard("jobs_total"))
    jobsOpen = DashboardNumber(rsDashboard("jobs_open"))
    jobsSubmitted = DashboardNumber(rsDashboard("jobs_submitted"))
    jobsAccepted = DashboardNumber(rsDashboard("jobs_accepted"))
    jobsDone = DashboardNumber(rsDashboard("jobs_done"))
    jobsPosted = DashboardNumber(rsDashboard("jobs_posted"))
    jobsCancelled = DashboardNumber(rsDashboard("jobs_cancelled"))
End If
Set rsDashboard = rsDashboard.NextRecordset

' Result 3: daily finance/customer KPIs and stock approval backlog.
If Not rsDashboard.EOF Then
    invoicesTotal = DashboardNumber(rsDashboard("invoices_total"))
    invoicesValue = DashboardNumber(rsDashboard("invoices_value"))
    invoicesOutstanding = DashboardNumber(rsDashboard("invoices_outstanding"))
    receiptsTotal = DashboardNumber(rsDashboard("receipts_total"))
    receiptsValue = DashboardNumber(rsDashboard("receipts_value"))
    customersTotal = DashboardNumber(rsDashboard("customers_total"))
    stockInPending = DashboardNumber(rsDashboard("stockin_pending"))
    stockOutPending = DashboardNumber(rsDashboard("stockout_pending"))
    stockTransferPending = DashboardNumber(rsDashboard("stocktransfer_pending"))
    stockAdjustmentPending = DashboardNumber(rsDashboard("stockadjustment_pending"))
End If
stockApprovalTotal = stockInPending + stockOutPending + stockTransferPending + stockAdjustmentPending
Set rsDashboard = rsDashboard.NextRecordset

' Result 4: Top 10 Cities. Materialize it before advancing to the next result set.
If Not rsDashboard.EOF Then
    topCitiesData = rsDashboard.GetRows()
End If
Set rsDashboard = rsDashboard.NextRecordset

' Result 5: Top 10 Parts Replaced.
If Not rsDashboard.EOF Then
    topPartsData = rsDashboard.GetRows()
End If
Set rsDashboard = rsDashboard.NextRecordset

' Result 6: Top 10 Technicians.
If Not rsDashboard.EOF Then
    topTechniciansData = rsDashboard.GetRows()
End If
rsDashboard.Close
Set rsDashboard = Nothing
Set cmdDashboard = Nothing
conDashboard.Close
Set conDashboard = Nothing
%>
<style type="text/css">
.crm-dashboard { padding:16px; background:#f4f7fb; font-family:Arial, Helvetica, sans-serif; color:#1f2937; }
.crm-dashboard * { box-sizing:border-box; }
.crm-dashboard-header { display:flex; justify-content:space-between; align-items:flex-end; gap:12px; margin-bottom:12px; }
.crm-dashboard-title { margin:0; color:#183b66; font-size:20px; line-height:1.2; }
.crm-dashboard-subtitle { margin:3px 0 0; color:#6b7280; font-size:11px; }
.crm-dashboard-date { padding:7px 10px; border:1px solid #d9e2ef; border-radius:7px; background:#fff; color:#334155; font-size:11px; font-weight:bold; white-space:nowrap; }
.crm-kpi-grid { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:10px; margin-bottom:12px; }
.crm-kpi-card { position:relative; min-height:94px; padding:12px; overflow:hidden; border:1px solid #e2e8f0; border-radius:8px; background:#fff; box-shadow:0 3px 12px rgba(15, 23, 42, .06); color:inherit; text-decoration:none; }
.crm-kpi-card:hover { border-color:#93b4d8; box-shadow:0 6px 18px rgba(15, 23, 42, .10); }
.crm-kpi-accent { position:absolute; top:0; right:0; width:5px; height:100%; }
.crm-kpi-label { color:#64748b; font-size:11px; font-weight:bold; letter-spacing:.3px; text-transform:uppercase; }
.crm-kpi-value { margin-top:6px; color:#0f2947; font-size:18px; font-weight:bold; line-height:1; }
.crm-kpi-note { margin-top:7px; color:#64748b; font-size:11px; line-height:1.3; }
.crm-money { font-size:18px; }
.crm-panel-grid { display:grid; grid-template-columns:1.45fr 1fr; gap:10px; }
.crm-approval-panel { margin-bottom:12px; border:1px solid #e2e8f0; border-radius:8px; background:#fff; box-shadow:0 3px 12px rgba(15, 23, 42, .05); }
.crm-approval-grid { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:8px; padding:10px; }
.crm-approval-card { display:flex; align-items:center; justify-content:space-between; gap:8px; padding:10px; border:1px solid #f0d8a8; border-radius:7px; background:#fffaf0; color:#334155; text-decoration:none; }
.crm-approval-card:hover { border-color:#d9a441; background:#fff5dd; }
.crm-approval-label { display:block; color:#7c5a16; font-size:11px; font-weight:bold; text-transform:uppercase; }
.crm-approval-note { display:block; margin-top:3px; color:#8a7652; font-size:10px; }
.crm-approval-number { min-width:28px; color:#b45309; font-size:18px; font-weight:bold; text-align:right; }
.crm-approval-total { color:#b45309; font-size:11px; font-weight:bold; }
.crm-panel { border:1px solid #e2e8f0; border-radius:8px; background:#fff; box-shadow:0 3px 12px rgba(15, 23, 42, .05); }
.crm-panel-header { display:flex; justify-content:space-between; align-items:center; padding:12px 14px; border-bottom:1px solid #edf2f7; }
.crm-panel-title { margin:0; color:#183b66; font-size:14px; }
.crm-panel-link { color:#2563a6; font-size:11px; font-weight:bold; text-decoration:none; }
.crm-status-grid { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:8px; padding:10px; }
.crm-status { display:block; padding:9px 8px; border:1px solid #e5eaf1; border-radius:6px; background:#f8fafc; color:#334155; text-align:center; text-decoration:none; }
.crm-status:hover { background:#eef5fc; border-color:#aac5e1; }
.crm-status-number { display:block; margin-bottom:3px; color:#183b66; font-size:18px; font-weight:bold; }
.crm-status-label { display:block; font-size:10px; font-weight:bold; text-transform:uppercase; }
.crm-finance-list { padding:4px 14px 8px; }
.crm-finance-row { display:flex; justify-content:space-between; gap:12px; padding:9px 0; border-bottom:1px solid #edf2f7; }
.crm-finance-row:last-child { border-bottom:0; }
.crm-finance-label { color:#64748b; font-size:11px; }
.crm-finance-value { color:#183b66; font-size:13px; font-weight:bold; text-align:right; }
.crm-quick-links { display:flex; flex-wrap:wrap; gap:6px; padding:0 14px 12px; }
.crm-quick-link { padding:6px 8px; border-radius:5px; background:#eaf2fb; color:#215b91; font-size:11px; font-weight:bold; text-decoration:none; }
.crm-quick-link:hover { background:#d8e8f8; }
.crm-insight-grid { display:grid; grid-template-columns:repeat(3, minmax(0, 1fr)); gap:10px; margin-top:12px; }
.crm-ranking-list { padding:8px 18px 14px; }
.crm-ranking-row { display:grid; grid-template-columns:28px minmax(0, 1fr) auto; align-items:center; gap:10px; min-height:42px; border-bottom:1px solid #edf2f7; }
.crm-ranking-row:last-child { border-bottom:0; }
.crm-ranking-position { display:inline-block; width:24px; height:24px; border-radius:50%; background:#eaf2fb; color:#215b91; font-size:11px; font-weight:bold; line-height:24px; text-align:center; }
.crm-ranking-name { min-width:0; color:#334155; font-size:13px; font-weight:bold; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.crm-ranking-description { display:block; margin-top:3px; color:#7a8797; font-size:11px; font-weight:normal; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.crm-ranking-value { padding:5px 8px; border-radius:5px; background:#f8fafc; color:#183b66; font-size:13px; font-weight:bold; white-space:nowrap; }
.crm-empty-state { padding:28px 18px; color:#7a8797; font-size:13px; text-align:center; }
.crm-period-label { color:#64748b; font-size:11px; }
@media (max-width:1050px) {
    .crm-kpi-grid { grid-template-columns:repeat(2, minmax(0, 1fr)); }
    .crm-approval-grid { grid-template-columns:repeat(2, minmax(0, 1fr)); }
    .crm-panel-grid { grid-template-columns:1fr; }
    .crm-insight-grid { grid-template-columns:1fr; }
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

      <div class="crm-insight-grid">
        <div class="crm-panel">
          <div class="crm-panel-header">
            <h2 class="crm-panel-title">Top 10 Cities</h2>
            <span class="crm-period-label">Last 3 months</span>
          </div>
          <div class="crm-ranking-list">
            <%
            If Not IsArray(topCitiesData) Then
            %>
              <div class="crm-empty-state">No city activity found for this period.</div>
            <%
            Else
                For cityRow = 0 To UBound(topCitiesData, 2)
            %>
              <div class="crm-ranking-row">
                <span class="crm-ranking-position"><%=cityRow + 1%></span>
                <span class="crm-ranking-name"><%=DashboardText(topCitiesData(0, cityRow))%></span>
                <span class="crm-ranking-value"><%=DashboardCount(topCitiesData(1, cityRow))%> jobs</span>
              </div>
            <%
                Next
            End If
            %>
          </div>
        </div>

        <div class="crm-panel">
          <div class="crm-panel-header">
            <h2 class="crm-panel-title">Top 10 Parts Replaced</h2>
            <span class="crm-period-label">Last 3 months</span>
          </div>
          <div class="crm-ranking-list">
            <%
            If Not IsArray(topPartsData) Then
            %>
              <div class="crm-empty-state">No replaced parts found for this period.</div>
            <%
            Else
                For partRow = 0 To UBound(topPartsData, 2)
            %>
              <div class="crm-ranking-row">
                <span class="crm-ranking-position"><%=partRow + 1%></span>
                <span class="crm-ranking-name"><%=DashboardText(topPartsData(0, partRow))%><span class="crm-ranking-description"><%=DashboardText(topPartsData(1, partRow))%></span></span>
                <span class="crm-ranking-value"><%=DashboardCount(topPartsData(2, partRow))%> replaced</span>
              </div>
            <%
                Next
            End If
            %>
          </div>
        </div>

        <div class="crm-panel">
          <div class="crm-panel-header">
            <h2 class="crm-panel-title">Top 10 Technicians</h2>
            <span class="crm-period-label">Last 3 months</span>
          </div>
          <div class="crm-ranking-list">
            <%
            If Not IsArray(topTechniciansData) Then
            %>
              <div class="crm-empty-state">No posted technician jobs found for this period.</div>
            <%
            Else
                For technicianRow = 0 To UBound(topTechniciansData, 2)
            %>
              <div class="crm-ranking-row">
                <span class="crm-ranking-position"><%=technicianRow + 1%></span>
                <span class="crm-ranking-name"><%=DashboardText(topTechniciansData(0, technicianRow))%><span class="crm-ranking-description"><%=DashboardText(topTechniciansData(1, technicianRow))%></span></span>
                <span class="crm-ranking-value"><%=DashboardCount(topTechniciansData(2, technicianRow))%> jobs</span>
              </div>
            <%
                Next
            End If
            %>
          </div>
        </div>
      </div>
    </div>
  </td>
</tr>
<!-- #include file="footer.asp" -->

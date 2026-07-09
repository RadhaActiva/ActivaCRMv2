<!-- #include file="database/datastore.asp" -->

<%
job_date_from = request("job_date_from")
job_date_to=request("job_date_to") 
job_status=request("job_status") 
jobrange=request("jobrange") 
jobrangestatus=request("jobrangestatus") 
job_tech_code=request("job_tech_code") 

%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>
	
<%
i = 1
sql2 = "SELECT tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
		"tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
		"tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
		"tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
		"tbljob.job_Model, tbljob.job_Model_desc, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
		"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
		"tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN, tbljob.job_tech_faulty_code, " & _
		"tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
		"tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo, tbljob.job_hq_remark,  " & _
		"tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
		"tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
		"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbltechnician.tech_name, tbltechnician.tech_tel1, " & _
		"tbljob.job_submitteddate, tbljob.job_submittedby, tbljob.job_donedate, tbljob.job_posteddate, tbljob.job_cancelleddate " & _
		"FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
	    "and tbljob.job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' "

if job_tech_code <> "All" then 
sql2 = sql2 & " and job_tech_code = '" & job_tech_code & "' "
end if
		
if jobrangestatus="completed" then 
   sql2 = sql2 & " and tbljob.job_status in ('Done', 'Posted') "

	if jobrange="1t3" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate,job_donedate) < 4 and DATEDIFF(day, job_submitteddate, job_donedate) >= 0"
	elseif jobrange="4t7" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate,job_donedate) > 3 and DATEDIFF(day,job_submitteddate,job_donedate) < 8 "
	elseif jobrange="8t14" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate,job_donedate) > 7 and DATEDIFF(day,job_submitteddate,job_donedate) < 15 "
	elseif jobrange="more14" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate,job_donedate) > 14 "
	elseif jobrange="incompletedate" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate,job_donedate) is null "
	elseif jobrange="all" then 
	'sql2 = sql2 & " and DATEDIFF(job_donedate,job_submitteddate) > 0 "
	end if

elseif jobrangestatus="pending" then 
    sql2 = sql2 & " and tbljob.job_status in ('Submitted', 'Accepted') "
	
	if jobrange="1t3" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) < 4 "
	elseif jobrange="4t7" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) > 3 and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) < 8 "
	elseif jobrange="8t14" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) > 7 and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) < 15 "
	elseif jobrange="more14" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) > 14 "
	elseif jobrange="incompletedate" then 
	sql2 = sql2 & " and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) is null "
	elseif jobrange="all" then 
	'sql2 = sql2 & " and DATEDIFF(job_JS_receiveddate,job_submitteddate) is not null "
	end if
	
end if
		


sql2 = sql2 & " order by tbljob.job_code"

response.Cookies("GAPS")("sqlexcel") = sql2
set rs = server.CreateObject("adodb.recordset")
set rs1 = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if rs.eof then
   norecord = "There is no record found."
end if

If Not rs.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs.RecordCount
Remain = TotalRecord - Showed

If Remain > row Then
  LoopMax = Showed + row
Else
  LoopMax = Showed + Remain
End If

	If Int(TotalRecord/row) <> TotalRecord/row Then
	  pgCount = Int(TotalRecord/row) + 1
	Else
	  pgCount = TotalRecord/row
	End If

	if LoopMax mod row = 0 then
		pagestartno = LoopMax/row
	else
		pagestartno = pgCount
	end if		
end if

count = count + Showed
link = "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & "&job_status=" & job_status & "&jobrange=" & jobrange & "&jobrangestatus=" & jobrangestatus & "&job_tech_code=" & job_tech_code   
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Job Listing</strong></font></td>
    <td align="right" class="style21"><span class="titlegrey1"><a href="rm_rpt_tech_servicekpi_detail_excel.asp?" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></span></td>
  </tr>
  <tr> 
    <td colspan="2" align="right" valign="top"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%>
      </font>of <font color="3366ff"> 
      <%=pgCount%>
      </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_tech_servicekpi_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_tech_servicekpi_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td colspan="2" valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#FFFF00"> 
          <td width="3%"><strong>No.</strong></td>
          <td><strong>Job No</strong></td>
          <td><strong>Status</strong></td>
          <td align="left" class='tktTotals'><strong>Open Date</strong></td>
          <td class='tktTotals'><strong>Submitted Date</strong></td>
          <td class='tktTotals'><strong>Accepted Date</strong></td>
          <td class='tktTotals'><strong>Done Date</strong></td>
          <td class='tktTotals'><strong>Posted Date</strong></td>
          <td class='tktTotals'><strong>Cancel Date</strong></td>
          <td class='tktTotals'><strong>Technician</strong></td>
          <td class='tktTotals'><strong>Technician Code</strong></td>
          <td class='tktTotals'><strong>Created By</strong></td>
          <td class='tktTotals'><strong>Faulty Reason</strong></td>
          <td class='tktTotals'><strong>Faulty Desc</strong></td>
        </tr>
        <% 
if not rs.eof then
rs.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
        
        <tr valign="top" bgcolor="<%=nbgcolor%>"> 
          <td nowrap><%=count%>.</td>
          <td nowrap><%=rs("job_code")%></td>
          <td><%=rs("job_status")%></td>
          <td align="left" nowrap><%=chkdate(rs("job_createddate"))%></td>
          <td><%=chkdate(rs("job_submitteddate"))%></td>
          <td><%=chkdate(rs("job_JS_receiveddate"))%></td>
          <td><%=chkdate(rs("job_donedate"))%></td>
          <td><%=chkdate(rs("job_posteddate"))%></td>
          <td><%=chkdate(rs("job_cancelleddate"))%></td>
          <td><%=rs("tech_name")%></td>
          <td><%=rs("job_tech_code")%></td>
          <td><%=rs("job_createdby")%></td>
          <td><%=rs("job_tech_faulty_reason")%></td>
          <td><%=rs("job_tech_faulty_action")%></td>
        </tr>
        <%
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
    </table></td>
  </tr>
  <tr valign="top"> 
    <td colspan="10" align="right"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_tech_servicekpi_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_tech_servicekpi_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

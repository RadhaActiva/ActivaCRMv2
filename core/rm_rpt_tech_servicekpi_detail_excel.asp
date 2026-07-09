<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=rm_rpt_tech_servicekpi_detail_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="1" cellpadding="4" cellspacing="0">
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
          <td class='tktTotals'><strong>Created by</strong></td>
          <td class='tktTotals'><strong>Submitted by</strong></td>            
          <td class='tktTotals'><strong>Faulty Reason</strong></td>
          <td class='tktTotals'><strong>Faulty Desc</strong></td>
        </tr>
  <%
i = 1  
sql2 = request.Cookies("GAPS")("sqlexcel")  
'sql2 = "select * from tbljob where job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "'"	
       
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open

while not rs.eof 
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
          <td><%=rs("job_submittedby")%></td>
          <td><%=rs("job_tech_faulty_reason")%></td>
          <td><%=rs("job_tech_faulty_action")%></td>
        </tr>
  <%
count = count + 1 
i = i + 1
rs.MoveNext
wend
rs.Close

%>
   
</table>

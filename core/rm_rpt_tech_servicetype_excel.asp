<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=servicetype_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<%
if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if request("jobmonth") <> "" then
   jobmonth = request("jobmonth")
else
   jobmonth = month(date())
end if

if request("jobyear") <> "" then
   jobyear = request("jobyear")
else
   jobyear = year(date())
end if
%>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Service Type Summary</div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_tech_servicetype.asp">
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Job Date<br />
          </strong></td>
          <td width="38"><div align="left">
            <%=job_date_from%> to <%=job_date_to%>  <span class="titlegrey1">            </span></div></td>
        </tr>
      </table>
    </form></td>
  </tr>
  <tr>
    <td colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td width="43%" valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
      <tr>
        <td width="9%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
        <td width="75%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Service Status</span></strong></font></td>
        <td width="16%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
      </tr>
      <%
i = 1
sql1 = "select job_status, count(job_id) as totalcnt from tbljob where " & _
       "job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _	
       "group by job_status  "

totalcnt = 0
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
      <tr>
        <td height="40" align="center"><%=i%></td>
        <td align="left"><%=rs1("job_status")%></td>
        <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs1("totalcnt")%></td>
      </tr>
      <%
 if rs1("job_status") = "Posted" then 
 %>
      <%
    withouappointment = 0
    sql1 = "select count(job_id) from tbljob where job_status='Posted' and job_tech_code='resolved_no_appt' " & _
	       "and job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
    withouappointment =  selectid(sql1)
    %>
      <tr>
        <td height="40" align="center">&nbsp;</td>
        <td align="left"><strong> - Resolved Issue without Appointment = <%=withouappointment%></strong></td>
        <td align="center" nowrap="nowrap" bgcolor="#F3F3F3">&nbsp;</td>
      </tr>
      <%
	withppointment=0
	sql1 = "select count(job_id) from tbljob where job_status='Posted' and job_tech_code<>'resolved_no_appt' " & _
	       "and job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
    withppointment =  selectid(sql1)
    %>
      <tr>
        <td height="40" align="center">&nbsp;</td>
        <td align="left"><strong>- Resolved Issue with Techician Appointment = <%=withppointment%></strong></td>
        <td align="center" nowrap="nowrap" bgcolor="#F3F3F3">&nbsp;</td>
      </tr>
      <%
 end if
 %>
      <%
totalcnt = totalcnt + cint(rs1("totalcnt"))
i = i + 1
rs1.movenext
wend
rs1.close
totalcnt = totalcnt 
%>
      <tr bgcolor="#F3F3F3">
        <td height="40" colspan="2" align="right" bgcolor="#FFFFFF"><strong>Total</strong></td>
        <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=totalcnt%></strong></td>
      </tr>
    </table></td>
    <td width="57%" valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
      <tr>
        <td width="8%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
        <td width="80%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Service Type</span></strong></font></td>
        <td width="12%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
      </tr>
      <%
i = 1
sql1 = "select job_reportedby, count(job_id) as totalcnt from tbljob where " & _
       "job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _	 
       " group by job_reportedby  "
	   
totalcnt = 0
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
      <tr bgcolor="<%=nbgcolor%>">
        <td height="40" align="center"><%=i%></td>
        <td align="left"><strong> <font color="#0000FF"><%=rs1("job_reportedby")%></font></strong></td>
        <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs1("totalcnt")%></font></strong></strong></td>
      </tr>
      <%

totalcnt = totalcnt + cint(rs1("totalcnt"))
i = i + 1
rs1.movenext
wend
rs1.close
%>
      <tr bgcolor="#F3F3F3">
        <td height="40" colspan="2" align="right" bgcolor="#FFFFFF"><strong>Total</strong></td>
        <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=totalcnt%></strong></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
</table>

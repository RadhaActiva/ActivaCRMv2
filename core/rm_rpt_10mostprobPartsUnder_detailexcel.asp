<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=most10UnderParts_Detail_" & year(date()) & month(date()) & day(date()) & ".xls"


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

%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                 <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td><strong>Job No</strong></td>
          <td><strong>Status</strong></td>
          <td><strong>Job Date</strong></td>
          <td align="left" class='tktTotals'><strong>Model</strong></td>
          <td class='tktTotals'><strong>Wrty</strong></td>
          <td class='tktTotals'><strong>Part Code </strong></td>
          <td align="center" class='tktTotals'><strong>Qty</strong></td>
          <td class='tktTotals'><strong>Unit Price</strong></td>
          <td class='tktTotals'><strong>Subtotal</strong></td>
          <td class='tktTotals'><strong>Faulty</strong></td>
          <td class='tktTotals'><strong>Faulty Desc</strong></td>
          <td class='tktTotals'><strong>Category Code</strong></td>
        </tr>
              <%
sql2 = request.Cookies("GAPS")("sqlexcel2")

i = 1	   
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
while not rs.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
          
           <tr valign="top" bgcolor="<%=nbgcolor%>"> 
          <td nowrap><%=i%>.</td>
          <td nowrap><%=rs("job_code")%></td>
          <td><%=rs("job_status")%></td>
          <td nowrap><%=chkdate(rs("job_date"))%></td>
          <td align="left" nowrap><%=rs("job_tech_model")%></td>
          <td><%=rs("job_actual_wrty_status")%></td>
          <td><%=rs("job_code")%></td>
          <td align="center"><%=rs("jobp_qty")%></td>
          <td><%=rs("jobp_unitcost")%></td>
          <td><%=rs("jobp_subtotal")%></td>
          <td><%=rs("jobp_faultycode")%></td>
          <td><%=rs("fr_description")%></td>
          <td><%=rs("job_hq_category_code")%></td>
        </tr>
              <%
totaljob = totaljob + cint(rs("jobp_qty"))
i = i + 1
rs.movenext
wend
rs.close
%>
 <tr valign="top" bgcolor="<%=nbgcolor%>">
                <td colspan="7" align="right" nowrap>Total</td>
                <td align="center"><%=totaljob%></td>
                <td colspan="5">&nbsp;</td>
  </tr>
 
</table>

<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=most10OverParts_detail_" & year(date()) & month(date()) & day(date()) & ".xls"


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
        <tr> 
    <td colspan="2" valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
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
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%> 
             
              <tr valign="top" bgcolor="<%=nbgcolor%>"> 
          <td nowrap><%=i%>.</td>
          <td nowrap><%=rs1("job_code")%></td>
          <td><%=rs1("job_status")%></td>
          <td nowrap><%=chkdate(rs1("job_date"))%></td>
          <td align="left" nowrap><%=rs1("job_tech_model")%></td>
          <td><%=rs1("job_actual_wrty_status")%></td>
          <td><%=rs1("job_code")%></td>
          <td align="center"><%=rs1("jobp_qty")%></td>
          <td><%=rs1("jobp_unitcost")%></td>
          <td><%=rs1("jobp_subtotal")%></td>
          <td><%=rs1("jobp_faultycode")%></td>
          <td><%=rs1("fr_description")%></td>
          <td><%=rs1("job_hq_category_code")%></td>
        </tr>
              <%
totaljob = totaljob + cint(rs1("jobp_qty"))
i = i + 1
rs1.movenext
wend
rs1.close
%>
 <tr valign="top" bgcolor="<%=nbgcolor%>">
                <td colspan="7" align="right" nowrap>Total</td>
                <td align="center"><%=totaljob%></td>
                <td colspan="5">&nbsp;</td>
              </tr>
            </table></td>
  </tr>
</table>

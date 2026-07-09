<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=inventory_warehouse_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->


<table border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td width="40" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>No</strong></font></td>
    <td width="198" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Item  Code.</strong></font></td>
    <td width="75" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Item  Name</strong></font></td>
    <td width="108" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><font color="#FFFFFF"><strong>Total </strong></font>Stock-In</strong></font></td>
    <td width="107" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><font color="#FFFFFF"><strong>Total</strong></font> Stock-Out</strong></font></td>
    <td width="54" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Balance</strong></font></td>
  </tr>
<%
sql1 = request.Cookies("GAPS")("sqlexcel")
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql1
rs.CursorLocation  = 3
rs.Open
while not rs.eof

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
  <tr bgcolor="<%=nbgcolor%>">
    <td height="40" align="center"><%=j%></td>
    <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("rpi_item_code")%></font></strong></td>
    <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><%=rs("rpi_item_name")%></td>
    <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_warehouse_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;stk_reference=<%=wh_code%>&amp;stk_itm_code=<%=rs("rpi_item_code")%>&amp;stype=In&amp;wh_code=<%=wh_code%>','cb18','scrollbars=yes,resizable=yes,width=500,height=500')"> <%=rs("rpi_total_in")%></a></strong></td>
    <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_warehouse_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;stk_reference=<%=wh_code%>&amp;stk_itm_code=<%=rs("rpi_item_code")%>&amp;stype=Out&amp;wh_code=<%=wh_code%>','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"> <%=rs("rpi_total_out")%></a></strong></td>
    <td align="center" nowrap="nowrap" bgcolor="#B9B9FF"><strong><%=rs("rpi_total_in") + rs("rpi_total_out")%></strong></td>
  </tr>
  <% 
rpi_total_in = rpi_total_in + rs("rpi_total_in")
rpi_total_out = rpi_total_out + rs("rpi_total_out")
rpi_balance = rpi_balance + (rs("rpi_total_in") + rs("rpi_total_out"))
count = count + 1 
i = i + 1
rs.MoveNext
wend
rs.Close
Set rs = Nothing
%>
  <tr bgcolor="#F3F3F3">
    <td height="40" colspan="3" align="right" bgcolor="#CCCCCC"><strong>Total</strong><strong></strong><strong></strong></td>
    <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_total_in%></strong></td>
    <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_total_out%></strong></td>
    <td align="center" bgcolor="#6F6FFF"><strong><%=rpi_balance%></strong></td>
  </tr>
</table>

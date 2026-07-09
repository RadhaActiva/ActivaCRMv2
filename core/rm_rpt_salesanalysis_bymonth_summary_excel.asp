<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=salesanalysis_bymonth_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Invoice No</span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Date</span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer</span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Technician<span></span></strong></font></td>
    <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Qty</strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Amt</strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GST Amt</span></strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Parts Amt</span></strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Labour Amt</span></strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Transport Amt</span></strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Payment Amt</strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CN Amt</strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Balance Amt</strong></font></td>
  </tr>
  <%
i = 1  
sql2 = request.Cookies("GAPS")("sqlexcel")  
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
while not rs.eof 
%>
  <tr>
    <td height="40" align="center"><%=i%></td>
    <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_invoice_new.asp?inv_no=<%=rs("inv_no")%>" target="_blank"><%=rs("inv_no")%></a></font></strong></td>
    <td align="left" nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
    <td align="left" nowrap="nowrap"><%=(rs("inv_status"))%></td>
    <td align="left"><%=rs("inv_cust_code") & " " & rs("inv_cust_name") %></td>
    <td align="left"><%=rs("inv_tech_code") &  " " & rs("tech_name")%></td>
    <td align="center" nowrap="nowrap"><strong> <%=rs("inv_totalqty")%></strong></td>
    <td align="right" nowrap="nowrap"><strong> <%=chknumber2(rs("inv_totalAmt"))%></strong></td>
    <td align="right"><strong><%=chknumber2(rs("inv_gstAmt"))%></strong></td>
    <td align="right"><strong><%=chknumber2(rs("inv_totalPartsAmt")*0.9433962264)%></strong></td>
    <td align="right"><strong><%=chknumber2(rs("inv_labourAmt")*0.9433962264)%></strong></td>
    <td align="right"><strong><%=chknumber2(rs("inv_transportAmt")*0.9433962264)%></strong></td>
    <td align="right"><strong><%=chknumber2(rs("inv_payment"))%></strong></td>
     <td align="right"><strong><%=chknumber2(rs("inv_cnamount"))%></strong></td>
    <td align="right"><strong><%=chknumber2(rs("inv_balance"))%></strong></td>
  </tr>
  <%
inv_totalAmt = inv_totalAmt + ccur(rs("inv_totalAmt"))
inv_gstAmt = inv_gstAmt + ccur(rs("inv_gstAmt"))
inv_totalPartsAmt = inv_totalPartsAmt + ccur(rs("inv_totalPartsAmt")*0.9433962264)
inv_labourAmt = inv_labourAmt + ccur(rs("inv_labourAmt")*0.9433962264)
inv_transportAmt = inv_transportAmt + ccur(rs("inv_transportAmt")*0.9433962264)
inv_payment = inv_payment + ccur(rs("inv_payment"))

if not isnull(rs("inv_cnamount")) then 
inv_cnamount = inv_cnamount + ccur(rs("inv_cnamount"))
end if

inv_balance = inv_balance + ccur(rs("inv_balance"))
i = i + 1
rs.MoveNext
wend
rs.Close

%>
  <tr bgcolor="#F3F3F3">
    <td height="40" colspan="7" align="right" bgcolor="#F3F3F3"><strong>Total</strong></td>
    <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=chknumber2(inv_totalAmt)%></strong></td>
    <td align="right" bgcolor="#F3F3F3"><strong><%=chknumber2(inv_gstAmt)%></strong></td>
    <td align="right" bgcolor="#F3F3F3"><strong><%=chknumber2(inv_totalPartsAmt)%></strong></td>
    <td align="right" bgcolor="#F3F3F3"><strong><%=chknumber2(inv_labourAmt)%></strong></td>
    <td align="right" bgcolor="#F3F3F3"><strong><%=chknumber2(inv_transportAmt)%></strong></td>
    <td align="right"><strong><%=chknumber2(inv_payment)%></strong></td>
    <td align="right" bgcolor="#F3F3F3"><strong><%=chknumber2(inv_cnamount)%></strong></td>
    <td align="right" bgcolor="#F3F3F3"><strong><%=chknumber2(inv_balance)%></strong></td>
  </tr>
</table>

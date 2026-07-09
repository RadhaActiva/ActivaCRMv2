<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=CnSummary_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>CN No</span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> CN Date</span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer <br />
    </span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>DO No<span><br />
    </span></strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
    <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
    <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CN Qty</strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CN Amt</strong></font></td>
    <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GST Amt</span></strong></font></td>
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
    <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("cn_no")%></font></strong></td>
    <td align="left" nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
    <td align="left" nowrap="nowrap"><%=(rs("cn_status"))%></td>
    <td align="left"><%=rs("cn_cust_code") & " " & rs("cn_cust_name") %></td>
    <td align="left"><%=rs("cn_do_no")%></td>
    <td align="left"><%=rs("inv_no")%></td>
    <td align="left"><%=chkdate(rs("inv_date"))%></td>
    <td align="center" nowrap="nowrap"><strong> <%=rs("cn_totalqty")%></strong></td>
    <td align="right" nowrap="nowrap"><strong> <%=chknumber2(ccur(rs("cn_totalAmt")*0.9433962264))%></strong></td>
    <td align="right"><strong><%=chknumber2(rs("cn_gstAmt"))%></strong></td>
  </tr>
  <%
cn_totalqty = cn_totalqty + ccur(rs("cn_totalqty"))
cn_totalAmt = cn_totalAmt + ccur(rs("cn_totalAmt")*0.9433962264)
cn_gstAmt = cn_gstAmt + ccur(rs("cn_gstAmt"))

i = i + 1
rs.MoveNext
wend
rs.Close

%>
  <tr bgcolor="#F3F3F3">
    <td height="40" colspan="8" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
    <td height="40" align="center" bgcolor="#CCCCCC"><strong><%=(cn_totalqty)%></strong></td>
    <!--Open-->
    <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(cn_totalAmt)%></strong></td>
    <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(cn_gstAmt)%></strong></td>
  </tr>
</table>

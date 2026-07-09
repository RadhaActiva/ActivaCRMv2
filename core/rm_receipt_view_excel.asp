<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=receipt_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<h2>Receipt List</h2>

<table width="100%" border="0" cellpadding="4" cellspacing="0">

  <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Tel 1</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer City</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Payment Type</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted By</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Receipt Amt</strong></font></td>
  </tr>
  <%
i = 1  
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
                      <td height="40"> <%=i%> </td>
                      <td nowrap="nowrap"><strong><font color="#0000FF"><%=rs("receipt_no")%></font></strong><br></td>
                      <td nowrap="nowrap"><%=rs("receipt_status")%></td>
                      <td nowrap="nowrap"><%=chkdate(rs("receipt_date"))%></td>
                      <td nowrap="nowrap"><%=rs("receipt_inv_no")%></td>
                      <td> <%=(rs("receipt_cust_name"))%></td>
                      <td> <%=rs("receipt_cust_tel1")%></td>
                      <td> <%=rs("receipt_cust_city")%></td>
                      <td><%=rs("receipt_paymenttype")%></td>
                      <td><%=chkdate(rs("receipt_createdby"))%></td>
                      <td><%=(rs("receipt_createdby"))%></td>
                      <td align="center"><%=chknumber2(rs("receipt_totalpayment"))%></td>
  </tr>
  <%
receipt_totalpayment = receipt_totalpayment + rs("receipt_totalpayment")
i = i + 1
rs.MoveNext
wend
rs.Close
Set rs = Nothing
%>
  <tr>
                       <td height="30" colspan="11" align="right" bgcolor="#CCCCCC"><strong> Total </strong></td>
                       <td height="30" align="center" bgcolor="#CCCCCC"><strong><%=chknumber2(receipt_totalpayment)%></strong></td>
  </tr>
</table>

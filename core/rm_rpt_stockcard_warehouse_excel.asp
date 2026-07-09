<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=warehouse_location_" & year(date()) & month(date()) & day(date()) & ".xls"

job_date_from = request("job_date_from")
%>
<!-- #include file="database/datastore.asp" -->


<h2>Report Summary By Store Location</h2>
<h3>For Period Ending <%=job_date_from%></h3>

<table border="0" cellpadding="4" cellspacing="0">
     <tr>
                      <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No.</span></strong></font></td>
                      <td height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Store</span></strong></font></td>
                      <td width="151" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
                      <td width="139" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total&nbsp;</strong></font></td>
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
   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("stk_reference")%> - <%=rs("wh_name")%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=rs("Totalstockqty")%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber(rs("totalvalue"))%>&nbsp;</strong></td>
                    </tr>
  <%
totalqty = totalqty + chknumber(rs("Totalstockqty")) 
totalvalue = totalvalue + chknumber(rs("totalvalue"))
i = i + 1
rs.MoveNext
wend
rs.Close

%>
  <tr bgcolor="#F3F3F3">
                     <td height="40" colspan="2" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=totalqty%></strong></td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(totalvalue)%>&nbsp;</strong></td>
                   </tr>
</table>

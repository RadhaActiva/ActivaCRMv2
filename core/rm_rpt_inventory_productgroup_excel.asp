<%  


Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=inventory_productcost_" & year(date()) & month(date()) & day(date()) & ".xls"
job_date_from = request("job_date_from")
%>
<!-- #include file="database/datastore.asp" -->
<h2>Report Inventory By Product Group for Period Ending <%=job_date_from%></h2>

<table width="100%" border="0" cellpadding="4" cellspacing="0">

  <tr>
    <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
    <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Category</span></strong></font></td>
    <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Item  Code.</strong></font></td>
    <td width="343" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item  Name</span></strong></font></td>
    <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
    <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Costing (Unit)</span></strong></font></td>
    <td width="139" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total Value</strong></font></td>
    <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Average Cost<span><br /></span></strong></font></td>
  </tr>
  <%
  
'sql3 = request.Cookies("GAPS")("sqlexcel2")

'rs.ActiveConnection = strconnect
'rs.Source = sql3
'rs.CursorLocation  = 3
'rs.Open
'if not rs.eof then
'   t_totalqty = rs("totalqty")
'   t_totalvalue = rs("totalvalue")
'end if
'rs.close

t_totalqty = 0
t_totalvalue = 0

sql1 = request.Cookies("GAPS")("sqlexcel")

set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql1
rs.CursorLocation  = 3
rs.Open
i=1
while not rs.eof
if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

if (IsNull(rs("md_costprice"))) Then
	total = 0
else
	total = rs("totalqty") * rs("md_costprice")
end if 

totalavg = total

%>
  <tr bgcolor="<%=nbgcolor%>">
    <td height="40" align="center"><%=i%></td>
    <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("md_category")%></font></strong></td>
    <td align="left" nowrap="nowrap"><strong><font color="#0000FF"><%=rs("stk_itm_code")%></font></strong></td>
    <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><%=rs("md_desc")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=rs("totalqty")%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2(rs("md_costprice"))%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=chknumber2(total)%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2(rs("avgcost"))%></strong></td>
  </tr>
  <%

t_totalqty = t_totalqty + cint(rs("totalqty")) 
t_totalvalue = t_totalvalue + total 

'totalqty = totalqty + chknumber(rs("totalqty")) 
'totalvalue = totalvalue + chknumber(rs("totalvalue"))

count = count + 1 
i = i + 1
rs.MoveNext
wend
rs.Close
Set rs = Nothing



If Err.Number <> 0 Then

  Response.Write (Err.Description)   

  Response.End 

End If

%>
  <tr bgcolor="#F3F3F3">
    <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
    <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=t_totalqty%></strong></td>
    <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
    <td align="right" bgcolor="#999999"><strong><%=chknumber2(t_totalvalue)%>&nbsp;</strong></td>
    <td align="right" bgcolor="#999999"></td>
  </tr>
</table>

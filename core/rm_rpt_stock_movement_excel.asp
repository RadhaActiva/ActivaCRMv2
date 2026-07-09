
<%  
searchvalue = request("searchvalue")
wh_code = request("wh_code")
whchk = request("whchk")
job_date_to = request("job_date_to")
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=Stock_Movement_" & searchvalue & year(date()) & month(date()) & day(date()) & ".xls"
    if whchk = "Yes" then
        wh_code = "All"
    end if
%>

<!-- #include file="database/datastore.asp" -->
<h2>Report Stock Movement By Product For Period Ending <%=job_date_to%></h2>

<table>
<tr><td>Item Code: </td><td><%=searchvalue%></td></tr>
<tr><td>Location: </td><td><%=wh_code%></td></tr>
</table>
<table width="100%" border="0" cellpadding="4" cellspacing="0">
     <tr>
     <td width="100" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Doc Date</span></strong></font></td>
     <td width="200" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Doc Type</span></strong></font></td>
     <td width="160" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Doc No</strong></font></td>
     <td width="50" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><span><strong>Qty</strong></span></font></td>
     <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty C/F</span></strong></font></td>
     <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Purchase Unit Cost</span></strong></td>
     <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total Purchase Cost</span></strong></td>
     <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Weighted Avg unit Cost</span></strong></td>
     <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total Cost</span></strong><br/> 
  </tr>
  <%
  
sql4 = request.Cookies("GAPS")("sqlexcel")
set rs4 = server.CreateObject("adodb.recordset")
rs4.ActiveConnection = strconnect
rs4.Source = sql4
rs4.CursorLocation  = 3
rs4.Open
while not rs4.eof
if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if


%>
          <tr bgcolor="<%=nbgcolor%>">
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chkdate(rs4("stk_date"))%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"> <%=rs4("stk_doc_type")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"> <%=rs4("stk_doc_no")%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_qty")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_bf_qty")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=chknumber(rs4("stk_purchase_cost"))%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=chknumber(rs4("stk_purchase_totalcost"))%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=chknumber(rs4("stk_avg_cost"))%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=chknumber(rs4("stk_total_cost"))%></td>
           </tr>

  <%

count = count + 1 
i = i + 1
rs4.MoveNext
wend
rs4.Close
Set rs = Nothing

If Err.Number <> 0 Then
  Response.Write (Err.Description)   
  Response.End 
End If
%>
</table>

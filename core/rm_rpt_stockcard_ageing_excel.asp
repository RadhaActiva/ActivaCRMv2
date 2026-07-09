<!-- #include file="database/datastore.asp" -->

<%
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=rm_rpt_stockcard_ageing_excel_" & year(date()) & month(date()) & day(date()) & ".xls"
stock_aging_date = request("stock_aging_date")
%>
<html>
<head>
</head>

<body>

<%
i = 1
sql2 = request.Cookies("GAPS")("stockageingsql")
'response.write sql2
'response.End()
set rs = server.CreateObject("adodb.recordset")
set rs1 = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr>
    <td colspan="15" align="left" class="style1"><font size="4"><strong>Stock Ageing Up to Date: <%=stock_aging_date%></strong></font></td>
  </tr>
  <tr>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Stock Code</span></strong></font></td>
                      <td bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock Name </strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>RCP</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>AVG Cost</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Brand</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Status</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 0 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 1 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 2 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 3 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 4 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 5 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 6 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock Bal</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total Value</strong></font></td>
                    </tr>
        <% 
count = 1		
while not rs.eof 
%>
        
         <tr bgcolor="<%=nbgcolor%>">
                     <td height="40" align="center"><%=count%> </td>
                      <td align="left" nowrap="nowrap"><a href="rm_stock_new.asp?md_code=<%=rs("ag_stock_code")%>"><strong><%=rs("ag_stock_code")%></strong></a></td>
                      <td><%=rs("md_desc")%> </td>
                      <td align="center"><%=chknumber2(rs("md_rcpprice"))%></td>
                      <td align="center" nowrap="nowrap">
					   <%if request.Cookies("GAPS")("view_cost")="Y" then %>
					  <%=chknumber2(rs("ag_averagecost"))%>
                      <% subtotal=chknumber(rs("ag_current_stock")) * chknumber2(rs("ag_averagecost")) %>
                      <%else%>
                      Restricted View
                      <%end if%>
                       </td>
                      <td align="center"><%=rs("md_brands")%></td>
                      <td align="center"><%=rs("md_status")%></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y0"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y1"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y2"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y3"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y4"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y5"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y6"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("ag_current_stock"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(subtotal)%></strong></td>
                 </tr>
         
        <%
t_Y0 =t_Y0+rs("Y0")
t_Y1 =t_Y1+rs("Y1")
t_Y2 =t_Y2+rs("Y2")
t_Y3 =t_Y3+rs("Y3")
t_Y4 =t_Y4+rs("Y4")
t_Y5 =t_Y5+rs("Y5")
t_Y6 =t_Y6+rs("Y6")
grand_total = grand_total + subtotal
ag_current_stock = ag_current_stock + rs("ag_current_stock")
count = count + 1 
rs.MoveNext
wend
rs.Close
Set rs = Nothing
%>

 <tr>
                      <td height="40" colspan="7" align="right"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y0%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y1%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y2%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y3%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y4%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y5%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y6%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=ag_current_stock%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=grand_total%></strong></td>
                    </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

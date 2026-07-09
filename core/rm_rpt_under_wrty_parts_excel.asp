<%  
job_to = request("job_to")
etype = request("etype")
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=monthly_under_warranty_expenses_" & etype & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<table border="0" width="100%" cellpadding="5" cellspacing="0">
   <tr>
     <td height="50" colspan="5" align="left" class="style1"><nowrap><h3>Monthly Under Warranty Parts Expenses For Month Ending <%=job_to%> for type <%=etype%></h3></nowrap></td>
   </tr>
   <tr>                  
                      <td align="center" width="50" bgcolor="#666666" class="style1"><font color="#FFFFFF">No</font></td>
                      <td align="left" width="560" bgcolor="#666666" class="style1"><font color="#FFFFFF"><nowrap>Spare Part Name</nowrap></font></td>
                      <td align="left" width="110" bgcolor="#666666" class="style1"><font color="#FFFFFF"><nowrap>Part Code</nowrap></font></td>
                      <td align="left" width="100" bgcolor="#666666" class="style1"><font color="#FFFFFF">Avg Cost</font></td>
                      <td align="left" width="100" bgcolor="#666666" class="style1"><font color="#FFFFFF">Qty</font></td>
                      <td align="left" width="100" bgcolor="#666666" class="style1"><font color="#FFFFFF">Total</font></td>       
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
    if not isnull(rs("avgcost")) then
        totalavgcost=rs("qty") * rs("avgcost")
    end if
    
%>
   <tr bgcolor="<%=nbgcolor%>">
               <tr bgcolor="<%=nbgcolor%>">
                      <td height="20" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><%=rs("md_desc")%></td>
                      <td align="left" nowrap="nowrap"><strong> <%=rs("jobp_partcode")%></strong></td>
                      <td align="left" nowrap="nowrap"><%=rs("avgcost")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("qty")%></td>
                      <td align="left" nowrap="nowrap"><%=totalavgcost%></td>
                    </tr>
  <%
totalqty = totalqty + rs("qty")
totalcost = totalcost + totalavgcost

i = i + 1
rs.MoveNext
totalavgcost=0
wend
rs.Close

%>
     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#CCCCCC"></td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                      <td align="right" bgcolor="#CCCCCC"></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#999999"><strong><%=(totalqty)%></strong></td>
                      <td align="left" bgcolor="#999999"><strong><%=chknumber2(totalcost)%></strong></td>
                    </tr>
</table>

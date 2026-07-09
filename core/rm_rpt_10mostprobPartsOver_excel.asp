<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=most10OverParts_" & year(date()) & month(date()) & day(date()) & ".xls"

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
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="15%" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Stock Code (Part No)</font></strong></td>
                      <td width="30%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Part Description</span></strong></font></td>
                      <td width="13%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Type</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Sales Qty (Fault)</strong></font></td>
                      <td width="9%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>% of Fault</span></strong></font></td>
                    </tr>
              <%
sql2 = request.Cookies("GAPS")("sqlexcel")
i = 1	   
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
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
                      <td height="40" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("jobp_partcode")%></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=rs("jobp_desc")%></td>
                      <td align="center" nowrap="nowrap"><%=rs("job_tech_type")%></td>
                      <td align="center" nowrap="nowrap"><strong><%=rs("totalParts")%></strong></td>
                      <td align="center"><strong>0</strong></td>
                    </tr>
              <%
totalParts = totalParts + cint(rs("totalParts"))
i = i + 1
rs.movenext
wend
rs.close
%>
   <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=totalParts%></strong></td>
                      <td align="center" bgcolor="#999999"><strong>0</strong></td>
                    </tr>
 
</table>

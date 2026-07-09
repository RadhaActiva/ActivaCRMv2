<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=most10Over_" & year(date()) & month(date()) & day(date()) & ".xls"


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
            <td width="100%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
              <tr>
                <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                <td width="15%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item Code.</span></strong></font></td>
                <td width="30%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Model Description</span></strong></font></td>
                <td width="13%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Type<br />
                </span></strong></font></td>
                <td width="13%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Sales Quanlity<br />
                </span></strong></font></td>
                <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cases of Fault</strong></font></td>
                <td width="9%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>% of Fault</span></strong></font></td>
              </tr>
              <%
i = 1


sql2 = request.Cookies("GAPS")("sqlexcel")

'response.write sql2
'response.End()
i = 1	   
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
              <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs1("job_Model")%></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=rs1("md_desc")%></td>
                      <td align="center" nowrap="nowrap"><%=rs1("job_tech_type")%></td>
                      <td align="center" nowrap="nowrap"><strong> 0</strong></td>
                      <td align="center" nowrap="nowrap"><strong> <%=rs1("totaljob")%></strong></td>
                      <td align="center"><strong>0</strong></td>
                    </tr>
              <%
totaljob = totaljob + cint(rs1("totaljob"))
i = i + 1
rs1.movenext
wend
rs1.close

grandtotal = total_over+total_under
%>
              <tr bgcolor="#F3F3F3">
                <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Total</strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong> 0</strong></td>
                <!--Open-->
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=totaljob%></strong></td>
                <td align="center" bgcolor="#999999"><strong>0</strong></td>
              </tr>
            </table></td>
          </tr>
</table>

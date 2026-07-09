<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=monthtech_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="100%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
              <tr>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1">&nbsp;</td>
                <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1">&nbsp;</td>
                <td colspan="3" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Accepted</font></strong></td>
                <td colspan="3" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Done</font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Posted</font></strong></td>
              </tr>
              <tr>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Technician</span></strong></font></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>1-3 days</span></strong></font></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>4-6 days</span></strong></font></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>7 days &amp; above</span></strong></font></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>1-3 days</span></strong></font></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>4-6 days</span></strong></font></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>7 days &amp; above</span></strong></font></td>
                <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1">&nbsp;</td>
              </tr>
              <%
i = 1


sql2 = "SELECT id, tech_code, tech_name, pending_1to3d, pending_4to6d, pending_7above, done_1to3d, done_4to6d, done_7above, posted_qty " & _
	   "FROM tblrpr_monthtech where id is not null order by tech_code "

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
                <td height="40" align="center" nowrap="nowrap"><%=i%></td>
                <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs1("tech_code")%> - <%=rs1("tech_name")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><%=rs1("pending_1to3d")%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><%=rs1("pending_4to6d")%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><%=rs1("pending_7above")%></strong></td>
                <td align="center" nowrap="nowrap"><strong><%=rs1("done_1to3d")%></strong></td>
                <td align="center" nowrap="nowrap"><strong><%=rs1("done_4to6d")%></strong></td>
                <td align="center" nowrap="nowrap"><strong><%=rs1("done_7above")%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><%=rs1("posted_qty")%></strong></td>
              </tr>
              <%
pending_1to3d = pending_1to3d + rs1("pending_1to3d")
pending_4to6d = pending_4to6d + rs1("pending_4to6d")
pending_7above = pending_7above + rs1("pending_7above")
done_1to3d = done_1to3d + rs1("done_1to3d")
done_4to6d = done_4to6d + rs1("done_4to6d")
done_7above = done_7above + rs1("done_7above") 
posted_qty = posted_qty + rs1("posted_qty") 

i = i + 1
rs1.movenext
wend
rs1.close
%>
              <tr bgcolor="<%=nbgcolor%>">
                <td height="40" colspan="2" align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Total</strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=pending_1to3d%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=pending_4to6d%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=pending_7above%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=done_1to3d%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=done_4to6d%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=done_7above%></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=posted_qty%></strong></td>
              </tr>
            </table></td>
          </tr>
</table>

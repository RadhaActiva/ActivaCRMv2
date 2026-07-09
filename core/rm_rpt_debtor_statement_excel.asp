<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=debtorstmt_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
Debtor Statement<br />
<br />
<table width="100%" border="0" cellpadding="4" cellspacing="0">
 <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="4%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="16%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Document No</span></strong></font></td>
                      <td width="11%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Doc. Date</span></strong></font></td>
                      <td width="7%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td width="11%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Cust. Code<br />
                      </span></strong></font></td>
                      <td width="16%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cust. Name</strong></font></td>
                      <td width="10%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Tech. Code</strong></font></td>
                      <td width="7%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Type<span><br />
                      </span></strong></font></td>
                      <td width="9%" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Debit</strong></font></td>
                      <td width="9%" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Credit</span></strong></font></td>
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
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF">
                      <% if rs("cn") = "INV" then %>
                      <a href="rm_invoice_new.asp?inv_no=<%=rs("cn_no")%>" target="_blank"><%=rs("cn_no")%></a>
                      <% elseif rs("cn") = "CN" then %>
                      <a href="rm_cn_new.asp?cn_no=<%=rs("cn_no")%>" target="_blank"><%=rs("cn_no")%></a>
                      <% elseif rs("cn") = "Pay" then %>
                      <a href="rm_receipt_new.asp?receipt_no=<%=rs("cn_no")%>" target="_blank"><%=rs("cn_no")%></a>
                      <% End if %>
                      
                      </font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("cn_status"))%></td>
                      <td align="left"><%=rs("cn_cust_code") %></td>
                      <td align="left"><%=rs("cn_cust_name") %></td>
                      <td align="left"><%=rs("inv_tech_code") %></td>
                      <td align="center"><%=rs("CN")%></td>
                      <td align="right" nowrap="nowrap"><strong> 
					  <% 
					  if rs("cn") = "INV" then 
					  response.write chknumber2(rs("cn_totalAmt"))
					  end if
					  %></strong></td>
                      <td align="right"><strong><%
					  if rs("cn") = "Pay" or rs("cn") = "CN" then 
					  response.write chknumber2(rs("cn_totalAmt"))
					  end if
					  %></strong></td>
                    </tr>
  <%
if rs("cn") = "INV" then    
inv_debitAmt = inv_debitAmt + ccur(rs("cn_totalAmt"))
end if

if (rs("cn") = "Pay" or rs("cn") = "CN") then 
inv_creditAmt = inv_creditAmt + ccur(rs("cn_totalAmt"))
end if

i = i + 1
rs.MoveNext
wend
rs.Close

%>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="8" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(inv_debitAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(inv_creditAmt)%></strong></td>
                    </tr>
                  </table>

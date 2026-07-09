<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=debtor_aging_detail_" & year(date()) & month(date()) & day(date()) & ".xls"

%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
    <td colspan="2" valign="top"><table width="100%" border="0" cellpadding="4" cellspacing="0">
      <tr>
        <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
        <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Document No</span></strong></font></td>
        <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Doc. Date</span></strong></font></td>
        <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
        <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer<br />
        </span></strong></font></td>
        <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Job Code</strong></font></td>
        <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Ref</strong></font></td>
        <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Type<span><br />
        </span></strong></font></td>
        <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Subtotal</strong></font></td>
        <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>GST</strong></font></td>
        <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
      </tr>
              <%
sql2 = request.Cookies("AlphaCRM")("sqlexcel")
'response.write sql2
'response.End()
i = 1	   
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
        <td align="left"><%=rs("cn_cust_code") %><br />
          <%=rs("cust_name") %><br /></td>
        <td align="left"><%=rs("cn_job_code") %></td>
        <td align="left"><%=rs("cn_inv_no") %></td>
        <td align="center"><%=rs("CN")%></td>
        <td align="right"><strong><%=chknumber2(rs("cn_subtotal"))%></strong></td>
        <td align="right"><strong><%=chknumber2(rs("cn_gstAmt"))%></strong></td>
        <td align="right" nowrap="nowrap"><strong><%=chknumber2(rs("cn_totalAmt"))%>
        </strong></td>
      </tr>
<%
inv_subtotal = inv_subtotal + ccur(rs("cn_subtotal"))
inv_gstamt = inv_gstamt + ccur(rs("cn_gstAmt"))
inv_totalamt = inv_totalamt + ccur(rs("cn_totalAmt"))
i = i + 1
rs.movenext
wend
rs.close
%>
      <tr bgcolor="#F3F3F3">
        <td colspan="5" align="left" bgcolor="#FFFFFF"></td>
        <td height="40" colspan="3" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
        <td height="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_subtotal)%></strong></td>
        <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_gstamt)%></strong></td>
        <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_totalamt)%></strong></td>
      </tr>
            </table></td>
          </tr>
</table>

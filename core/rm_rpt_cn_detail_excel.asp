<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=CnDetail_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="0" cellpadding="4" cellspacing="0">
  <tr>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>CN No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> CN Date</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer Code<br />
                        </span></strong></font><font color="#FFFFFF"><strong><span><br />
                      </span></strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Customer Name</strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>DO No</strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Code</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Description</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Cost</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong> CN<br />
                      </strong></font><font color="#FFFFFF"><strong>Qty</strong></font></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>CN<br /> 
                      Amt</span></strong></font></td>
                      <!--<td height="30" align="right" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Gst<br />
                      Amt</font></strong></td>-->
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
    <tr>
                      <td height="40" align="center" nowrap="nowrap"><%=i%>.</td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_cn_new.asp?cn_no=<%=rs("cnd_cn_no")%>" target="_blank"><%=rs("cnd_cn_no")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("cn_status"))%></td>
                      <td align="left"><%=rs("cn_cust_code")%></td>
                      <td align="left"><%=rs("cn_cust_name")%></td>
                      <td align="left"><%=rs("cn_do_no")%></td>
                      <td align="left"><%=rs("inv_no")%></td>
                      <td align="left"><%=chkdate(rs("inv_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=rs("cnd_partcode")%></td>
                      <td align="left"><%=rs("cnd_desc")%></td>
                      <td align="right"><strong><%=round(rs("md_averageecost"),2)%></strong></td>
                      <td align="right" nowrap="nowrap"><strong> <%=rs("cnd_qty")%></strong></td>
                      <td align="right"><strong><%=round((rs("cnd_subtotal")*0.9433962264),2) + round((rs("cnd_subtotal")*0.0566037735849057),2) %></strong></td>
                   <!--   <td align="right"><strong><%=round((rs("cnd_subtotal")*0.0566037735849057),2)%></strong></td>-->
                    </tr>
  <%
cnd_qty = cnd_qty + ccur(rs("cnd_qty"))
cnd_subtotal = cnd_subtotal + round((rs("cnd_subtotal")*0.9433962264),2) + (cnd_gstAmt + round((rs("cnd_subtotal")*0.0566037735849057),2))
'cnd_gstAmt = cnd_gstAmt + round((rs("cnd_subtotal")*0.0566037735849057),2)

i = i + 1
rs.MoveNext
wend
rs.Close

%>
  <tr bgcolor="#F3F3F3">
    <td height="40" colspan="12" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
    <td height="40" align="center" bgcolor="#CCCCCC"><strong><%=(cnd_qty)%></strong></td>
    <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(cnd_subtotal)%></strong></td>
   <!-- <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(cnd_gstAmt)%></strong></td>-->
  </tr>
</table>

<!-- #include file="database/datastore.asp" -->
<html>
<head>
<!-- #include file="meta.asp" -->
</head>
<body>
<%
cn_cust_code = request("cn_cust_code")
i = 1
sql2 = "SELECT     v_debtor_balance.cn_id, v_debtor_balance.cn_no, v_debtor_balance.cn_date, v_debtor_balance.cn_cust_code, v_debtor_balance.cn_job_code, v_debtor_balance.cn_inv_no, " & _
        "v_debtor_balance.cn_totalAmt, v_debtor_balance.cn_status, v_debtor_balance.cn_gstAmt, v_debtor_balance.cn_subtotal, v_debtor_balance.CN, v_debtor_balance.cn_rouding,  " & _
        "v_debtor_balance.cn_createdby, tblcustomer.cust_name  " & _
        "FROM v_debtor_balance INNER JOIN  " & _
        "tblcustomer ON v_debtor_balance.cn_cust_code = tblcustomer.cust_code " & _
	    "where v_debtor_balance.cn_status = 'Posted' and v_debtor_balance.cn_cust_code like '%" & cn_cust_code & "%' and v_debtor_balance.cn_status = 'Posted' " & _
        "order by v_debtor_balance.cn_no"	
response.Cookies("AlphaCRM")("sqlexcel") = sql2
''response.write sql2
'response.End()
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if rs.eof then
   norecord = "There is no record found."
end if

If Not rs.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs.RecordCount
Remain = TotalRecord - Showed

If Remain > row Then
  LoopMax = Showed + row
Else
  LoopMax = Showed + Remain
End If

	If Int(TotalRecord/row) <> TotalRecord/row Then
	  pgCount = Int(TotalRecord/row) + 1
	Else
	  pgCount = TotalRecord/row
	End If

	if LoopMax mod row = 0 then
		pagestartno = LoopMax/row
	else
		pagestartno = pgCount
	end if		
end if

count = count + Showed

link = "&cn_cust_code=" & cn_cust_code 
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>List of Invoice, Receipts and CN for &quot;<%=cn_cust_code%>&quot;</strong></font></td>
    <td align="right" class="style21"><span class="titlegrey1"><a href="rm_rpt_debtor_aging_detail_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></span></td>
  </tr>
  <tr>
    <td colspan="2" align="left" valign="top"> **** Only Posted Status</td>
  </tr>
  <tr> 
    <td colspan="2" align="right" valign="top"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
    <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_debtor_aging_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_debtor_aging_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
  </tr>
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

if not rs.eof then
rs.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
      <tr bgcolor="<%=nbgcolor%>">
        <td height="40" align="center"><%=j%></td>
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
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
      <tr bgcolor="#F3F3F3">
        <td colspan="5" align="left" bgcolor="#FFFFFF"></td>
        <td height="40" colspan="3" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
        <td height="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_subtotal)%></strong></td>
        <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_gstamt)%></strong></td>
        <!--Open-->
        <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_totalamt)%></strong></td>
      </tr>
    </table></td>
  </tr>
  <tr valign="top"> 
    <td colspan="10" align="right"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
    <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_debtor_aging_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_debtor_aging_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
  </tr>
</table>
</body>
</html>

<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
orderby = request("orderby")
ordertype = request("ordertype")
cn_totalAmt = request("cn_totalAmt") 
balanceAmtCompare = request("balanceAmtCompare") 

if request("cn_totalAmt") = "" then 
   cn_totalAmt = "0"
end if

if request("balanceAmtCompare") = "" then 
   balanceAmtCompare = "notequal"
end if

''''''''''Total Service''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql2 = "SELECT   SUM(v_debtor_balance.cn_totalAmt) AS cn_totalAmt " & _
		"FROM         v_debtor_balance INNER JOIN " & _
		"tblcustomer ON v_debtor_balance.cn_cust_code = tblcustomer.cust_code " & _
		"GROUP BY v_debtor_balance.cn_cust_code, v_debtor_balance.cn_status, tblcustomer.cust_name, tblcustomer.cust_type, tblcustomer.cust_status, tblcustomer.cust_reg_no,  " & _
		"tblcustomer.cust_company, tblcustomer.cust_address, tblcustomer.cust_postcode, tblcustomer.cust_state, tblcustomer.cust_city, tblcustomer.cust_email, tblcustomer.cust_tel1,  " & _
		"tblcustomer.cust_pic " & _
		"HAVING  v_debtor_balance.cn_cust_code is not null and (v_debtor_balance.cn_status = 'Posted') "

if balanceAmtCompare = "notequal" then 
   sql2 = sql2 & " and (SUM(v_debtor_balance.cn_totalAmt) <> " & cn_totalAmt & ") "
elseif balanceAmtCompare = "equal" then 
   sql2 = sql2 & " and (SUM(v_debtor_balance.cn_totalAmt) = " & cn_totalAmt & ") "
elseif balanceAmtCompare = "greater" then 
   sql2 = sql2 & " and (SUM(v_debtor_balance.cn_totalAmt) > " & cn_totalAmt & ") "
elseif balanceAmtCompare = "lesses" then 
   sql2 = sql2 & " and (SUM(v_debtor_balance.cn_totalAmt) < " & cn_totalAmt & ") "
end if

if searchitem <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue & "%' "
end if

grandtotal = selectid(sql2)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
i = 1
sql2 = "SELECT     v_debtor_balance.cn_cust_code, SUM(v_debtor_balance.cn_totalAmt) AS cn_totalAmt, tblcustomer.cust_name, tblcustomer.cust_type, tblcustomer.cust_status, tblcustomer.cust_reg_no, " & _
		"tblcustomer.cust_company, tblcustomer.cust_address, tblcustomer.cust_postcode, tblcustomer.cust_state, tblcustomer.cust_city, tblcustomer.cust_email, tblcustomer.cust_tel1,  " & _
		"tblcustomer.cust_pic " & _
		"FROM         v_debtor_balance INNER JOIN " & _
		"tblcustomer ON v_debtor_balance.cn_cust_code = tblcustomer.cust_code " & _
		"GROUP BY v_debtor_balance.cn_cust_code, v_debtor_balance.cn_status, tblcustomer.cust_name, tblcustomer.cust_type, tblcustomer.cust_status, tblcustomer.cust_reg_no,  " & _
		"tblcustomer.cust_company, tblcustomer.cust_address, tblcustomer.cust_postcode, tblcustomer.cust_state, tblcustomer.cust_city, tblcustomer.cust_email, tblcustomer.cust_tel1,  " & _
		"tblcustomer.cust_pic " & _
		"HAVING  v_debtor_balance.cn_cust_code is not null and (v_debtor_balance.cn_status = 'Posted') "

if balanceAmtCompare = "notequal" then 
   sql2 = sql2 & " and (round(SUM(v_debtor_balance.cn_totalAmt),0) <> " & cn_totalAmt & ") "
elseif balanceAmtCompare = "equal" then 
   sql2 = sql2 & " and (round(SUM(v_debtor_balance.cn_totalAmt),0) = " & cn_totalAmt & ") "
elseif balanceAmtCompare = "greater" then 
   sql2 = sql2 & " and (round(SUM(v_debtor_balance.cn_totalAmt),0) > " & cn_totalAmt & ") "
elseif balanceAmtCompare = "lesser" then 
   sql2 = sql2 & " and (round(SUM(v_debtor_balance.cn_totalAmt),0) < " & cn_totalAmt & ") "
end if

if searchitem <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue & "%' "
end if
		
if orderby <> "" then
	sql2 = sql2 & " order by " & orderby & " " & ordertype
else
	sql2 = sql2 & " order by cn_totalAmt DESC"
end if

'response.write sql2
response.Cookies("AlphaCRM")("sqlexcel") = sql2
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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&orderby=" & orderby & "&ordertype=" & ordertype & "&cn_totalAmt=" & cn_totalAmt & "&balanceAmtCompare=" & balanceAmtCompare 
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Debtor Aging Report</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1">&nbsp;</td>
                      <td width="20%" align="right" class="titlegrey1"><a href="rm_rpt_debtor_aging_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="rm_rpt_aging_statement.asp?type=searchdata">
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="15%" class="titlegrey1"><select name="searchitem" id="searchitem">
                          <option value="v_debtor_balance.cn_cust_code"  <% if searchitem = "v_debtor_balance.cn_cust_code" then response.write " selected" %>>Cust. Code</option>
                          <option value="tblcustomer.cust_name" <% if searchitem = "tblcustomer.cust_name" then response.write " selected" %>>Cust. Name</option>
                          <option value="tblcustomer.cust_tel1" <% if searchitem = "tblcustomer.cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tblcustomer.cust_email" <% if searchitem = "tblcustomer.cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tblcustomer.cust_branch" <% if searchitem = "tblcustomer.cust_branch" then response.write " selected" %>>Branch</option>
                          <option value="tblcustomer.cust_status" <% if searchitem = "tblcustomer.cust_status" then response.write " selected" %>>Customer Status</option>
                        </select></td>
                        <td><input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                          <option value="cn_totalAmt"  <% if orderby = "cn_totalAmt" then response.write " selected" %>>Balance</option>                          
                          <option value="v_debtor_balance.cn_cust_code"  <% if orderby = "v_debtor_balance.cn_cust_code" then response.write " selected" %>>Cust. Code</option>
                          <option value="tblcustomer.cust_name" <% if orderby = "tblcustomer.cust_name" then response.write " selected" %>>Cust. Name</option>
                          <option value="tblcustomer.cust_tel1" <% if orderby = "tblcustomer.cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tblcustomer.cust_email" <% if orderby = "tblcustomer.cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tblcustomer.cust_branch" <% if orderby = "tblcustomer.cust_branch" then response.write " selected" %>>Branch</option>
                          <option value="tblcustomer.cust_status" <% if orderby = "tblcustomer.cust_status" then response.write " selected" %>>Customer Status</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>                           
                        </select>  
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td align="left" bgcolor="#FFFFFF">Balance Amount 
                    <select name="balanceAmtCompare" id="balanceAmtCompare">
                      <option value="notequal" <% if balanceAmtCompare = "notequal" then response.write " selected" %>>Not Equal</option>                      
                      <option value="greater" <% if balanceAmtCompare = "greater" then response.write " selected" %>>Greater than</option>
                      <option value="lesser" <% if balanceAmtCompare = "lesser" then response.write " selected" %>>Less than</option>
                      <option value="equal" <% if balanceAmtCompare = "equal" then response.write " selected" %>>Equal</option>
                  </select>                    <input name="cn_totalAmt" type="text" id="cn_totalAmt" value="<%=cn_totalAmt%>" size="20" maxlength="20" /></td>
                </tr>
                </form>
                
                <tr>
                  <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_aging_statement.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_aging_statement.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cust. Code</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Cust. Name<br />
                      </span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cust. Address</strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cust. Tel 1</strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cust. Email</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Balance</strong></font></td>
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
                      <td align="left" nowrap="nowrap"><%=rs("cn_cust_code")%></td>
                      <td align="left"><%=rs("cust_name") %></td>
                      <td align="left"><%=rs("cust_address") %></td>
                      <td align="left"><%=rs("cust_tel1") %></td>
                      <td align="left"><%=rs("cust_email") %></td>
                      <td align="center"><%=rs("cust_status") %></td>
                      <td align="right" nowrap="nowrap"><strong>
					  <a href="javascript:popup('rm_rpt_debtor_aging_detail.asp?cn_cust_code=<%=rs("cn_cust_code")%>','cb18','scrollbars=yes,resizable=yes,width=900,height=500')">
					  <%=chknumber2(rs("cn_totalAmt"))%></a></strong></td>
                    </tr>
<%
cn_totalAmt = cn_totalAmt + ccur(rs("cn_totalAmt"))
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
                    
                   
                    <tr bgcolor="#F3F3F3">
                      <td colspan="3" align="left" bgcolor="#FFFFFF"><br />
                      <br /></td>
                      <td height="40" colspan="4" align="right" bgcolor="#CCCCCC"><strong>SubTotal</strong></td>
                      <!--Open-->
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(cn_totalAmt)%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                       <td colspan="3" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(grandtotal)%></strong></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_aging_statement.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_aging_statement.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
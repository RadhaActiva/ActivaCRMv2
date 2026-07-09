<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")

if ordertype = "" then 
   ordertype = "desc"
end if

if request("receipt_createddate_from") <> "" then
   receipt_createddate_from = request("receipt_createddate_from")
else
   receipt_createddate_from = chkdate(DateAdd("d",-90,date()))
end if

if request("receipt_createddate_to") <> "" then
   receipt_createddate_to = request("receipt_createddate_to")
else
   receipt_createddate_to = chkdate(date())
end if

if request("receipt_status") <> "" then
   receipt_status = request("receipt_status")
else
   receipt_status = "Posted"
end if


'sql = "SELECT sum(receipt_totalpayment) as receipt_grandtotal," &_
'		"tblinvoice.inv_no, tbltechnician.tech_name " & _
'		"FROM tblreceipt " & _ 
'		"left join tblinvoice on tblinvoice.inv_no = tblreceipt.receipt_inv_no " & _ 
'		"left join tbltechnician on tbltechnician.tech_code = tblinvoice.inv_tech_code " & _
'		"where receipt_id is not null " & _
'		"and  tblreceipt.receipt_date >= '" & ChkDateYYYYMMDD(receipt_createddate_from) & "' and tblreceipt.receipt_date <= '" & ChkDateYYYYMMDD(receipt_createddate_to) & "' " 

'if searchvalue <> "" then 
'   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
'end if

'if receipt_status <> "All" and receipt_status <> "" then
'   sql = sql & " and  tblreceipt.receipt_status = '" & receipt_status & "' "
'end if

'sql = sql & "group by tblinvoice.inv_no, tbltechnician.tech_name"
		
sql = "SELECT sum(receipt_totalpayment) as receipt_grandtotal " & _		
		"FROM tblreceipt " & _
		"left join tblinvoice on tblinvoice.inv_no = tblreceipt.receipt_inv_no " & _ 
		"left join tbltechnician on tbltechnician.tech_code = tblinvoice.inv_tech_code " & _
		"where tblreceipt.receipt_id is not null " & _
		"and  tblreceipt.receipt_date >= '" & ChkDateYYYYMMDD(receipt_createddate_from) & "' and tblreceipt.receipt_date <= '" & ChkDateYYYYMMDD(receipt_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if receipt_status <> "All" and receipt_status <> "" then
   sql = sql & " and  tblreceipt.receipt_status = '" & receipt_status & "' "
end if


set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if not rs.eof then 
   receipt_grandtotal = rs("receipt_grandtotal")
end if
rs.close


i = 1
'sql = "SELECT receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
'		"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
'		"receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
'		"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
'		"FROM tblreceipt where receipt_id is not null " & _
'		"and  tblreceipt.receipt_date >= '" & ChkDateYYYYMMDD(receipt_createddate_from) & "' and tblreceipt.receipt_date <= '" & ChkDateYYYYMMDD(receipt_createddate_to) & "' "
'inv_tech_code
sql = "SELECT tblreceipt.receipt_id, tblreceipt.receipt_no, tblreceipt.receipt_status, tblreceipt.receipt_date, tblreceipt.receipt_inv_no, tblreceipt.receipt_inv_date, tblreceipt.receipt_cust_code, tblreceipt.receipt_cust_name, " & _
		"tblreceipt.receipt_cust_address, tblreceipt.receipt_cust_postcode, tblreceipt.receipt_cust_state, tblreceipt.receipt_cust_state_id, tblreceipt.receipt_cust_city, tblreceipt.receipt_cust_city_id,  " & _
		"tblreceipt.receipt_cust_email, tblreceipt.receipt_cust_tel1, tblreceipt.receipt_cust_tel2, tblreceipt.receipt_createddate, tblreceipt.receipt_createdby, tblreceipt.receipt_job_code, tblreceipt.receipt_remark, " & _ 
		"tblreceipt.receipt_paymenttype, tblreceipt.receipt_totalpayment, tblreceipt.receipt_emailsent, tblreceipt.receipt_emailsentdate, tblreceipt.receipt_cancelleddate, tblreceipt.receipt_cancelledby, " & _
		"tblinvoice.inv_no, tbltechnician.tech_name " & _
		"FROM tblreceipt " & _
		"left join tblinvoice on tblinvoice.inv_no = tblreceipt.receipt_inv_no " & _ 
		"left join tbltechnician on tbltechnician.tech_code = tblinvoice.inv_tech_code " & _
		"where tblreceipt.receipt_id is not null " & _
		"and  tblreceipt.receipt_date >= '" & ChkDateYYYYMMDD(receipt_createddate_from) & "' and tblreceipt.receipt_date <= '" & ChkDateYYYYMMDD(receipt_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if receipt_status <> "All" and receipt_status <> "" then
   sql = sql & " and  tblreceipt.receipt_status = '" & receipt_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tblreceipt.receipt_id desc"
end if

'response.write("1")

response.Cookies("GAPS")("sqlexcel") = sql

'response.write request.Cookies("GAPS")("slevel") & "<br>"
'response.write sql
'response.End()
rs.ActiveConnection = strconnect
rs.Source = sql
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
slink = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&receipt_createddate_from=" & receipt_createddate_from & "&receipt_createddate_to=" & receipt_createddate_to & "&ordertype=" & ordertype 
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby")  & "&receipt_status=" & receipt_status & "&receipt_createddate_from=" & receipt_createddate_from & "&receipt_createddate_to=" & receipt_createddate_to & "&ordertype=" & ordertype 
%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Receipt</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><span class="titlegrey1"><a href="rm_receipt_view_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></span></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_receipt_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> Receipt Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"><strong><font color="#000000"><strong>
                          <input name="receipt_createddate_from" type="text" id="receipt_createddate_from" value="<%=receipt_createddate_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.receipt_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="receipt_createddate_to" type="text" id="receipt_createddate_to" value="<%=receipt_createddate_to%>" size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.receipt_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td><select name="searchitem" id="searchitem">
                          <option value="tblreceipt.receipt_no"  <% if searchitem = "tblreceipt.receipt_no" then response.write " selected" %>>Receipt No</option>
                          <option value="tblreceipt.receipt_cust_name" <% if searchitem = "tblreceipt.receipt_cust_name" then response.write " selected" %>>Customer Name</option>
                          <option value="tblreceipt.receipt_cust_tel1" <% if searchitem = "tblreceipt.receipt_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tblreceipt.receipt_cust_email" <% if searchitem = "tblreceipt.receipt_cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tblreceipt.receipt_cust_city" <% if searchitem = "tblreceipt.receipt_cust_city" then response.write " selected" %>>Customer City</option>
                          <option value="tblreceipt.receipt_inv_no" <% if searchitem = "tblreceipt.receipt_inv_no" then response.write " selected" %>>Invoice No</option>
						  <option value="tbltechnician.tech_name" <% if searchitem = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                        </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                          <option value="tblreceipt.receipt_no"  <% if orderby = "tblreceipt.receipt_no" then response.write " selected" %>>Receipt No</option>
                          <option value="tblreceipt.receipt_cust_name" <% if orderby = "tblreceipt.receipt_cust_name" then response.write " selected" %>>Customer Name</option>
                          <option value="tblreceipt.receipt_cust_tel1" <% if orderby = "tblreceipt.receipt_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tblreceipt.receipt_cust_email" <% if orderby = "tblreceipt.receipt_cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tblreceipt.receipt_cust_city" <% if orderby = "tblreceipt.receipt_cust_city" then response.write " selected" %>>Customer City</option>
                           <option value="tbltechnician.tech_name" <% if orderby = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="Submit43" value="Display" />
                          <input name="receipt_status" type="hidden" id="receipt_status" value="<%=receipt_status%>" /></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_receipt_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_receipt_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      <td nowrap="nowrap"><div align="center" class="titlegrey1"> Status :</div> </td>
                      <td <%if receipt_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_receipt_view.asp?receipt_status=All<%=slink%>"><font color="#FFFFFF">All</font></a></strong></td>
                      <td <%if receipt_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_receipt_view.asp?receipt_status=Posted<%=slink%>"><font color="#FFFFFF">Posted</font></a></strong></td>
                      <td <%if receipt_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_receipt_view.asp?receipt_status=Cancel<%=slink%>"><font color="#FFFFFF">Cancel</font></a></strong></td>
                    </tr>
                  </table>
                    <div align="right"></div>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Tel 1</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer City</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Payment Type</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted By</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Receipt Amt</strong></font></td>
                    </tr>
                    
<% 
receipt_totalqty = 0
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
                      <td height="40"> <%=j%> </td>
                      <td nowrap="nowrap"><strong><font color="#0000FF"><a href="rm_receipt_new.asp?receipt_no=<%=rs("receipt_no")%>&receipt_inv_no=<%=rs("receipt_inv_no") %>"><%=rs("receipt_no")%></a></font></strong><br><%=rs("receipt_status")%></td>
                      <td nowrap="nowrap"><%=chkdate(rs("receipt_date"))%></td>
                      <td nowrap="nowrap"><%=rs("receipt_inv_no")%></td>
                      <td> <%=(rs("receipt_cust_name"))%></td>
                      <td> <%=rs("receipt_cust_tel1")%></td>
                      <td> <%=rs("receipt_cust_city")%></td>
                      <td><%=rs("receipt_paymenttype")%></td>
                      <td><%=chkdate(rs("receipt_createdby"))%></td>
                      <td><%=(rs("receipt_createdby"))%></td>
                      <td align="center"><%=chknumber2(rs("receipt_totalpayment"))%></td>
                    </tr>
                    
                    
<%
receipt_totalpayment = receipt_totalpayment + rs("receipt_totalpayment")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
                     <tr>
                       <td height="30" colspan="10" align="right" bgcolor="#CCCCCC"><strong> Total </strong></td>
                       <td height="30" align="center" bgcolor="#CCCCCC"><strong><%=chknumber2(receipt_totalpayment)%></strong></td>
                     </tr>
                    <tr>
                      <td height="30" colspan="10" align="right" bgcolor="#CCCCCC"><strong> Grand Total </strong></td>
                      <td height="30" align="center" bgcolor="#CCCCCC"><strong><%=chknumber2(receipt_grandtotal)%></strong></td>
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
					Response.Write " <a href='rm_receipt_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_receipt_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
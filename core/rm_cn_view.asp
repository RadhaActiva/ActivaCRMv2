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

if request("cn_createddate_from") <> "" then
   cn_createddate_from = request("cn_createddate_from")
else
   cn_createddate_from = chkdate(DateAdd("d",-90,date()))
end if

if request("cn_createddate_to") <> "" then
   cn_createddate_to = request("cn_createddate_to")
else
   cn_createddate_to = chkdate(date())
end if

if request("cn_status") <> "" then
   cn_status = request("cn_status")
else
   cn_status = "Open"
end if


sql = "SELECT sum(cn_totalqty) as total_qty, sum(cn_totalAmt) as total_amt FROM tblcn where cn_id is not null " & _
		"and  tblcn.cn_date >= '" & ChkDateYYYYMMDD(cn_createddate_from) & "' and tblcn.cn_date <= '" & ChkDateYYYYMMDD(cn_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if cn_status <> "All" and cn_status <> "" then
   sql = sql & " and  tblcn.cn_status = '" & cn_status & "' "
end if
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if not rs.eof then 
   total_qty = rs("total_qty")
   total_amt = rs("total_amt")
end if 
rs.close


i = 1
sql = "SELECT tblcn.cn_id, tblcn.cn_no, tblcn.cn_status, tblcn.cn_date, tblcn.cn_inv_no, tblcn.cn_inv_date, tblcn.cn_cust_code, tblcn.cn_cust_name, tblcn.cn_cust_address, tblcn.cn_cust_postcode, " & _
	  "tblcn.cn_cust_state, tblcn.cn_cust_state_id, tblcn.cn_cust_city, tblcn.cn_cust_city_id, tblcn.cn_cust_email, tblcn.cn_cust_tel1, tblcn.cn_cust_tel2, tblcn.cn_createddate, tblcn.cn_createdby,  " & _
	  "tblcn.cn_job_code, tblcn.cn_do_no, tblcn.cn_invoice_no, tblcn.cn_totalqty, tblcn.cn_totalPartsAmt, tblcn.cn_remark, tblcn.cn_labourAmt, tblcn.cn_transportAmt, tblcn.cn_gstAmt, tblcn.cn_totalAmt,  " & _
	  "tblcn.cn_emailsent, tblcn.cn_emailsentdate, tblcn.cn_returnedby, tblcn.cn_returneddate, tblcn.cn_submittedby, tblcn.cn_submitteddate, tblcn.cn_doneby, tblcn.cn_donedate, tblcn.cn_postedby,  " & _
	  "tblcn.cn_posteddate, tblcn.cn_cancelledby, tblcn.cn_cancelleddate,tblinvoice.inv_no, tblinvoice.inv_date " & _
	  "FROM tblcn left join tblinvoice on tblcn.cn_inv_no=tblinvoice.inv_no where tblcn.cn_id is not null " & _
		"and  tblcn.cn_date >= '" & ChkDateYYYYMMDD(cn_createddate_from) & "' and tblcn.cn_date <= '" & ChkDateYYYYMMDD(cn_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if cn_status <> "All" and cn_status <> "" then
   sql = sql & " and  tblcn.cn_status = '" & cn_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tblcn.cn_id desc"
end if

'response.write request.Cookies("GAPS")("slevel") & "<br>"
'response.write sql
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

link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&Searchor_date=" & request("Searchor_date") & "&orderby=" & request("orderby") & "&ordertype=" & request("ordertype") & "&cn_createddate_from=" & cn_createddate_from & "&cn_createddate_to=" & cn_createddate_to & "&cn_status=" & cn_status  

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>CN</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_cn_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> CN Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"><strong><font color="#000000"><strong>
                          <input name="cn_createddate_from" type="text" id="cn_createddate_from" value="<%=cn_createddate_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.cn_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="cn_createddate_to" type="text" id="cn_createddate_to" value="<%=cn_createddate_to%>" size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.cn_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td><select name="searchitem" id="searchitem">
                          <option value="tblcn.cn_no"  <% if searchitem = "tblcn.cn_no" then response.write " selected" %>>CN No</option>
                          <option value="tblcn.cn_cust_name" <% if searchitem = "tblcn.cn_cust_name" then response.write " selected" %>>Customer Name</option>
                          <option value="tblcn.cn_cust_tel1" <% if searchitem = "tblcn.cn_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tblcn.cn_cust_email" <% if searchitem = "tblcn.cn_cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tblcn.cn_cust_city" <% if searchitem = "tblcn.cn_cust_city" then response.write " selected" %>>Customer City</option>
                        </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                          <option value="tblcn.cn_no"  <% if orderby = "tblcn.cn_no" then response.write " selected" %>>CN No</option>
                          <option value="tblcn.cn_cust_name" <% if orderby = "tblcn.cn_cust_name" then response.write " selected" %>>Customer Name</option>
                          <option value="tblcn.cn_cust_tel1" <% if orderby = "tblcn.cn_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tblcn.cn_cust_email" <% if orderby = "tblcn.cn_cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tblcn.cn_cust_city" <% if orderby = "tblcn.cn_cust_city" then response.write " selected" %>>Customer City</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="Submit43" value="Display" />
                          <input name="cn_status" type="hidden" id="cn_status" value="<%=cn_status%>" /></td>
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
					Response.Write " <a href='rm_cn_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cn_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      <td nowrap="nowrap"><div align="center" class="titlegrey1"> Status :</div> </td>
                      <td <%if cn_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_cn_view.asp?cn_status=All"><font color="#FFFFFF">All</font></a></strong></td>
                      <td <%if cn_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_cn_view.asp?cn_status=Open"><font color="#FFFFFF">Open</font></a></strong></td>
                      <td <%if cn_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_cn_view.asp?cn_status=Submitted"><font color="#FFFFFF">Submitted</font></a></strong></td>
                      <td <%if cn_status="Done" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_cn_view.asp?cn_status=Done"><font color="#FFFFFF">Done</font></a></strong></td>
                      <td <%if cn_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_cn_view.asp?cn_status=Posted"><font color="#FFFFFF">Posted</font></a></strong></td>
                      <td <%if cn_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_cn_view.asp?cn_status=Cancel"><font color="#FFFFFF">Cancel</font></a></strong></td>
                    </tr>
                  </table>
                    <div align="right"></div>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>CN  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>CN  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Customer</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer City</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Submitted<span> Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Done Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted Date</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>CN Qty</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>CN Amt</strong></font></td>
                    </tr>
                    
<% 
cn_totalqty = 0
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
                      <td nowrap="nowrap"><strong><font color="#0000FF"><a href="rm_cn_new.asp?cn_no=<%=rs("cn_no")%>"><%=rs("cn_no")%></a></font></strong><br><%=rs("cn_status")%></td>
                      <td nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
                      <td nowrap="nowrap"><%=rs("inv_no")%></td>
                      <td><%=chkdate(rs("inv_date"))%></td>
                      <td><%=(rs("cn_cust_name"))%></td>
                      <td> <%=rs("cn_cust_city")%></td>
                      <td> <%=chkdate(rs("cn_returneddate"))%></td>
                      <td> <%=chkdate(rs("cn_donedate"))%></td>
                      <td> <%=chkdate(rs("cn_posteddate"))%></td>
                      <td align="center"><%=rs("cn_totalqty")%></td>
                      <td align="center"> <%=rs("cn_totalAmt")%></td>
                    </tr>
                   
                    
<%
cn_totalqty = cn_totalqty + rs("cn_totalqty")
cn_totalAmt = cn_totalAmt + rs("cn_totalAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
                    
                      <tr>
                       <td height="30" colspan="10" align="right" bgcolor="#CCCCCC"><strong> Total </strong></td>
                       <td height="30" align="center" bgcolor="#CCCCCC"><strong> <%=cn_totalqty%></strong></td>
                       <td height="30" align="center" bgcolor="#CCCCCC"><strong> <%=chknumber2(cn_totalAmt)%></strong></td>
                    </tr>
                    <tr>
                      <td height="30" colspan="10" align="right" bgcolor="#999999"><strong> Grand Total </strong></td>
                      <td height="30" align="center" bgcolor="#999999"><strong> <%=total_qty%></strong></td>
                      <td height="30" align="center" bgcolor="#999999"><strong> <%=chknumber2(total_amt)%> </strong></td>
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
					Response.Write " <a href='rm_cn_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cn_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
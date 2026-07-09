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

if request("sp_date_from") <> "" then
   sp_date_from = request("sp_date_from")
else
   sp_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("sp_date_to") <> "" then
   sp_date_to = request("sp_date_to")
else
   sp_date_to = chkdate(date())
end if

if request("sp_status") <> "" then
   sp_status = request("sp_status")
else

  if request.Cookies("GAPS")("slevel") = "technician" then 
   sp_status = "Open"
   else
   sp_status = "Submitted"
   end if
end if


sql = "SELECT sum(sp_totalqty) as totalqty, sum(sp_totalAmt) as totalAmt FROM tblsparepartrequest WHERE sp_id is not null "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if Searchor_date = "Y" then
   sql = sql & " and  sp_date >= '" & sp_date_from & "' and sp_date <= '" & sp_date_to & "' "
end if

if request.Cookies("GAPS")("slevel") = "technician" then 
   sql = sql & " and sp_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
end if

if sp_status <> "All" and sp_status <> "" then
   sql = sql & " and sp_status = '" & sp_status & "' "
end if
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if not rs.eof then 
   totalqty = rs("totalqty")
   totalAmt = rs("totalAmt")
end if
rs.close

i = 1
sql = "SELECT sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
		"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, sp_deliverydate, sp_deliveryby, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
		"sp_rejecteddate, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
		"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate " & _
		"FROM tblsparepartrequest WHERE sp_id is not null "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if Searchor_date = "Y" then
   sql = sql & " and  sp_date >= '" & sp_date_from & "' and sp_date <= '" & sp_date_to & "' "
end if

if request.Cookies("GAPS")("slevel") = "technician" then 
   sql = sql & " and sp_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
end if

if sp_status <> "All" and sp_status <> "" then
   sql = sql & " and sp_status = '" & sp_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by sp_id desc"
end if
'response.write request.Cookies("GAPS")("slevel") & "<br>"

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
'link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby")
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&sortby=" & sortby & "&job_date_from=" & sp_date_from & "&job_date_to=" & sp_date_to & "&ordertype=" & ordertype

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Spare Part Request</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_consignment_view.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> Request Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"> <strong><font color="#000000"><strong>
                          <input name="sp_date_from" type="text" id="sp_date_from" value="<%=sp_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.sp_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="sp_date_to" type="text" id="sp_date_to" value="<%=sp_date_to%>"
                                            size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.sp_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td>
                          <select name="searchitem" id="searchitem">
                            <option value="sp_no"  <% if searchitem = "sp_no" then response.write " selected" %>>Request No</option>
                            <option value="sp_tech_code" <% if searchitem = "sp_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="sp_tech_name" <% if searchitem = "sp_tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="sp_tech_tel1" <% if searchitem = "sp_tech_tel1" then response.write " selected" %>>Technician Tel 1</option>
                            <option value="sp_tech_email" <% if searchitem = "sp_tech_email" then response.write " selected" %>>Technician Email</option>
                          </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                            <option value="sp_no"  <% if orderby = "sp_no" then response.write " selected" %>>Request No</option>
                            <option value="sp_tech_code" <% if orderby = "sp_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="sp_tech_name" <% if orderby = "sp_tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="sp_tech_tel1" <% if orderby = "sp_tech_tel1" then response.write " selected" %>>Technician Tel 1</option>
                            <option value="sp_tech_email" <% if orderby = "sp_tech_email" then response.write " selected" %>>Technician Email</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="Submit43" value="Display" />
                          <input name="sp_status" type="hidden" id="sp_status" value="<%=sp_status%>" /></td>
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
					Response.Write " <a href='rm_consignment_view.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_consignment_view.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      <td nowrap="nowrap"> <strong>Status : </strong></td>
                      
                     
                      <td <%if sp_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=All"><font color="#FFFFFF"> <strong>All</strong></font></a></td>
                      <td <%if sp_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=Open"><font color="#FFFFFF"><strong>Open</strong></font></a></td>
                      <td <%if sp_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=Submitted"><font color="#FFFFFF"> <strong>Submitted</strong></font></a>
                      </td>
                      <td <%if sp_status="Approved" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=Approved"><font color="#FFFFFF"><strong>Approved</strong></font></a></td>
                      <td <%if sp_status="Delivered" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=Delivered"><font color="#FFFFFF"> <strong>Delivered</strong></font></a></td>
                      <td <%if sp_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=Posted"><font color="#FFFFFF"> <strong>Posted</strong></font></a></td>
                      <td <%if sp_status="Rejected" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=Rejected"><font color="#FFFFFF"> <strong>Rejected</strong></font></a></td>
                      <td <%if sp_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_consignment_view.asp?sp_status=Cancel"><font color="#FFFFFF"> <strong>Cancel</strong></font></a></td>
                    </tr>
                  </table>
                   </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Request  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Request  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Technician Code<br />
                      </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Mobile</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Location</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Submitted Date</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Approve Date</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Confirmed Received Date</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Qty</strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Amount (RM)</strong></font></td>
                    </tr>
                    
<% 
job_totalAmt = 0
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
                    <tr bgcolor="#FFFFFF">
                      <td height="40"><%=j%></td>
                      <td align="left" nowrap="nowrap"><strong><font color="#0000FF"><a href="rm_consignment_new.asp?sp_no=<%=rs("sp_no")%>"><%=rs("sp_no")%></a></font></strong><br><%=rs("sp_status")%></td>
                      <td nowrap="nowrap"> <%=chkdate(rs("sp_date"))%> </td>
                      <td> <%=rs("sp_tech_code")%></td>
                      <td><%=rs("sp_tech_name")%></td>
                      <td nowrap="nowrap"> <%=rs("sp_tech_tel1")%> </td>
                      <td> <%=rs("sp_tech_city")%> </td>
                      <td align="center"> <%=chkdate(rs("sp_submitteddate"))%></td>
                      <td align="center" nowrap="nowrap"><%=chkdate(rs("sp_approveddate"))%></td>
                      <td align="center" nowrap="nowrap"><%=chkdate(rs("sp_confirmedreceiveddate"))%></td>
                      <td align="center"> <%=rs("sp_totalqty")%> </td>
                      <td align="right"> <%=chknumber2(rs("sp_totalAmt"))%> </td>
                    </tr>
                  
<%
sp_totalqty = sp_totalqty + rs("sp_totalqty")
sp_totalAmt = sp_totalAmt + rs("sp_totalAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>

  					<tr>
                      <td height="30" colspan="10" align="right" bgcolor="#CCCCCC"><strong> Total</strong></td>
                      <td height="30" align="center" bgcolor="#CCCCCC"><strong> <%=sp_totalqty%></strong></td>
                      <td height="30" align="right" bgcolor="#CCCCCC"><strong> <%=chknumber2(sp_totalAmt)%></strong></td>
                    </tr>
                    <tr>
                      <td height="30" colspan="10" align="right" bgcolor="#CCCCCC"><strong>Grand Total</strong></td>
                      <td height="30" align="center" bgcolor="#CCCCCC"><strong> <%=totalqty%></strong></td>
                      <td height="30" align="right" bgcolor="#CCCCCC"><strong> <%=chknumber2(totalAmt)%></strong></td>
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
					Response.Write " <a href='rm_consignment_view.asp?num=" & (j-1) * row & link & "&sp_status=" & sp_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_consignment_view.asp?num=" & Showed+row & link & "&sp_status=" & sp_status & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
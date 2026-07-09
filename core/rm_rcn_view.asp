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

if request("rcn_date_from") <> "" then
   rcn_date_from = request("rcn_date_from")
else
   rcn_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("rcn_date_to") <> "" then
   rcn_date_to = request("rcn_date_to")
else
   rcn_date_to = chkdate(date())
end if

if request("rcn_status") <> "" then
   rcn_status = request("rcn_status")
else

   if request.Cookies("GAPS")("slevel") = "technician" then 
   rcn_status = "Submitted"
   else
   rcn_status = "Open"
   end if
end if

i = 1
sql="SELECT tblrcn.rcn_id, tblrcn.rcn_no, tblrcn.rcn_date, tblrcn.rcn_status, tblrcn.rcn_job_code, tblrcn.rcn_onlineWrtyNo, tblrcn.rcn_SN_no, " & _ 
	"tblrcn.rcn_onlinewrtyStatus, tblrcn.rcn_modelcode,  " & _
	"tblrcn.rcn_modeltype, tblrcn.rcn_tech_code, tblrcn.rcn_cust_code, tblrcn.rcn_cust_name, tblrcn.rcn_cust_address, tblrcn.rcn_cust_postcode, " & _
	"tblrcn.rcn_cust_state, tblrcn.rcn_cust_state_id, tblrcn.rcn_cust_city, tblrcn.rcn_cust_city_id, tblrcn.rcn_cust_email, tblrcn.rcn_cust_tel1,  " & _
	"tblrcn.rcn_cust_tel2, tblrcn.rcn_remark,   " & _
	"tblrcn.rcn_createddate, tblrcn.rcn_createdby, tblrcn.rcn_submitteddate, tblrcn.rcn_submittedby,  " & _
	"tblrcn.rcn_posteddate, tblrcn.rcn_postedby, tblrcn.rcn_cancelleddate, tblrcn.rcn_cancelledby, tblrcn.rcn_totalqty, tblrcn.rcn_totalPartsAmt,  " & _
	"tblrcn.rcn_labourAmt, tblrcn.rcn_transportAmt, tblrcn.rcn_gstAmt, tblrcn.rcn_gstRate,  " & _
	"tblrcn.rcn_gstCode, tblrcn.rcn_totalAmt, tblrcn.rcn_emailsentdate, tbltechnician.tech_name, tbltechnician.tech_tel1  " & _
	"FROM tblrcn left join tbltechnician on tblrcn.rcn_tech_code = tbltechnician.tech_code where tblrcn.rcn_id is not null "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if Searchor_date = "Y" then
   sql = sql & " and  tblrcn.rcn_date >= '" & rcn_date_from & "' and tblrcn.rcn_date <= '" & rcn_date_to & "' "
end if

if request.Cookies("GAPS")("slevel") = "technician" then 
   sql = sql & " and tblrcn.rcn_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
end if

if rcn_status <> "All" and rcn_status <> "" then
   sql = sql & " and  tblrcn.rcn_status = '" & rcn_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tblrcn.rcn_id desc"
end if


'response.write request.Cookies("GAPS")("slevel") & "<br>"
'response.write sql

set rs = server.CreateObject("adodb.recordset")
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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&orderby=" & orderby & "&ordertype=" & ordertype & "&rcn_date_from=" & rcn_date_from & "&rcn_date_to=" & rcn_date_to 

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Return Credit Note (RCN)</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_rcn_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> RCN Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"><strong><font color="#000000"><strong>
                          <input name="rcn_date_from" type="text" id="rcn_date_from" value="<%=rcn_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.rcn_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="rcn_date_to" type="text" id="rcn_date_to" value="<%=rcn_date_to%>" size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.rcn_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td>
                          <select name="searchitem" id="searchitem">
                            <option value="tblrcn.rcn_no"  <% if searchitem = "tblrcn.rcn_no" then response.write " selected" %>>Invoice No</option>
                            <option value="tblrcn.rcn_cust_name" <% if searchitem = "tblrcn.rcn_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblrcn.rcn_cust_tel1" <% if searchitem = "tblrcn.rcn_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tblrcn.rcn_cust_email" <% if searchitem = "tblrcn.rcn_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tblrcn.rcn_tech_code" <% if searchitem = "tblrcn.rcn_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tblrcn.tech_name" <% if searchitem = "tblrcn.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tblrcn.rcn_cust_state" <% if searchitem = "tblrcn.rcn_cust_state" then response.write " selected" %>>Customer State</option>
                          </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                            <option value="tblrcn.rcn_no"  <% if orderby = "tblrcn.rcn_no" then response.write " selected" %>>Invoice No</option>
                            <option value="tblrcn.rcn_cust_name" <% if orderby = "tblrcn.rcn_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblrcn.rcn_cust_tel1" <% if orderby = "tblrcn.rcn_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tblrcn.rcn_cust_email" <% if orderby = "tblrcn.rcn_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tblrcn.rcn_tech_code" <% if orderby = "tblrcn.rcn_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tblrcn.tech_name" <% if orderby = "tblrcn.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tblrcn.rcn_cust_state" <% if orderby = "tblrcn.rcn_cust_state" then response.write " selected" %>>Customer State</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="Submit43" value="Display" />
                          <input name="rcn_status" type="hidden" id="rcn_status" value="<%=rcn_status%>" /></td>
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
					Response.Write " <a href='rm_rcn_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rcn_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      <td nowrap="nowrap"> <strong>Status</strong> : </td>
                      <td <%if rcn_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_rcn_view.asp?rcn_status=All<%=link%>"><font color="#FFFFFF"><strong>All</strong></font></a></td>
                      <td <%if rcn_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_rcn_view.asp?rcn_status=Open<%=link%>"><font color="#FFFFFF"><strong>Open</strong></font></a></td>
                      <td <%if rcn_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_rcn_view.asp?rcn_status=Submitted<%=link%>"><font color="#FFFFFF"> <strong>Submitted</strong></font></a></td>
                       <td <%if rcn_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_rcn_view.asp?rcn_status=Posted<%=link%>"><font color="#FFFFFF"> <strong>Posted</strong></font></a></td> 
                      <td <%if rcn_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_rcn_view.asp?rcn_status=Cancel<%=link%>"><font color="#FFFFFF"> <strong>Cancel</strong></font></a></td>
                    </tr>
                  </table>
                    <div align="right"></div>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>RCN  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>RCN  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer<br />
                      </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Online Wrty No. </strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Tel 1</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer City</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Technician </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Technician<span> Mobile</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted By/<br />
Date</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>GRN Amount <span>(RM)</span></strong></font></td>
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
                    
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40"><%=j%></td>
                      <td nowrap="nowrap"><strong><a href="rm_rcn_new.asp?rcn_no=<%=rs("rcn_no")%>"> <font color="#0000FF"><%=rs("rcn_no")%></font></a></strong><br />
                       <%=rs("rcn_status")%></td>
                      <td nowrap="nowrap"><%=chkdate(rs("rcn_date"))%>  </td>
                      <td> <%=rs("rcn_cust_name")%></td>
                      <td nowrap="nowrap"> <%=rs("rcn_onlineWrtyNo")%></td>
                      <td> <%=rs("rcn_cust_tel1")%></td>
                      <td> <%=rs("rcn_cust_city")%></td>
                      <td> <%=rs("tech_name")%></td>
                      <td> <%=rs("tech_tel1")%></td>
                      <td nowrap="nowrap"><%=rs("rcn_postedby")%> <br />
                        <%=chkdate(rs("rcn_posteddate"))%></td>
                      <td align="right"> <%=chknumber2(rs("rcn_totalAmt"))%> </td>
                    </tr>
                    
<%
rcn_totalAmt = rcn_totalAmt + rs("rcn_totalAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
  
                    <tr>
                      <td height="30" colspan="10" align="right" bgcolor="#CCCCCC"><strong>Grand Total</strong></td>
                      <td height="30" align="right" bgcolor="#CCCCCC"><strong> <%=chknumber2(rcn_totalAmt)%> </strong></td>
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
					Response.Write " <a href='rm_rcn_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rcn_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
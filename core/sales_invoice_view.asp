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

if request("inv_createddate_from") <> "" then
   inv_createddate_from = request("inv_createddate_from")
else
   inv_createddate_from = chkdate(DateAdd("d",-90,date()))
end if

if request("inv_createddate_to") <> "" then
   inv_createddate_to = request("inv_createddate_to")
else
   inv_createddate_to = chkdate(date())
end if

if request("inv_status") <> "" then
   inv_status = request("inv_status")
else

   if request.Cookies("GAPS")("slevel") = "technician" then 
   inv_status = "Submitted"
   else
   inv_status = "Open"
   end if
end if

i = 1
sql = "SELECT tblinvoice.inv_id, tblinvoice.inv_no, tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address, tblinvoice.inv_cust_postcode, " & _
"tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id, tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, " & _
"tblinvoice.inv_createddate, tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt, " & _
"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate, tblinvoice.inv_status, inv_approvedby, inv_approveddate, " & _
"tbltechnician.tech_name, tbltechnician.tech_tel1 FROM tblinvoice inner join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code where tblinvoice.inv_id is not null "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if Searchor_date = "Y" then
   sql = sql & " and  tblinvoice.inv_date >= '" & job_date_from & "' and tblinvoice.inv_date <= '" & job_date_to & "' "
end if

if request.Cookies("GAPS")("slevel") = "technician" then 
   sql = sql & " and tblinvoice.inv_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
end if

if inv_status <> "All" and inv_status <> "" then
   sql = sql & " and  tblinvoice.inv_status = '" & inv_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tblinvoice.inv_id desc"
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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby")

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Invoice</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="sales_invoice_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> Invoice Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"><strong><font color="#000000"><strong>
                          <input name="inv_createddate_from" type="text" id="inv_createddate_from" value="<%=inv_createddate_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.inv_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="inv_createddate_to" type="text" id="inv_createddate_to" value="<%=inv_createddate_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.inv_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td>
                          <select name="searchitem" id="searchitem">
                            <option value="tblinvoice.inv_no"  <% if searchitem = "tblinvoice.inv_no" then response.write " selected" %>>Invoice No</option>
                            <option value="tblinvoice.inv_cust_name" <% if searchitem = "tblinvoice.inv_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblinvoice.inv_cust_tel1" <% if searchitem = "tblinvoice.inv_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tblinvoice.inv_cust_email" <% if searchitem = "tblinvoice.inv_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tblinvoice.inv_tech_code" <% if searchitem = "tblinvoice.inv_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tbltechnician.tech_name" <% if searchitem = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tblinvoice.inv_cust_state" <% if searchitem = "tblinvoice.inv_cust_state" then response.write " selected" %>>Customer State</option>
                          </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                            <option value="tblinvoice.inv_no"  <% if orderby = "tblinvoice.inv_no" then response.write " selected" %>>Invoice No</option>
                            <option value="tblinvoice.inv_cust_name" <% if orderby = "tblinvoice.inv_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblinvoice.inv_cust_tel1" <% if orderby = "tblinvoice.inv_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tblinvoice.inv_cust_email" <% if orderby = "tblinvoice.inv_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tblinvoice.inv_tech_code" <% if orderby = "tblinvoice.inv_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tbltechnician.tech_name" <% if orderby = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tblinvoice.inv_cust_state" <% if orderby = "tblinvoice.inv_cust_state" then response.write " selected" %>>Customer State</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                        <input type="submit" name="Submit43" value="Display" /></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">
                                    <strong>Page</strong> <font color="3366ff"> 
                        <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
                        <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='sales_invoice_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='sales_invoice_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> 
                  </td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      <td nowrap="nowrap"><div align="center" class="titlegrey1"> Status :</div>          </td>
                      
                      <td <%if inv_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#461C13'" end if%>><a href="sales_invoice_view.asp?inv_status=Open"><font color="#FFFFFF"><strong>Open</strong></font></a></td>
                      <td <%if inv_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#461C13'" end if%>><a href="sales_invoice_view.asp?inv_status=Submitted"><font color="#FFFFFF"> <strong>Submitted</strong></font></a>                       </td>
                       <td <%if inv_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#461C13'" end if%>><a href="sales_invoice_view.asp?inv_status=Posted"><font color="#FFFFFF"> <strong>Posted</strong></font></a>                       </td>
                      <td <%if inv_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#461C13'" end if%>><a href="sales_invoice_view.asp?inv_status=Cancel"><font color="#FFFFFF"> <strong>Cancel</strong></font></a></td>
                    </tr>
                  </table>
                    <div align="right"></div>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Invoice  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Invoice  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Online Wrty No. </strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Mobile</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Location</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Technician </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Approved By</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Approve Date</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice Amount <span>(RM)</span></strong></font></td>
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
                      <td height="40"><div align="center"><%=j%> </div></td>
                      <td nowrap="nowrap"><strong><a href="sales_invoice_new.asp?inv_no=<%=rs("inv_no")%>"> <font color="#0000FF"><%=rs("inv_no")%></font></a></strong></td>
                      <td nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
                      <td nowrap="nowrap"><%=rs("inv_cust_name")%></td>
                      <td nowrap="nowrap">&nbsp;</td>
                      <td align="left"><%=rs("inv_cust_tel1")%></td>
                      <td align="left"><%=rs("inv_cust_city")%></td>
                      <td><%=rs("tech_name")%></td>
                      <td><%=rs("inv_approvedby")%></td>
                      <td align="center" nowrap="nowrap"><%=chkdate(rs("inv_approveddate"))%></td>
                      <td align="right"><%=chknumber2(rs("inv_totalAmt"))%> </td>
                    </tr>
<%
inv_totalAmt = inv_totalAmt + rs("inv_totalAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
                    <tr>
                      <td colspan="10" align="right"><strong>Grand Total</strong></td>
                      <td align="right"><strong> <%=chknumber2(inv_totalAmt)%> </strong></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">
                                    <strong>Page</strong> <font color="3366ff"> 
                        <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
                        <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='sales_invoice_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='sales_invoice_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
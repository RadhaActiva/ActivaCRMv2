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

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if request("job_status") <> "" then
   job_status = request("job_status")
else
   job_status = "Submitted"
end if

i = 1
sql = "SELECT tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
		"tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
		"tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
		"tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
		"tbljob.job_Model, tbljob.job_Model_desc, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
		"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
		"tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN,  " & _
		"tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
		"tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo, tbljob.job_hq_remark,  " & _
		"tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
		"tbljob.job_totalAmt, tbljob.job_amount_received, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
		"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbljob.job_submitforclaims, tbltechnician.tech_name, tbltechnician.tech_tel1 " & _
		"FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if Searchor_date = "Y" then
   sql = sql & " and  tbljob.job_date >= '" & job_date_from & "' and tbljob.job_date <= '" & job_date_to & "' "
end if

if request.Cookies("GAPS")("slevel") = "technician" then 
   sql = sql & " and tbljob.job_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
elseif request.Cookies("GAPS")("slevel") = "technician2" then 
   sql = sql & " and tbljob.job_createdby = '" & request.Cookies("GAPS")("sloginid") & "' " 
end if

if job_status <> "All" and job_status <> "" then
   sql = sql & " and  tbljob.job_status = '" & job_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tbljob.job_id desc"
end if

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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&sortby=" & sortby & "&job_date_from=" & job_date_from & "&job_date_to" & "&ordertype=" & ordertype

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left">View Job Sheet</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rmtech_jobsheet_view.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> Job Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"> <strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td><label for="select"></label>
                          <select name="searchitem" id="searchitem">
                            <option value="tbljob.job_code"  <% if searchitem = "tbljob.job_code" then response.write " selected" %>>Job No</option>
                            <option value="tbljob.job_cust_name" <% if searchitem = "tbljob.job_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tbljob.job_cust_tel1" <% if searchitem = "tbljob.job_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tbljob.job_cust_email" <% if searchitem = "tbljob.job_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tbljob.job_tech_code" <% if searchitem = "tbljob.job_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tbltechnician.tech_name" <% if searchitem = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tbljob.job_SN_no" <% if searchitem = "tbljob.job_SN_no" then response.write " selected" %>>Serial No</option>
                            <option value="tbljob.job_cust_state" <% if searchitem = "tbljob.job_cust_state" then response.write " selected" %>>Customer State</option>
                          </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>">
                          <select name="orderby" id="orderby">
                            <option value="tbljob.job_code"  <% if orderby = "tbljob.job_code" then response.write " selected" %>>Job No</option>
                            <option value="tbljob.job_cust_name" <% if orderby = "tbljob.job_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tbljob.job_cust_tel1" <% if orderby = "tbljob.job_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tbljob.job_cust_email" <% if orderby = "tbljob.job_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tbljob.job_tech_code" <% if orderby = "tbljob.job_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tbltechnician.tech_name" <% if orderby = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tbljob.job_SN_no" <% if orderby = "tbljob.job_SN_no" then response.write " selected" %>>Serial No</option>
                            <option value="tbljob.job_cust_state" <% if orderby = "tbljob.job_cust_state" then response.write " selected" %>>Customer State</option>
                          </select>
                         <select name="ordertype" id="ordertype">
                                  <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                                  <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                                </select> 
                          <input type="submit" name="Submit43" value="Display">
                        <input type="hidden" name="job_status" id="job_status" value="<%=job_status%>" /></td>
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
					Response.Write " <a href='rmtech_jobsheet_view.asp?num=" & (j-1) * row & link & "&job_status=" & job_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_jobsheet_view.asp?num=" & Showed+row & link & "&job_status=" & job_status & "'> Next >></a>"
	End If  %> 
                  
                  </td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table width="50%" border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      <td nowrap="nowrap"><div align="center" class="titlegrey1"> Status :</div><div align="center" class="titlewhite1"></div></td>
                      <td width="7%" <%if job_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><div align="center" class="titlewhite1"> <a href="rmtech_jobsheet_view.asp?job_status=Submitted<%=link%>"><font color="#FFFFFF"> Submitted</font></a></div></td>
                      <td width="9%" <%if job_status="Accepted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><div align="center" class="titlewhite1"> <a href="rmtech_jobsheet_view.asp?job_status=Accepted<%=link%>"><font color="#FFFFFF"> Accepted</font></a></div></td>
                      <td width="5%" <%if job_status="Done" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><div align="center" class="titlewhite1"> <a href="rmtech_jobsheet_view.asp?job_status=Done<%=link%>"><font color="#FFFFFF">Done</font></a></div></td>
                      <td width="11%" <%if job_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><div align="center" class="titlewhite1"> <a href="rmtech_jobsheet_view.asp?job_status=Posted<%=link%>"><font color="#FFFFFF"> Posted</font></a></div></td>
                      <td width="12%" <%if job_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><div align="center" class="titlewhite1"> <a href="rmtech_jobsheet_view.asp?job_status=Cancel<%=link%>"><font color="#FFFFFF"> Cancel</font></a></div></td>
                    </tr>
                  </table>
                    <div align="right"></div>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Job  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Job  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer<br />
                      </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Model No </strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Mobile</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer State</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Technician </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Technician<span> Mobile</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Amount <span>(RM)</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Action</strong></font></td>
                    </tr>

<form name="formorder" id="formorder" method="post" action="action.asp?type=JobActionClaimandAccept" >

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
                      <td nowrap="nowrap"><strong><a href="rmtech_jobsheet.asp?job_code=<%=rs("job_code")%>"> <font color="#0000FF"><%=rs("job_code")%> </font></a></strong></td>
                      <td nowrap="nowrap"><%=chkdate(rs("job_date"))%></td>
                      <td><%=rs("job_cust_name")%></td>
                      <td><%=rs("job_Model_desc")%></td>
                      <td><%=rs("job_cust_tel1")%></td>
                      <td><%=rs("job_cust_state")%><br />
                      <%=rs("job_cust_city")%>
                      </td>
                      <td><%=rs("tech_name")%></td>
                      <td align="center"> <%=rs("tech_tel1")%></td>                      
                      <td align="right"> <%=chknumber2(rs("job_totalAmt"))%> </td> <!--this is the amount CS calculated based on totalparts + totallablour + totaltransport -->
                        

                      <td align="right">
                       <%if job_status="Submitted" then %>
                            <input type="checkbox" name="Accepted" id="Accepted" value="<%=rs("job_code")%>" />                         
                       <%end if%>
                     
                     <!--  02/10/23 implemented logic for tech checkbox. Ticked ones cannot be undone-->
                      <%if job_status="Done" then %> 
                          <%if rs("job_submitforclaims") = "Yes" then%>                          
                            <input type="checkbox" name="Claimed" id="Claimed" checked disabled="disabled" />
                          <%else%> 
                            <input type="checkbox" name="Claim" id="Claim" value="<%=rs("job_code")%>" />
                          <%end if%>
                      <%end if%>
                      </td>
                    </tr>
 <%
job_totalAmt = job_totalAmt + rs("job_totalAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>

                    <tr>
                      <td colspan="9" align="right"><strong>Grand Total</strong></td>
                      <td align="right"><strong> <%=chknumber2(job_totalAmt)%></strong></td>
                      <td align="right">
                      <%if job_status="Submitted" then %>
                        <input type="submit" name="button2" id="button" value="Accept Jobs" />
                      <%end if%>   

                      <%if job_status="Done" then %>
                        <input type="submit" name="button2" id="claimbutton" value="Submit for Claim" />
                      <%end if%>
                      </td>
                    </tr>
</form>                    
                    
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> 
                        <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
                        <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rmtech_jobsheet_view.asp?num=" & (j-1) * row & link & "&job_status=" & job_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rmtech_jobsheet_view.asp?num=" & Showed+row & link & "&job_status=" & job_status & "'> Next >></a>"
	End If
	
                    %> </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
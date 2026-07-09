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

i = 1
sql = "SELECT tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
		"tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
		"tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
		"tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
		"tbljob.job_Model, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
		"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
		"tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN,  " & _
		"tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
		"tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo, tbljob.job_hq_remark,  " & _
		"tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
		"tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
		"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbltechnician.tech_name, tbltechnician.tech_tel1 " & _
		"FROM tbljob left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
        "and tbljob.job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' and tbljob.job_status <> 'Posted' "
else
   sql = sql & " and tbljob.job_status <> 'Posted' "
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Job Sheet</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="sales_jobsheet_view.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                        <td nowrap="nowrap" class="titlegrey1"> Job Date <br/></td>
                        <td width="84%"><div align="left"> <strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015
                        <input name="job_status" type="hidden" id="job_status" value="<%=job_status%>" />
                        </div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td width="84%"><label for="select"></label>
                          <select name="searchitem" id="searchitem">
                            <option value="tbljob.job_code"  <% if searchitem = "tbljob.job_code" then response.write " selected" %>>Job No</option>
                            <option value="tbljob.job_cust_name" <% if searchitem = "tbljob.job_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tbljob.job_cust_tel1" <% if searchitem = "tbljob.job_cust_tel1" then response.write " selected" %>>Customer Mobile</option>
                            <option value="tbljob.job_SN_no" <% if searchitem = "tbljob.job_SN_no" then response.write " selected" %>>Serial No</option>
                          </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>">
                          <select name="orderby" id="orderby">
                            <option value="tbljob.job_code"  <% if orderby = "tbljob.job_code" then response.write " selected" %>>Job No</option>
                            <option value="tbljob.job_cust_name" <% if orderby = "tbljob.job_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tbljob.job_cust_tel1" <% if orderby = "tbljob.job_cust_tel1" then response.write " selected" %>>Customer Mobile</option>
                            <option value="tbljob.job_SN_no" <% if orderby = "tbljob.job_SN_no" then response.write " selected" %>>Serial No</option>
                          </select>
                         <select name="ordertype" id="ordertype">
                                  <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                                  <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                                </select> 
                          <input type="submit" name="Submit43" value="Display">
                        </td>
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
					Response.Write " <a href='sales_jobsheet_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='sales_jobsheet_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> 
                  
                  </td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=norecord%></font></strong></td>
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
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                     <!-- <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Amount <span>(RM)</span></strong></font></td>-->
                    </tr>

<form name="formorder" id="formorder" method="post" action="action.asp?type=Acceptedjob" >

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
                      <td nowrap="nowrap"><strong><a href="sales_jobsheet_new.asp?job_code=<%=rs("job_code")%>"> <font color="#0000FF"><%=rs("job_code")%> </font></a></strong></td>
                      <td nowrap="nowrap"><%=chkdate(rs("job_date"))%></td>
                      <td width="15%"><%=rs("job_cust_name")%></td>
                      <td nowrap="nowrap"><%=rs("job_Model")%></td>
                      <td><%=rs("job_cust_tel1")%></td>
                      <td><%=rs("job_cust_state")%></td>
                      <td><%=rs("tech_name")%></td>
                      <td align="center"> <%=rs("job_status")%></td>
                      <!--<td align="right"> <%=chknumber2(rs("job_totalAmt"))%> </td>-->
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
					Response.Write " <a href='sales_jobsheet_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='sales_jobsheet_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
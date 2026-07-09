<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            height: 30px;
        }
    </style>
</head>
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")

if ordertype = "" then 
   ordertype = "desc"
end if

if request("do_createddate_from") <> "" then
   do_createddate_from = request("do_createddate_from")
else
   do_createddate_from = chkdate(DateAdd("d",-90,date()))
end if

if request("do_createddate_to") <> "" then
   do_createddate_to = request("do_createddate_to")
else
   do_createddate_to = chkdate(date())
end if

if request("do_status") <> "" then
   do_status = request("do_status")
else
   do_status = "Open"
end if

sql = "SELECT sum(do_totalqty) as grandtotal_qty FROM tbldo where do_id is not null " & _
		"and  tbldo.do_date >= '" & ChkDateYYYYMMDD(do_createddate_from) & "' and tbldo.do_date <= '" & ChkDateYYYYMMDD(do_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if do_status <> "All" and do_status <> "" then
   sql = sql & " and  tbldo.do_status = '" & do_status & "' "
end if
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if not rs.eof then 
   grandtotal_qty = rs("grandtotal_qty")
end if
rs.close


i = 1
sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, do_cust_state, " & _ 
		"do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby, do_job_code, do_tech_code,  " & _
		"do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, do_emailsentdate, do_deliveredby,  " & _
		"do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate, do_purchase_date, do_onlineWrtyNo,  " & _
		"do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time, do_appointment_remark " & _
		"FROM tbldo where do_id is not null " & _
		"and  tbldo.do_date >= '" & ChkDateYYYYMMDD(do_createddate_from) & "' and tbldo.do_date <= '" & ChkDateYYYYMMDD(do_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if do_status <> "All" and do_status <> "" then
   sql = sql & " and  tbldo.do_status = '" & do_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tbldo.do_id desc"
end if

'response.write request.Cookies("GAPS")("slevel") & "<br>"
'response.write sql
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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&do_status=" & do_status & "&do_createddate_from=" & do_createddate_from  & "&do_createddate_to=" & do_createddate_to 
slink = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&do_createddate_from=" & do_createddate_from  & "&do_createddate_to=" & do_createddate_to 
%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>DO</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_do_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> DO Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"><strong><font color="#000000"><strong>
                          <input name="do_createddate_from" type="text" id="do_createddate_from" value="<%=do_createddate_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.do_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="do_createddate_to" type="text" id="do_createddate_to" value="<%=do_createddate_to%>" size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.do_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td><select name="searchitem" id="searchitem">
                          <option value="tbldo.do_no"  <% if searchitem = "tbldo.do_no" then response.write " selected" %>>DO No</option>
                          <option value="tbldo.do_cust_name" <% if searchitem = "tbldo.do_cust_name" then response.write " selected" %>>Customer Name</option>
                          <option value="tbldo.do_cust_tel1" <% if searchitem = "tbldo.do_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tbldo.do_cust_email" <% if searchitem = "tbldo.do_cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tbldo.do_cust_city" <% if searchitem = "tbldo.do_cust_city" then response.write " selected" %>>Customer City</option>
                          <option value="tbldo.do_inv_no" <% if searchitem = "tbldo.do_inv_no" then response.write " selected" %>>Invoice No</option>
                        </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                           <option value="tbldo.do_no"  <% if orderby = "tbldo.do_no" then response.write " selected" %>>DO No</option>
                          <option value="tbldo.do_cust_name" <% if orderby = "tbldo.do_cust_name" then response.write " selected" %>>Customer Name</option>
                          <option value="tbldo.do_cust_tel1" <% if orderby = "tbldo.do_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                          <option value="tbldo.do_cust_email" <% if orderby = "tbldo.do_cust_email" then response.write " selected" %>>Customer Email</option>
                          <option value="tbldo.do_cust_city" <% if orderby = "tbldo.do_cust_city" then response.write " selected" %>>Customer City</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="Submit43" value="Display" />
                          <input name="do_status" type="hidden" id="do_status" value="<%=do_status%>" /></td>
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
					Response.Write " <a href='rm_do_view.asp?num=" & (j-1) * row & link & "&do_status=" & do_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_do_view.asp?num=" & Showed+row & link & "&do_status=" & do_status & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="5" cellspacing="1">
                    <tr>
                      <td nowrap="nowrap" class="auto-style1"><div align="center" class="titlegrey1"> Status :</div> </td>
                      <td <%if do_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%> class="auto-style1"><strong><a href="rm_do_view.asp?do_status=All<%=slink%>"><font color="#FFFFFF">All</font></a></strong></td>
                      <td <%if do_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%> class="auto-style1"><strong><a href="rm_do_view.asp?do_status=Open<%=slink%>"><font color="#FFFFFF">Open</font></a></strong></td>
                      <td <%if do_status="Delivered" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%> class="auto-style1"><strong><a href="rm_do_view.asp?do_status=Delivered<%=slink%>"><font color="#FFFFFF">Delivered</font></a></strong></td>
                      <td <%if do_status="Done" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%> class="auto-style1"><strong><a href="rm_do_view.asp?do_status=Done<%=slink%>"><font color="#FFFFFF">Done</font></a></strong></td>
                      <td <%if do_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%> class="auto-style1"><strong><a href="rm_do_view.asp?do_status=Posted<%=slink%>"><font color="#FFFFFF">Posted</font></a></strong></td>
                      <td <%if do_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%> class="auto-style1"><strong><a href="rm_do_view.asp?do_status=Cancel<%=slink%>"><font color="#FFFFFF">Cancel</font></a></strong></td>
                    </tr>
                  </table>
                    <div align="right"></div>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>DO  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>DO  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Mobile</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer City</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Delivered<span> Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Done Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>DO Qty</strong></font></td>
                    </tr>
                    
<% 
do_totalqty = 0
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
                      <td nowrap="nowrap"><strong><font color="#0000FF"><a href="rm_do_new.asp?do_no=<%=rs("do_no")%>"><%=rs("do_no")%></a></font></strong><br /><%=rs("do_status")%></td>
                      <td nowrap="nowrap"> <%=chkdate(rs("do_date"))%></td>
                      <td> <%=rs("do_cust_name")%></td>
                      <td> <%=rs("do_onlineWrtyNo")%></td>
                      <td> <%=rs("do_cust_tel1")%></td>
                      <td> <%=rs("do_cust_city")%></td>
                      <td> <%=chkdate(rs("do_delivereddate"))%></td>
                      <td> <%=chkdate(rs("do_donedate"))%></td>
                      <td> <%=chkdate(rs("do_posteddate"))%></td>
                      <td><%=rs("do_inv_no")%></td>
                      <td nowrap="nowrap"><%=chkdate(rs("do_inv_date"))%></td>
                      <td align="center"> <%=rs("do_totalqty")%></td>
                    </tr>
                    
                    
<%
do_totalqty = do_totalqty + rs("do_totalqty")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
 <tr>
                       <td height="30" colspan="12" align="right" bgcolor="#CCCCCC"><strong> Total DO Qty</strong></td>
                       <td height="30" align="center" bgcolor="#CCCCCC"><strong> <%=do_totalqty%></strong></td>
                     </tr>
                    
                    <tr>
                      <td height="30" colspan="12" align="right" bgcolor="#CCCCCC"><strong> Grand Total DO Qty</strong></td>
                      <td height="30" align="center" bgcolor="#CCCCCC"><strong> <%=grandtotal_qty%> </strong></td>
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
					Response.Write " <a href='rm_do_view.asp?num=" & (j-1) * row & link  & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_do_view.asp?num=" & Showed+row & link  & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
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

if request("purchase_date_from") <> "" then
   purchase_date_from = request("purchase_date_from")
else
   purchase_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("purchase_date_to") <> "" then
   purchase_date_to = request("purchase_date_to")
else
   purchase_date_to = chkdate(date())
end if

if request("warranty_status") <> "" then
   warranty_status = request("warranty_status")
else
   warranty_status = "All"
end if

i = 1
sql = "SELECT " & _
		"tblonlinewarranty.refer_id, tblonlinewarranty.warrantyno, tblonlinewarranty.productmodel,  " & _
		"tblonlinewarranty.othermodel, tblonlinewarranty.serialno, tblonlinewarranty.dealername,  " & _
		"tblonlinewarranty.purchase_date, tblonlinewarranty.invoiceno,  " & _
		"tblonlinewarranty.deliveryno, tblonlinewarranty.customername, tblonlinewarranty.customericno, " & _ 
		"tblonlinewarranty.customeremail, tblonlinewarranty.customeraddress, tblonlinewarranty.customerpostcode,  " & _
		"tblonlinewarranty.customerstate, tblonlinewarranty.customercity, tblonlinewarranty.customertel1,  " & _
		"tblonlinewarranty.customertel2, tblonlinewarranty.customerfax, tbljob.job_code, tbljob.job_date " & _
		"FROM tblonlinewarranty left join tbljob on  " & _
		"tblonlinewarranty.warrantyno=tbljob.job_onlineWrtyNo  " & _
		"WHERE tblonlinewarranty.refer_id is not null  "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if Searchor_date = "Y" then
   sql = sql & " and  tblonlinewarranty.purchase_date >= '" & job_date_from & "' and tblonlinewarranty.purchase_date <= '" & job_date_to & "' "
end if

if warranty_status <> "All" and warranty_status <> "" then
   if warranty_status = "Over" then 
      sql = sql & " and  DATE_ADD(tblonlinewarranty.purchase_date,INTERVAL 365 DAY)  < now() "
   else
      sql = sql & " and  DATE_ADD(tblonlinewarranty.purchase_date,INTERVAL 365 DAY)  > now() "
   end if
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tblonlinewarranty.warrantyno desc"
end if

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

%>      <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Online Warranty</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_warrantty_view.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> Warranty Registed Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"><strong><font color="#000000"><strong>
                          <input name="purchase_date_from" type="text" id="purchase_date_from" value="<%=purchase_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.purchase_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="purchase_date_to" type="text" id="purchase_date_to" value="<%=purchase_date_to%>" size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.purchase_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td>
                          <select name="searchitem" id="searchitem">
                            <option value="tblonlinewarranty.warrantyno" <% if searchitem = "tblonlinewarranty.warrantyno" then response.write " selected" %>>Online Warranty No</option>
                            <option value="tblonlinewarranty.customername" <% if searchitem = "tblonlinewarranty.customername" then response.write " selected" %>>Customer Name</option>
                            <option value="tblonlinewarranty.customertel1" <% if searchitem = "tblonlinewarranty.customertel1" then response.write " selected" %>>Customer Mobile Number</option>
                            <option value="tblonlinewarranty.serialno" <% if searchitem = "tblonlinewarranty.serialno" then response.write " selected" %>>SN No.</option>
                            <option value="tblonlinewarranty.customeremail" <% if searchitem = "tblonlinewarranty.customeremail" then response.write " selected" %>>Email </option>
                            <option value="tbljob.job_code" <% if searchitem = "tbljob.job_code" then response.write " selected" %>>job Code </option>
                          </select>
                           <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                            <option value="tblonlinewarranty.warrantyno" <% if orderby = "tblonlinewarranty.warrantyno" then response.write " selected" %>>Online Warranty No</option>
                            <option value="tblonlinewarranty.customername" <% if orderby = "tblonlinewarranty.customername" then response.write " selected" %>>Customer Name</option>
                            <option value="tblonlinewarranty.customertel1" <% if orderby = "tblonlinewarranty.customertel1" then response.write " selected" %>>Customer Mobile Number</option>
                            <option value="tblonlinewarranty.serialno" <% if orderby = "tblonlinewarranty.serialno" then response.write " selected" %>>SN No.</option>
                            <option value="tblonlinewarranty.customeremail" <% if orderby = "tblonlinewarranty.customeremail" then response.write " selected" %>>Email </option>
                            <option value="tbljob.job_code" <% if orderby = "tbljob.job_code" then response.write " selected" %>>job Code </option>
                          </select>
                           <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="button" id="button3" value="Submit" />
                          </td>
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
					Response.Write " <a href='rm_warrantty_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_warrantty_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Purchased Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Expiry Date</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Mobile</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> State</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> City</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Model</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Serial No.</span></strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Dealer</span></strong></font></td>
                      <td align="right" nowrap="nowrap" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Last Job No.</strong></font></td>
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

if isdate(rs("purchase_date")) then 
   expiry_date = dateadd("d", 365, ChkDateYYYYMMDD(rs("purchase_date")))
else
   expiry_date = dateadd("d", 365, date())
end if

%>
                   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40"> <div align="center"><%=j%> </div></td>
                      <td><strong><a href="rm_warrantty_new.asp?warrantyno=<%=rs("warrantyno")%>"> <font color="#0000FF"><%=rs("warrantyno")%></font></a></strong></td>
                      <td nowrap="nowrap">
					  <%if DATEDIFF("d",chkdate(expiry_date),chkdate(date()))  > 1 then %>
                      <font color="#FF0000"><%=chkdate(rs("purchase_date"))%></font>
                      <%else%>
                      <font color="#000000"><%=chkdate(rs("purchase_date"))%></font>
                      <%end if%>
                      </td>
                      <td nowrap="nowrap">
					  <%if DATEDIFF("d",chkdate(expiry_date),chkdate(date()))  > 1 then %>
                      <font color="#FF0000"><%=chkdate(expiry_date)%></font>
                      <%else%>
                      <font color="#000000"><%=chkdate(expiry_date)%></font>
                      <%end if%>
					  </td>
                      <td><%=rs("customername")%></td>
                      <td><%=rs("customertel1")%><br><%=rs("customertel2")%></td>
                      <td><%=rs("customerstate")%></td>
                      <td><%=rs("customercity")%></td>
                      <td><%=rs("productmodel")%></td>
                      <td><%=rs("serialno")%></td>
                      <td align="left"><%=rs("dealername")%></td>
                      <td align="right"><a href="rm_jobsheet.asp?job_code=<%=rs("job_code")%>" target="_blank"><%=rs("job_code")%></a></td>
                    </tr>
                    
<%
 count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>      
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
					Response.Write " <a href='rm_warrantty_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_warrantty_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
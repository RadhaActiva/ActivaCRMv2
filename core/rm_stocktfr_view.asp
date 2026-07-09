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

if request("sf_createddate_from") <> "" then
   sf_createddate_from = request("sf_createddate_from")
else
   sf_createddate_from = chkdate(DateAdd("d",-90,date()))
end if

if request("sf_createddate_to") <> "" then
   sf_createddate_to = request("sf_createddate_to")
else
   sf_createddate_to = chkdate(date())
end if

if request("sf_status") <> "" then
   sf_status = request("sf_status")
else
   sf_status = "Open"
end if

i = 1
sql = "SELECT sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, " & _
		"sf_approveddate, sf_approvedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate " & _
		"FROM tblstocktransfer where sf_id is not null " & _
		"and  sf_date >= '" & ChkDateYYYYMMDD(sf_createddate_from) & "' and sf_date <= '" & ChkDateYYYYMMDD(sf_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if sf_status <> "All" and sf_status <> "" then
   sql = sql & " and  sf_status = '" & sf_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by sf_no desc"
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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&sf_status=" & sf_status & "&sf_createddate_from=" & sf_createddate_from  & "&sf_createddate_to=" & sf_createddate_to 
slink = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&sf_createddate_from=" & sf_createddate_from  & "&sf_createddate_to=" & sf_createddate_to 
%>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Stock Transfer</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="66%" valign="top"><div align="left">
                        <form name="form1" id="form1" method="post" action="rm_stockTfr_view.asp?type=reset">
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td nowrap="nowrap" class="titlegrey1"><strong> DO Date <br />
                              </strong></td>
                              <td><div align="left"><strong><font color="#000000"><strong>
                                <input name="sf_createddate_from" type="text" id="sf_createddate_from" value="<%=sf_createddate_from%>" size="15" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.form1.sf_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                                <input name="sf_createddate_to" type="text" id="sf_createddate_to" value="<%=sf_createddate_to%>" size="12" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.form1.sf_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                            </tr>
                            <tr>
                              <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                              <td><select name="searchitem" id="searchitem">
                                <option value="tblstocktransfer.sf_no"  <% if searchitem = "tblstocktransfer.sf_no" then response.write " selected" %>>Stock-Transfer No</option>
                                <option value="tblstocktransfer.sf_referenceno" <% if searchitem = "tblstocktransfer.sf_referenceno" then response.write " selected" %>>Reference No</option>
                                <option value="tblstocktransfer.sf_fromwarehouse" <% if searchitem = "tblstocktransfer.sf_fromwarehouse" then response.write " selected" %>>from Store</option>
                                <option value="tblstocktransfer.sf_towarehouse" <% if searchitem = "tblstocktransfer.sf_towarehouse" then response.write " selected" %>>to Store</option>
                                <option value="tblstocktransfer.sf_remark" <% if searchitem = "tblstocktransfer.sf_remark" then response.write " selected" %>>Remark</option>
                              </select>
                                <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                                <select name="orderby" id="orderby">
                                 <option value="tblstocktransfer.sf_no"  <% if orderby = "tblstocktransfer.sf_no" then response.write " selected" %>>Stock-Transfer No</option>
                                <option value="tblstocktransfer.sf_referenceno" <% if orderby = "tblstocktransfer.sf_referenceno" then response.write " selected" %>>Reference No</option>
                                <option value="tblstocktransfer.sf_fromwarehouse" <% if orderby = "tblstocktransfer.sf_fromwarehouse" then response.write " selected" %>>from Store</option>
                                <option value="tblstocktransfer.sf_towarehouse" <% if orderby = "tblstocktransfer.sf_towarehouse" then response.write " selected" %>>to Store</option>
                                <option value="tblstocktransfer.sf_remark" <% if orderby = "tblstocktransfer.sf_remark" then response.write " selected" %>>Remark</option>
                                </select>
                                <select name="ordertype" id="ordertype">
                                  <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                                  <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                                </select>
                                <input type="submit" name="Submit43" value="Display" />
                                <input name="sf_status" type="hidden" id="sf_status" value="<%=sf_status%>" /></td>
                            </tr>
                          </table>
                        </form>
                      </div></td>
                      <td width="18%">&nbsp;</td>
                      <td width="16%" align="right" valign="top"><input type="button" name="Submit333" value="Create New Stock-Transfer" onclick="document.location.href='rm_stocktfr_new.asp'" />
                        <br />
                        <font size="4"><strong><font color="#FF0000"></font></strong></font></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_stockTfr_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockTfr_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
                    <tr>
                      <td colspan="9" align="right">&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="9" align="center"><table border="0" align="right" cellpadding="5" cellspacing="0">
                        <tr>
                          <td nowrap="nowrap"><div align="center" class="titlegrey1"> Status :</div></td>
                          <td <%if sf_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockTfr_view.asp?sf_status=All<%=slink%>"><font color="#FFFFFF">All</font></a></strong></td>
                          <td <%if sf_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockTfr_view.asp?sf_status=Open<%=slink%>"><font color="#FFFFFF">Open</font></a></strong></td>
                          <td <%if sf_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockTfr_view.asp?sf_status=Submitted<%=slink%>"><font color="#FFFFFF">Submitted</font></a></strong></td>
                          <td <%if sf_status="Approved" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockTfr_view.asp?sf_status=Approved<%=slink%>"><font color="#FFFFFF">Approved</font></a></strong></td>
                          <td <%if sf_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockTfr_view.asp?sf_status=Cancel<%=slink%>"><font color="#FFFFFF">Cancel</font></a></strong></td>
                        </tr>
                      </table>
                        </td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#475387"><font color="#FFFFFF"><strong>No.</strong></font></td>
                      <td bgcolor="#475387"><font color="#FFFFFF"><strong>Stock-Transfer No</strong></font></td>
                      <td bgcolor="#475387"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                      <td bgcolor="#475387"><font color="#FFFFFF"><strong>Stock-Transfer Date</strong></font></td>
                      <td bgcolor="#475387"><strong><font color="#FFFFFF">From Store</font></strong></td>
                      <td bgcolor="#475387"><font color="#FFFFFF"><strong>To Store</strong></font></td>
                      <td align="right" bgcolor="#475387"><strong><font color="#FFFFFF">Total Qty</font></strong></td>
                      <td align="right" bgcolor="#475387"><font color="#FFFFFF"><strong>Total Amt</strong></font></td>
                      <td align="right" bgcolor="#475387"><font color="#FFFFFF"><strong>Last updated by</strong></font></td>
                    </tr>
                    
<% 
sf_totalqty = 0
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
                      <td><strong><font color="#000000"><a href="rm_stocktfr_new.asp?sf_no=<%=rs("sf_no")%>"><%=rs("sf_no")%></a></font></strong><br><%=rs("sf_status")%></td>
                      <td><%=rs("sf_remark")%></td>
                      <td><%=chkdate(rs("sf_date"))%></td>
                      <td><%=rs("sf_fromwarehouse")%></td>
                      <td><%=rs("sf_towarehouse")%></td>
                      <td align="right"><%=rs("sf_totalqty")%></td>
                      <td align="right"><%=rs("sf_totalaAmt")%></td>
                      <td align="right"><%=rs("sf_approvedby")%><br />
                      <%=chkdatetime(rs("sf_approveddate"))%></td>
                    </tr>
                    
<%
sf_totalqty = sf_totalqty + rs("sf_totalqty")
sf_totalaAmt = sf_totalaAmt + rs("sf_totalaAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
                    <tr>
                      <td height="30" colspan="6" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td height="30" align="right" bgcolor="#CCCCCC"><strong><%=sf_totalqty%></strong></td>
                      <td height="30" align="right" bgcolor="#CCCCCC"><%=sf_totalaAmt%></td>
                      <td height="30" align="center" bgcolor="#CCCCCC">&nbsp;</td>
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
					Response.Write " <a href='rm_stockTfr_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockTfr_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
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

if request("sj_createddate_from") <> "" then
   sj_createddate_from = request("sj_createddate_from")
else
   sj_createddate_from = chkdate(DateAdd("d",-90,date()))
end if

if request("sj_createddate_to") <> "" then
   sj_createddate_to = request("sj_createddate_to")
else
   sj_createddate_to = chkdate(date())
end if

if request("sj_status") <> "" then
   sj_status = request("sj_status")
else
   sj_status = "Open"
end if

i = 1
sql = "SELECT sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, " & _
		"sj_approveddate, sj_approvedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
		"FROM tblstockadj where sj_id is not null " & _
		"and  sj_date >= '" & ChkDateYYYYMMDD(sj_createddate_from) & "' and sj_date <= '" & ChkDateYYYYMMDD(sj_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if sj_status <> "All" and sj_status <> "" then
   sql = sql & " and  sj_status = '" & sj_status & "' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by sj_no desc"
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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&sj_status=" & sj_status & "&sj_createddate_from=" & sj_createddate_from  & "&sj_createddate_to=" & sj_createddate_to 
slink = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&sj_createddate_from=" & sj_createddate_from  & "&sj_createddate_to=" & sj_createddate_to
%>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Stock Adjustment (ADJ)</div></td>
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
                        <form name="form1" id="form1" method="post" action="rm_stockAdj_view.asp?type=reset">
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td nowrap="nowrap" class="titlegrey1"><strong> DO Date <br />
                              </strong></td>
                              <td><div align="left"><strong><font color="#000000"><strong>
                                <input name="sj_createddate_from" type="text" id="sj_createddate_from" value="<%=sj_createddate_from%>" size="15" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.form1.sj_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                                <input name="sj_createddate_to" type="text" id="sj_createddate_to" value="<%=sj_createddate_to%>" size="12" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.form1.sj_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                            </tr>
                            <tr>
                              <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                              <td><select name="searchitem" id="searchitem">
                                <option value="tblstockadj.sj_no"  <% if searchitem = "tblstockadj.sj_no" then response.write " selected" %>>Stock-In No</option>
                                <option value="tblstockadj.sj_referenceno" <% if searchitem = "tblstockadj.sj_referenceno" then response.write " selected" %>>Reference No</option>
                                <option value="tblstockadj.sj_fromwarehouse" <% if searchitem = "tblstockadj.sj_fromwarehouse" then response.write " selected" %>>from Store</option>
                                <option value="tblstockadj.sj_towarehouse" <% if searchitem = "tblstockadj.sj_towarehouse" then response.write " selected" %>>to Store</option>
                                <option value="tblstockadj.sj_remark" <% if searchitem = "tblstockadj.sj_remark" then response.write " selected" %>>Remark</option>
                              </select>
                                <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                                <select name="orderby" id="orderby">
                                 <option value="tblstockadj.sj_no"  <% if orderby = "tblstockadj.sj_no" then response.write " selected" %>>Stock-In No</option>
                                <option value="tblstockadj.sj_referenceno" <% if orderby = "tblstockadj.sj_referenceno" then response.write " selected" %>>Reference No</option>
                                <option value="tblstockadj.sj_fromwarehouse" <% if orderby = "tblstockadj.sj_fromwarehouse" then response.write " selected" %>>from Store</option>
                                <option value="tblstockadj.sj_towarehouse" <% if orderby = "tblstockadj.sj_towarehouse" then response.write " selected" %>>to Store</option>
                                <option value="tblstockadj.sj_remark" <% if orderby = "tblstockadj.sj_remark" then response.write " selected" %>>Remark</option>
                                </select>
                                <select name="ordertype" id="ordertype">
                                  <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                                  <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                                </select>
                                <input type="submit" name="Submit43" value="Display" />
                                <input name="sj_status" type="hidden" id="sj_status" value="<%=sj_status%>" /></td>
                            </tr>
                          </table>
                        </form>
                      </div></td>
                      <td width="18%">&nbsp;</td>
                      <td width="16%" align="right" valign="top"><input type="button" name="Submit333" value="Create New Stock-Adjustment" onclick="document.location.href='rm_stockAdj_new.asp'" />
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
					Response.Write " <a href='rm_stockAdj_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockAdj_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
                    <tr>
                      <td colspan="8" align="right">&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="8" align="center"><table border="0" align="right" cellpadding="5" cellspacing="0">
                        <tr>
                          <td nowrap="nowrap"><div align="center" class="titlegrey1"> Status :</div></td>
                          <td <%if sj_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockAdj_view.asp?sj_status=All<%=slink%>"><font color="#FFFFFF">All</font></a></strong></td>
                          <td <%if sj_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockAdj_view.asp?sj_status=Open<%=slink%>"><font color="#FFFFFF">Open</font></a></strong></td>
                          <td <%if sj_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockAdj_view.asp?sj_status=Submitted<%=slink%>"><font color="#FFFFFF">Submitted</font></a></strong></td>
                          <td <%if sj_status="Approved" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockAdj_view.asp?sj_status=Approved<%=slink%>"><font color="#FFFFFF">Approved</font></a></strong></td>
                          <td <%if sj_status="Cancel" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><strong><a href="rm_stockAdj_view.asp?sj_status=Cancel<%=slink%>"><font color="#FFFFFF">Cancel</font></a></strong></td>
                        </tr>
                      </table>
                        </td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#475387"><font color="#FFFFFF"><strong>No.</strong></font></td>
                      <td bgcolor="#475387"><font color="#FFFFFF"><strong>Stock-Adj No</strong></font></td>
                      <td bgcolor="#475387"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                      <td bgcolor="#475387"><font color="#FFFFFF"><strong>Stock-Adj Date</strong></font></td>
                      <td bgcolor="#475387"><strong><font color="#FFFFFF">Store</font></strong></td>
                      <td align="right" bgcolor="#475387"><strong><font color="#FFFFFF">Total Qty</font></strong></td>
                      <td align="right" bgcolor="#475387"><font color="#FFFFFF"><strong>Total Amt</strong></font></td>
                      <td align="right" bgcolor="#475387"><font color="#FFFFFF"><strong>Last updated by</strong></font></td>
                    </tr>
                    
<% 
sj_totalqty = 0
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
                      <td><strong><font color="#000000"><a href="rm_stockAdj_new.asp?sj_no=<%=rs("sj_no")%>"><%=rs("sj_no")%></a></font></strong><br><%=rs("sj_status")%></td>
                      <td><%=rs("sj_remark")%></td>
                      <td><%=chkdate(rs("sj_date"))%></td>
                      <td><%=rs("sj_fromwarehouse")%></td>
                      <td align="right"><%=rs("sj_totalqty")%></td>
                      <td align="right"><%=rs("sj_totalaAmt")%></td>
                      <td align="right"><%=rs("sj_approvedby")%><br />
                      <%=chkdatetime(rs("sj_approveddate"))%></td>
                    </tr>
                    
<%
sj_totalqty = sj_totalqty + rs("sj_totalqty")
sj_totalaAmt = sj_totalaAmt + rs("sj_totalaAmt")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
                    <tr>
                      <td height="30" colspan="5" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td height="30" align="right" bgcolor="#CCCCCC"><strong><%=sj_totalqty%></strong></td>
                      <td height="30" align="right" bgcolor="#CCCCCC"><%=sj_totalaAmt%></td>
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
					Response.Write " <a href='rm_stockAdj_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockAdj_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
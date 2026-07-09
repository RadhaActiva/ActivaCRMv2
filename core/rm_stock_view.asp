<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
md_status = request("md_status")

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(date())
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if ordertype = "" then 
   ordertype = "desc"
end if

i = 1

sql = "Select sum(tblstocktran.stk_qty) as totalqty, round(sum(tblstocktran.stk_qty*tblmodel.md_averageecost),2) as totalvalue " & _
		"from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
		"where tblmodel.md_code is not null and tblstocktran.stk_date < '" & ChkDateYYYYMMDD(DateAdd("d",1,job_date_from)) & "'" 
		
if md_status <> "" then 
   sql = sql & " and tblmodel.md_status = '" & md_status & "' "
end if

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if
response.Cookies("GAPS")("sqlexcel2") = sql
'response.write sql & "<br><br>"
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   t_totalqty = rs("totalqty")
   t_totalvalue = rs("totalvalue")
end if
rs.close


sql = "SELECT     md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_group_type, md_unitprice, md_brands, md_rcpprice, " & _
		"(SELECT     SUM(stk_qty) AS totalqty " & _
		"FROM          tblstocktran " & _
		"WHERE      (stk_itm_code = tblmodel.md_code) and tblstocktran.stk_date < '" & ChkDateYYYYMMDD(DateAdd("d",1,job_date_from)) & "') AS totalqty " & _
		"FROM         tblmodel " & _
		"WHERE     (md_code IS NOT NULL) " 
		
if md_status <> "" then 
   sql = sql & " and tblmodel.md_status = '" & md_status & "' "
end if

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by md_code desc"
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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&orderby=" & request("orderby") & "&md_status=" & request("md_status") & "&ordertype=" & request("ordertype")

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Stock Item Master</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">
                    <form name="form1" id="form1" method="post" action="rm_stock_view.asp?type=reset">
                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                       <tr>
                        <td class="titlegrey1">Stock Up to</td>
                        <td><input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                         <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </td>
                        </tr>
                      
                        <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td width="84%"><select name="searchitem" id="searchitem">
                          <option value="tblmodel.md_code" <% if searchitem = "tblmodel.md_code" then response.write " selected" %>>Stock Code</option>
                          <option value="tblmodel.md_desc" <% if searchitem = "tblmodel.md_desc" then response.write " selected" %>>Stock Desc</option>
                          <option value="tblmodel.md_category" <% if searchitem = "tblmodel.md_category" then response.write " selected" %>>Category</option>
                          <option value="tblmodel.md_group_type" <% if searchitem = "tblmodel.md_group_type" then response.write " selected" %>>Group Type</option>
                          <option value="tblmodel.md_model" <% if searchitem = "tblmodel.md_model" then response.write " selected" %>>Model</option>
                          <option value="tblmodel.md_type" <% if searchitem = "tblmodel.md_type" then response.write " selected" %>>Type</option>
                          <option value="tblmodel.md_brands" <% if searchitem = "tblmodel.md_brands" then response.write " selected" %>>Brand</option>
                        </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                          <option value="totalqty" <% if orderby = "totalqty" then response.write " selected" %>>Current Qty</option>
                          <option value="tblmodel.md_code" <% if orderby = "tblmodel.md_code" then response.write " selected" %>>Stock Code</option>
                          <option value="tblmodel.md_desc" <% if orderby = "tblmodel.md_desc" then response.write " selected" %>>Stock Desc</option>
                          <option value="tblmodel.md_category" <% if orderby = "tblmodel.md_category" then response.write " selected" %>>Category</option>
                          <option value="tblmodel.md_model" <% if orderby = "tblmodel.md_model" then response.write " selected" %>>Model</option>
                          <option value="tblmodel.md_type" <% if orderby = "tblmodel.md_type" then response.write " selected" %>>Type</option>
                          <option value="tblmodel.md_brands" <% if orderby = "tblmodel.md_brands" then response.write " selected" %>>Brand</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="button" id="button3" value="Submit" /></td>
                      </tr>
                      <tr>
                        <td class="titlegrey1">Status</td>
                        <td>
                          <select name="md_status" id="md_status">
                            <option value="">All</option>
                            <option value="Y" <%if md_status="Y" then response.write " selected"%>>Y</option>
                            <option value="N" <%if md_status="N" then response.write " selected"%>>N</option>
                        </select></td>
                      </tr>
                    </table>
                  </form></td>
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
					Response.Write " <a href='rm_stock_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stock_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <th align="center" bgcolor="#666666" width="5%"><font color="#FFFFFF"><strong><span>No</span></strong></font></th>
                      <th bgcolor="#666666" width="10%"><font color="#FFFFFF"><strong><span> Stock Code</span></strong></font></th>
                      <th bgcolor="#666666" width="30%"><font color="#FFFFFF"><strong>Description</strong></font></th>
                      <th align="center" bgcolor="#666666" width="15%"><font color="#FFFFFF"><strong><span>Category</span></strong></font></th>
                      <!--<td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Group Type</strong></font></td>-->
                      <th align="center" width="5%" bgcolor="#666666"><font color="#FFFFFF"><strong><span>Brand</span></strong></font></th>
                      <th align="center" width="5%" bgcolor="#666666" ><font color="#FFFFFF"><strong><span>Type</span></strong></font></th>
                      <th align="center" width="5%" bgcolor="#666666"><font color="#FFFFFF"><strong><span>Status</span></strong></font></th>
                      <th align="right" width="8%" bgcolor="#666666"><font color="#FFFFFF"><strong><span>RCP, RM</span></strong></font></th>
                      <th align="center" width="8%" bgcolor="#666666"><font color="#FFFFFF"><strong>Current <br /> Qty</strong></font></th>
                    </tr>
                   
                    
<% 
totalqty=0
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
                     <td height="40" align="center"><%=j%> </td>
                      <td align="left"><a href="rm_stock_new.asp?md_code=<%=rs("md_code")%>"><strong><%=rs("md_code")%></strong></a></td>
                      <td><%=rs("md_desc")%> </td>
                      <td align="center"><%=rs("md_category")%></td>
                      <!--<td align="center" nowrap="nowrap"><%=rs("md_group_type")%> </td>-->
                      <td align="center"><%=rs("md_brands")%></td>
                      <td align="center"><%=rs("md_type")%></td>
                      <td align="center"><%=rs("md_status")%></td>
                      <td align="right"><%=chknumber2(rs("md_rcpprice"))%></td>
                      <td align="center"><strong><%=chknumber(rs("totalqty"))%></strong></td>
                 </tr>
<%
totalqty=totalqty+chknumber(rs("totalqty"))
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
                 
                  <tr>
                      <td height="40" colspan="8" align="right"><strong>Total</strong></td>
                      <td align="center"><strong><%=totalqty%></strong></td>
                    </tr>
                     <tr>
                      <td height="40" colspan="8" align="right"><strong>Grand Total</strong></td>
                      <td align="center"><strong><%=t_totalqty%></strong></td>
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
					Response.Write " <a href='rm_stock_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stock_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
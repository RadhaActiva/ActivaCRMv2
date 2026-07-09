<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
md_status = request("md_status")

if ordertype = "" then 
   ordertype = "desc"
end if

i = 1
sql = "SELECT     md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_group_type, md_unitprice, md_brands, md_rcpprice, " & _
		"(SELECT     SUM(stk_qty) AS totalqty " & _
		"FROM          tblstocktran " & _
		"WHERE      (stk_itm_code = tblmodel.md_code) and stk_reference ='W1') AS totalqty " & _
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
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="sales_stock_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      
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
					Response.Write " <a href='sales_stock_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='sales_stock_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Stock Code</span></strong></font></td>
                      <td bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock Name </strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Category</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Group Type</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Brand</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Type</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Status</span></strong></font></td>
                      <td align="right" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>RCP, RM</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Current Qty</strong></font></td>
                    </tr>
                    
<% 
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
                     <td align="left" nowrap="nowrap"><a href="sales_stock_new.asp?md_code=<%=rs("md_code")%>"><strong><%=rs("md_code")%></strong></a></td>
                     <td nowrap="nowrap"><%=rs("md_desc")%> </td>
                     <td align="center"><%=rs("md_category")%></td>
                     <td align="center" nowrap="nowrap"><%=rs("md_group_type")%> </td>
                     <td align="center"><%=rs("md_brands")%></td>
                     <td align="center"><%=rs("md_type")%></td>
                     <td align="center"><%=rs("md_status")%></td>
                     <td align="right"><%=rs("md_rcpprice")%></td>
                     <td align="center"><strong><%=chknumber(rs("totalqty"))%></strong></td>
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
					Response.Write " <a href='sales_stock_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='sales_stock_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
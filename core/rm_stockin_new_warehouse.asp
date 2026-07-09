<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
formname=request("formname") 
fieldname=request("fieldname")
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
i = 1
sql = "SELECT wh_id, wh_code, wh_name, wh_address, wh_postcode, wh_state_id, wh_state, wh_city_id, wh_city, wh_tel, wh_fax, wh_remark, " & _
      "wh_contact_person, wh_email, wh_status " & _
	  "FROM tblwarehouse  where wh_id is not null  and wh_status='Y'" 

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

sql = sql & " order by wh_code"

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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&formname=" & formname & "&fieldname=" & fieldname & "&fieldname1=" & fieldname1

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Store List</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_stockin_new_warehouse.asp">
      <select name="searchitem">                   
		  <option value="tblwarehouse.wh_code" <%if searchitem = "tblwarehouse.md_code" then response.write " selected"%>>Store Code</option>
		  <option value="tblwarehouse.wh_name" <%if searchitem = "tblwarehouse.wh_name" then response.write " selected"%>>Store Name</option>
          <option value="tblwarehouse.wh_tel" <%if searchitem = "tblwarehouse.wh_tel" then response.write " selected"%>>Store Tel</option>
        </select>
        <input type="text" name="searchvalue" value="<%=searchvalue%>">
        <input type="hidden" name="formname" value="<%=formname%>">
        <input type="hidden" name="fieldname" value="<%=fieldname%>">
        <input type="submit" name="Submit" value="Submit">
      </form></td>
  </tr>
  <tr> 
    <td align="right" valign="top"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%>
      </font>of <font color="3366ff"> 
      <%=pgCount%>
      </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_stockin_new_warehouse.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockin_new_warehouse.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>Code</strong></td>
          <td width="12%"><strong>Name</strong></td>
          <td><strong>Tel</strong></td>
          <td width="14%" class='tktTotals'><strong>Person In Charge</strong></td>
          <td width="14%" class='tktTotals'><strong>Remark</strong></td>
        </tr>
        <% 
Set rs1 = Server.CreateObject("ADODB.Recordset")

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
        <tr valign="top" bgcolor="<%=nbgcolor%>"> 
          <td nowrap><%=count%>.</td>
          <td><a href="javascript:parent.opener.document.forms['<%=formname%>'].<%=fieldname%>.value='<%=rs("wh_code")%>';parent.window.close();"><%=rs("wh_code")%></a></td>
          <td><%=rs("wh_name")%></td>
          <td><%=rs("wh_tel")%></td>
          <td><%=rs("wh_contact_person")%></td>
          <td><%=rs("wh_remark")%></td>
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
  <tr valign="top"> 
    <td colspan="9" align="right"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_stockin_new_warehouse.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockin_new_warehouse.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

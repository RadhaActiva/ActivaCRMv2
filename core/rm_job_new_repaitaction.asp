<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
ra_type=request("ra_type") 
formname=request("formname") 
fieldname=request("fieldname")
fieldname1=request("fieldname1")
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
i = 1
sql = "SELECT ra_id, ra_type, ra_category, ra_repairaction FROM tbljob_repairaction where ra_id is not null and " & searchitem & " like '%" & searchvalue& "%' "

if request("ra_type") <> "All" then 
sql = sql & " and ra_type='" & request("ra_type") & "' "
end if

sql = sql & " order by ra_category"

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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&ra_type=" & ra_type & "&formname=" & formname & "&fieldname=" & fieldname & "&fieldname1=" & fieldname1

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Repair Action</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_job_new_repaitaction.asp">
<strong>Model Type</strong>
<select name="ra_type" id="ra_type">
          <option value="All" <%if ra_type = "All" then response.write " selected"%>>All</option>
          <option value="WH" <%if ra_type = "WH" then response.write " selected"%>>WH</option>
          <option value="CF" <%if ra_type = "CF" then response.write " selected"%>>CF</option>
        </select>
        <br>
        <select name="searchitem">                   
		  <option value="tbljob_repairaction.ra_category" <%if searchitem = "tbljob_repairaction.ra_category" then response.write " selected"%>>Category</option>
		  <option value="tbljob_repairaction.ra_repairaction" <%if searchitem = "tbljob_repairaction.md_desc" then response.write " selected"%>>Repair Action</option>
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
					Response.Write " <a href='rm_job_new_repaitaction.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_new_repaitaction.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>Repair Category</strong></td>
          <td width="12%"><strong>Repair Description</strong></td>
          <td><strong>Model Type</strong></td>
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
          <td><%=rs("ra_category")%></td>
          <td><a href="javascript:parent.opener.document.forms['formorder2'].job_tech_faulty_action.value='<%=chkstring(rs("ra_repairaction"))%>';parent.window.close();"><%=rs("ra_repairaction")%></a></td>
          <td><%=rs("ra_type")%></td>
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
					Response.Write " <a href='rm_job_new_repaitaction.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_new_repaitaction.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

<!-- #include file="database/datastore.asp" -->

<%
stk_itm_code = request("stk_itm_code")
stk_reference = request("stk_reference")
jobmonth=request("jobmonth") 
jobyear=request("jobyear") 
stype=request("stype") 
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
i = 1
sql = "SELECT stk_id, stk_voucherno, stk_reference, stk_date, stk_type, stk_itm_code, stk_fromwarehouse, stk_towarehouse, stk_desc, stk_qty, " & _
      "stk_balanceqty, stk_sales_price, stk_logby, stk_logdate " & _
      "FROM tblstocktran where stk_itm_code = '" & stk_itm_code & "' and stk_reference='" & stk_reference & "' " & _
	  "and month(stk_date)=" & jobmonth & " and year(stk_date)=" & jobyear & " "
'response.write sql
if stype = "In" then 
sql = sql & "and stk_qty > 0 "
else
sql = sql & "and stk_qty < 0 "
end if

sql = sql & " order by stk_id desc"

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
link = "&stk_itm_code=" & stk_itm_code & "&stk_reference=" & stk_reference & "&jobmonth=" & jobmonth & "&jobyear=" & jobyear 
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Stock Movement Detail</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_rpt_inventory_shockin_detail.asp">
<strong>Store</strong>
<select name="stk_reference" id="stk_reference">
  <option value="0"></option>
  <%			
				sql1 = "SELECT wh_code, wh_name FROM tblwarehouse order by wh_code"	
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if (stk_reference) = (rs1("wh_code")) then
					  response.write "<option value='" & rs1("wh_code") & "' selected>" & rs1("wh_code") & "</option>"
					  else
					  response.write "<option value='" & rs1("wh_code") & "'>" & rs1("wh_code") & "</option>"
					  end if 					  
				rs1.movenext
				wend
				rs1.close					
				%>
</select>
<select name="jobyear" id="jobyear">
  <option value="2016"<%if jobyear="2016" then response.write " selected"%>>2016</option>
</select>
<select name="jobmonth" id="jobmonth">
  <option value="1" <%if jobmonth="1" then response.write " selected"%>>Jan</option>
  <option value="2" <%if jobmonth="2" then response.write " selected"%>>Feb</option>
  <option value="3" <%if jobmonth="3" then response.write " selected"%>>Mar</option>
  <option value="4" <%if jobmonth="4" then response.write " selected"%>>Apr</option>
  <option value="5" <%if jobmonth="5" then response.write " selected"%>>May</option>
  <option value="6" <%if jobmonth="6" then response.write " selected"%>>Jun</option>
  <option value="7" <%if jobmonth="7" then response.write " selected"%>>Jul</option>
  <option value="8" <%if jobmonth="8" then response.write " selected"%>>Aug</option>
  <option value="9" <%if jobmonth="9" then response.write " selected"%>>Sep</option>
  <option value="10" <%if jobmonth="10" then response.write " selected"%>>Oct</option>
  <option value="11" <%if jobmonth="11" then response.write " selected"%>>Nov</option>
  <option value="12" <%if jobmonth="12" then response.write " selected"%>>Dec</option>
</select>
<br>
<strong>Item Code</strong>
<input type="text" name="stk_itm_code" value="<%=stk_itm_code%>">
<br>
<strong>Type
<select name="stype" id="stype">
  <option value="In" <%if stype="In" then response.write " selected"%>>In</option>
  <option value="Out" <%if stype="Out" then response.write " selected"%>>Out</option>
</select>
</strong>
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
					Response.Write " <a href='rm_rpt_inventory_shockin_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_shockin_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td><strong>Doc No</strong></td>
          <td><strong>Doc Type</strong></td>
          <td><strong>Doc Date</strong></td>
          <td align="center" class='tktTotals'><strong>Qty</strong></td>
          <td class='tktTotals'><strong>Item Desc</strong></td>
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
        
        <tr valign="top" bgcolor="<%=nbgcolor%>"> 
          <td nowrap><%=count%>.</td>
          <td><%=rs("stk_voucherno")%></td>
          <td><%=rs("stk_type")%></td>
          <td><%=chkdatetime(rs("stk_date"))%></td>
          <td align="center"><%=rs("stk_qty")%></td>
          <td><%=rs("stk_desc")%></td>
        </tr>
        <%
stk_qty = stk_qty + rs("stk_qty")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
<tr valign="top" bgcolor="<%=nbgcolor%>">
          <td colspan="4" align="right" nowrap><strong>Total</strong></td>
          <td align="center"><strong><%=stk_qty%></strong></td>
          <td>&nbsp;</td>
        </tr>
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
					Response.Write " <a href='rm_rpt_inventory_shockin_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_shockin_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

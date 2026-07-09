<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
md_type=request("md_type") 
sj_fromwarehouse=request("sj_fromwarehouse") 

%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
'modified by  110922 to exclude qty lesser than 1

i = 1
'sql = "SELECT tblwarehouse_stock.wst_id, tblwarehouse_stock.wst_wh_code, tblwarehouse_stock.wst_itm_code, " & _
'		"tblwarehouse_stock.wst_itm_current_qty, tblwarehouse_stock.wst_itm_min_qty,  " & _
'		"tblwarehouse_stock.wst_itm_remarks, tblwarehouse_stock.wst_lastupdateby, tblwarehouse_stock.wst_lastupdatedate, " & _
'		"tblmodel.md_code, tblmodel.md_category, tblmodel.md_desc, tblmodel.md_unitprice, tblmodel.md_type " & _
'		"FROM tblwarehouse_stock inner join tblmodel on tblwarehouse_stock.wst_itm_code = tblmodel.md_code " & _
'		"where tblwarehouse_stock.wst_id is not null and " & searchitem & " like '%" & searchvalue& "%' and tblwarehouse_stock.wst_wh_code='" & sj_fromwarehouse & "'"

'sql = "select * from ( select tblmodel.md_code, tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type, tblmodel.md_rcpprice,  " & _
'		"sum(tblstocktran.stk_qty) as wst_itm_current_qty " & _
'		"from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
'		"where (tblmodel.md_category = 'Parts' or tblmodel.md_category = 'Component') and tblstocktran.stk_reference <> '0' and " & searchitem & " like '%" & searchvalue& "%' and tblstocktran.stk_reference='" & sj_fromwarehouse & "' "

'modified by  160724 to show stock list from respective tech wh and exclude qty lesser than 1

sql2 = "select tblwarehouse_stock.wst_itm_code, tblmodel.md_code,tblmodel.md_rcpprice,tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type,tblwarehouse_stock.wst_itm_current_qty, md_averageecost " & _
	   "from tblwarehouse_stock inner join tblmodel on " & _
	   "tblwarehouse_stock.wst_itm_code=tblmodel.md_code where tblmodel.md_category = 'Parts' and " & searchitem & " like '%" & searchvalue& "%' and tblwarehouse_stock.wst_wh_code = '" &  sj_fromwarehouse & "' " & _
	   "and tblwarehouse_stock.wst_itm_current_qty > 0 "
	
if request("md_type") <> "All" then 
    sql2 = sql2 & " and tblmodel.md_type='" & request("md_type") & "' "
end if

sql2 = sql2 & " order by md_code"

set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&formname=" & formname & "&fieldname=" & fieldname & "&sj_fromwarehouse=" & sj_fromwarehouse

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Stock List for Store: <%=sj_fromwarehouse%></strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_stockAdj_new_item.asp">
<strong>Model Type</strong>
<select name="md_type" id="md_type">
          <option value="All" <%if md_type = "All" then response.write " selected"%>>All</option>
          <option value="WH" <%if md_type = "WH" then response.write " selected"%>>WH</option>
          <option value="CF" <%if md_type = "CF" then response.write " selected"%>>CF</option>
        </select>
        <br>
        <select name="searchitem">                   
		  <option value="tblmodel.md_code" <%if searchitem = "tblmodel.md_code" then response.write " selected"%>>Model Code</option>
		  <option value="tblmodel.md_desc" <%if searchitem = "tblmodel.md_desc" then response.write " selected"%>>Description</option>
          <option value="tblmodel.md_category" <%if searchitem = "tblmodel.md_category" then response.write " selected"%>>Category</option>
          <option value="tblmodel.md_barcode" <%if searchitem = "tblmodel.md_barcode" then response.write " selected"%>>Barcode</option>
        </select>
        <input type="text" name="searchvalue" value="<%=searchvalue%>">
        <input type="hidden" name="formname" value="<%=formname%>">
        <input type="hidden" name="fieldname" value="<%=fieldname%>">
        <input type="hidden" name="sj_fromwarehouse" value="<%=sj_fromwarehouse%>">
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
					Response.Write " <a href='rm_stockAdj_new_item.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockAdj_new_item.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>Code</strong></td>
          <td width="12%"><strong>Description</strong></td>
          <td><strong>category</strong></td>
          <td width="14%" class='tktTotals'><strong>Type</strong></td>
          <td class='tktTotals'><strong>Current Qty</strong></td>
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
          <td><a href="javascript:parent.opener.document.forms['formdodetail'].sjd_itm_code.value='<%=rs("md_code")%>';parent.opener.document.forms['formdodetail'].sjd_itm_desc.value='<%=ChkString(rs("md_desc"))%>';parent.opener.document.forms['formdodetail'].sjd_current_qty.value='<%=ChkString(rs("wst_itm_current_qty"))%>';parent.window.close();"><%=rs("md_code")%></a></td>
          <td><%=rs("md_desc")%></td>
          <td><%=rs("md_category")%></td>
          <td><%=rs("md_type")%></td>
          <td><%=rs("wst_itm_current_qty")%></td>
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
					Response.Write " <a href='rm_stockAdj_new_item.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_stockAdj_new_item.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

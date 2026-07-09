<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
md_type=request("md_type") 
formname=request("formname") 
fieldname=request("fieldname")
fieldname1=request("fieldname1")
tech_wh=request("tech_wh")   
faulty=request("faulty")   
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
  
'modified by  110922 to exclude qty lesser than 1
i = 1	
'sql2 = "select * from (select tblmodel.md_code, tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type, tblmodel.md_rcpprice,  " & _
'		"sum(tblstocktran.stk_qty) as totalqty " & _
'		"from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
'		"where (tblmodel.md_category = 'Parts' or tblmodel.md_category = 'Labour') and tblstocktran.stk_reference <> '0'  " 

'modified by  160724 to show stock list from respective tech wh and exclude qty lesser than 1 exception applies to labout/transport

'if request("md_category") = "Service" then
 '       sql2="select 0 as wst_itm_current_qty,md_code,md_rcpprice, md_model,md_desc, md_category,md_type,1 as current_qty,1 as avg_cost from tblmodel "  & _
  '      "where md_code in ('Labour','Service')"
'else

    if faulty <> "" then
        sql2 = "select tblmodel.md_code,tblmodel.md_model,tblmodel.md_rcpprice,tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type,'1' as wst_itm_current_qty " & _
	           "from tblmodel where tblmodel.md_category = 'Labour'"      
    else
    
       ' sql2 = "select tblwarehouse_stock.wst_itm_code, tblmodel.md_code,tblmodel.md_model,tblmodel.md_rcpprice,tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type,tblwarehouse_stock.wst_itm_current_qty, md_averageecost " & _
	    '       "from tblwarehouse_stock inner join tblmodel on " & _
	     '      "tblwarehouse_stock.wst_itm_code=tblmodel.md_code where tblmodel.md_category = 'Parts' and " & searchitem & " like '%" & searchvalue& "%' and tblwarehouse_stock.wst_wh_code = '" &  tech_wh & "' " & _
	      '     "and tblwarehouse_stock.wst_itm_current_qty > 0 "
        if request("md_type") = "Service" then
                sql2 = "select tblmodel.md_code,tblmodel.md_model,tblmodel.md_rcpprice,tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type,'1' as wst_itm_current_qty " & _
	           "from tblmodel where tblmodel.md_category = 'Service'"      
        else    
                sql2 = "select tblwarehouse_stock.wst_itm_code, tblmodel.md_code,tblmodel.md_model,tblmodel.md_rcpprice,tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type,tblwarehouse_stock.wst_itm_current_qty, md_averageecost " & _
	           "from tblwarehouse_stock inner join tblmodel on " & _
	           "tblwarehouse_stock.wst_itm_code=tblmodel.md_code where tblmodel.md_category in ('Parts','Labour') and " & searchitem & " like '%" & searchvalue& "%' and tblwarehouse_stock.wst_wh_code = '" &  tech_wh & "' " & _
	           "and tblwarehouse_stock.wst_itm_current_qty > 0 "
        end if

        if request("md_type") <> "All" then 
            sql2 = sql2 & " and tblmodel.md_type='" & request("md_type") & "' "
        end if

        sql2 = sql2 & " order by md_code"
    end if
    
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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&tech_wh=" & tech_wh & "&formname=" & formname & "&fieldname=" & fieldname & "&fieldname1=" & fieldname1
'link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&formname=" & formname & "&fieldname=" & fieldname & "&sf_fromwarehouse=" & sf_fromwarehouse

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Item / Model List</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_job_new_spareparts.asp?tech_wh=<%=tech_wh%>">
<strong>Model Type</strong>
<select name="md_type" id="md_type">
          <option value="All" <%if md_type = "All" then response.write " selected"%>>All</option>
          <option value="WH" <%if md_type = "WH" then response.write " selected"%>>WH</option>
          <option value="CF" <%if md_type = "CF" then response.write " selected"%>>CF</option>
          <option value="Service" <%if md_type = "Services" then response.write " selected"%>>Services</option>
        </select>
        <br>
        <select name="searchitem">                   
		  <option value="tblmodel.md_code" <%if searchitem = "tblmodel.md_code" then response.write " selected"%>>Model Code</option>
		  <option value="tblmodel.md_desc" <%if searchitem = "tblmodel.md_desc" then response.write " selected"%>>Description</option>
          <option value="tblmodel.md_category" <%if searchitem = "tblmodel.md_category" then response.write " selected"%>>Category</option>
        </select>
        <input type="text" name="searchvalue" value="<%=searchvalue%>">
        <input type="hidden" name="formname" value="<%=formname%>">
        <input type="hidden" name="fieldname" value="<%=fieldname%>">
        <input type="hidden" name="fieldname1" value="<%=fieldname1%>">
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
					Response.Write " <a href='rm_job_new_spareparts.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_new_spareparts.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>Item Code</strong></td>
          <td width="12%"><strong>Model Description</strong></td>
          <td><strong>Category</strong></td>
          <td width="14%" class='tktTotals'><strong>Type</strong></td>
          <td width="7%" class='tktTotals'><strong>RCP Price</strong></td>
          <td width="7%" align="center" class='tktTotals'><strong>Current Qty</strong></td>
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
          <td>
          <%if rs("wst_itm_current_qty") > 0 or rs("md_code")="Labour" or rs("md_code")="Service" then %>
          <a href="javascript:parent.opener.document.forms['<%=formname%>'].<%=fieldname%>.value='<%=rs("md_code")%>';parent.opener.document.forms['<%=formname%>'].<%=fieldname1%>.value='<%=ChkString(rs("md_desc"))%>';parent.opener.document.forms['<%=formname%>'].jobp_unitcost.value='<%=ChkNumber2(rs("md_rcpprice"))%>';parent.opener.document.forms['<%=formname%>'].jobp_qty.value='1';parent.opener.document.forms['<%=formname%>'].jobp_subtotal.value='<%=ChkNumber2(rs("md_rcpprice"))%>';parent.window.close();"><%=rs("md_code")%></a>
          <%else%>
          <%=rs("md_code")%>
          <%end if %></td>
          <td><%=rs("md_desc")%></td>
          <td><%=rs("md_category")%></td>
          <td><%=rs("md_type")%></td>
          <td><%=ChkNumber2(rs("md_rcpprice"))%></td>
          <%if rs("md_model")  = "Service" then%>
             <td align="center"><%=1%></td>
       `` <%else%>
             <td align="center"><%=rs("wst_itm_current_qty")%></td>
          <%end if %>
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
					Response.Write " <a href='rm_job_new_spareparts.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_new_spareparts.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
md_type=request("md_type") 
formname=request("formname") 
fieldname=request("fieldname")
inv_tech_code=request("inv_tech_code")

if request("md_type") <> "Service" then
    sql = "select tech_type from tbltechnician where tech_code='" & inv_tech_code & "'"
    tech_type = selectid(sql)

    if job_code <> "" then ''useful for walk-in when there's no job_code 23/09/24
        sql = "select wh_code from tblwarehouse where wh_contact_person='" & inv_tech_code & "'"
        wh_code = selectid(sql)
    else
        wh_code = "W1"
    end if
end if

if request("inv_tech_code") = "walk-in" then 
   pricelable = "RCP"
else
   if tech_type="IHC" then 
      pricelable = "Price 2"
   elseif tech_type="TPC" or tech_type="IHT" or tech_type="IC" then 
      pricelable = "Price 1"
   else  
      pricelable = "RCP"
   end if	  
end if
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
'modified by  110922 to exclude qty lesser than 1
i = 1
if request("md_type") = "Service" then
   sql = "select tblmodel.md_code,tblmodel.md_category,tblmodel.md_model,tblmodel.md_rcpprice,tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type,'1' as totalqty " & _
	      "from tblmodel where tblmodel.md_category = 'Service'"      
else
    sql = "select * from (select tblmodel.md_code, tblmodel.md_desc, tblmodel.md_category, tblmodel.md_type, tblmodel.md_rcpprice, tblmodel.md_unitprice1, tblmodel.md_unitprice2, " & _
	      "tblwarehouse_stock.wst_itm_current_qty as totalqty " & _
	      "from tblwarehouse_stock inner join tblmodel on tblwarehouse_stock.wst_itm_code=tblmodel.md_code " & _
	      "where (tblmodel.md_category = 'Parts' or tblmodel.md_category = 'Labour' ) and tblwarehouse_stock.wst_wh_code <> '0' and " & searchitem & " like '%" & searchvalue& "%' " 
			  
    if request("md_type") <> "All" and request("md_type") <> "" then 
    sql = sql & " and tblmodel.md_type='" & request("md_type") & "' "
    end if

    sql = sql & " and tblwarehouse_stock.wst_wh_code='" & wh_code & "')t"
    sql = sql & " where t.totalqty > 0 or t.md_code in ('Labour','Service') order by md_code"
end if

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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&formname=" & formname & "&fieldname=" & fieldname 

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Item List</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_invoice_new_spareparts.asp?formname=<%=formname%>&fieldname=<%=fieldname%>">
<strong>Model Type</strong>
<select name="md_type" id="md_type">
          <option value="All" <%if md_type = "All" then response.write " selected"%>>All</option>
          <option value="WH" <%if md_type = "WH" then response.write " selected"%>>WH</option>
          <option value="CF" <%if md_type = "CF" then response.write " selected"%>>CF</option>
          <option value="Service" <%if md_type = "Service" then response.write " selected"%>>Services</option>
        </select>
        <br>
        <strong>Store 
        <select name="inv_tech_code" id="inv_tech_code">
          <option value=""></option>
          <%			
				sql1 = "SELECT tech_id, tech_code, tech_name FROM tbltechnician where tech_status = 'Y' "	
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if (inv_tech_code) = (rs1("tech_code")) then
					  response.write "<option value='" & rs1("tech_code") & "' selected>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs1("tech_code") & "'>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  end if 					  
				rs1.movenext
				wend
				rs1.close					
				%>
        </select>
        </strong><br>
        <select name="searchitem">                   
		  <option value="tblmodel.md_code" <%if searchitem = "tblmodel.md_code" then response.write " selected"%>>Item Code</option>
		  <option value="tblmodel.md_desc" <%if searchitem = "tblmodel.md_desc" then response.write " selected"%>>Description</option>
          <option value="tblmodel.md_barcode" <%if searchitem = "tblmodel.md_barcode" then response.write " selected"%>>Barcode</option>
        </select>
        <input type="text" name="searchvalue" value="<%=searchvalue%>">
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
					Response.Write " <a href='rm_invoice_new_spareparts.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_invoice_new_spareparts.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>Code</strong></td>
          <td><strong>Description</strong></td>
          <td align="right" nowrap><strong><%=pricelable%></strong></td>
          <td width="7%" align="center" class='tktTotals'><strong>Type
          </strong></td>
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
          <td><%=count%>.</td>
          <td nowrap>
          <%if rs("totalqty") > 0 or rs("md_code")="Labour" or rs("md_category")="Service" then %>
			  <%if request("inv_tech_code") = "walk-in" then %>
              <a href="javascript:parent.opener.document.forms['forminvoicedetail'].invd_partcode.value='<%=rs("md_code")%>';parent.opener.document.forms['forminvoicedetail'].invd_desc.value='<%=chkstring(rs("md_desc"))%>';parent.opener.document.forms['forminvoicedetail'].invd_unitcost.value='<%=rs("md_rcpprice")%>';parent.opener.document.forms['forminvoicedetail'].invd_qty.value='1';parent.opener.document.forms['forminvoicedetail'].invd_subtotal.value='<%=rs("md_rcpprice")%>';parent.window.close();"><%=rs("md_code")%></a>
              <%else%>
                  <%if tech_type="IHC" then %>
                        <a href="javascript:parent.opener.document.forms['forminvoicedetail'].invd_partcode.value='<%=rs("md_code")%>';parent.opener.document.forms['forminvoicedetail'].invd_desc.value='<%=chkstring(rs("md_desc"))%>';parent.opener.document.forms['forminvoicedetail'].invd_unitcost.value='<%=rs("md_rcpprice")%>';parent.opener.document.forms['forminvoicedetail'].invd_qty.value='1';parent.opener.document.forms['forminvoicedetail'].invd_subtotal.value='<%=rs("md_rcpprice")%>';parent.window.close();"><%=rs("md_code")%></a>
                  <%elseif tech_type="TPC" or tech_type="IHT" or tech_type="IC" then %> 
                        <a href="javascript:parent.opener.document.forms['forminvoicedetail'].invd_partcode.value='<%=rs("md_code")%>';parent.opener.document.forms['forminvoicedetail'].invd_desc.value='<%=chkstring(rs("md_desc"))%>';parent.opener.document.forms['forminvoicedetail'].invd_unitcost.value='<%=rs("md_rcpprice")%>';parent.opener.document.forms['forminvoicedetail'].invd_qty.value='1';parent.opener.document.forms['forminvoicedetail'].invd_subtotal.value='<%=rs("md_rcpprice")%>';parent.window.close();"><%=rs("md_code")%></a>
                  <%else%>
                        <a href="javascript:parent.opener.document.forms['forminvoicedetail'].invd_partcode.value='<%=rs("md_code")%>';parent.opener.document.forms['forminvoicedetail'].invd_desc.value='<%=chkstring(rs("md_desc"))%>';parent.opener.document.forms['forminvoicedetail'].invd_unitcost.value='<%=rs("md_rcpprice")%>';parent.opener.document.forms['forminvoicedetail'].invd_qty.value='1';parent.opener.document.forms['forminvoicedetail'].invd_subtotal.value='<%=rs("md_rcpprice")%>';parent.window.close();"><%=rs("md_code")%></a>
                  <%end if%>     
              <%end if%>
          <%else%>
          <%=rs("md_code")%>
          <%end if %>    
          </td>
          <td><%=rs("md_desc")%></td>
          <td align="right">
		  <%if request("inv_tech_code") = "walk-in" then %>
		  <%=chknumber2(rs("md_rcpprice"))%>
          <%else%>
                <%
                if tech_type="IHC" then 
                   response.write chknumber2(rs("md_unitprice2"))
                elseif tech_type="TPC" or tech_type="IHT" or tech_type="IC" then 
                   response.write chknumber2(rs("md_unitprice1"))
				else
				    response.write chknumber2(rs("md_rcpprice"))
                end if
				%>
          <%end if%>
          </td>
          <td align="center"><%=rs("md_type")%></td>
          <td align="center"><%=rs("totalqty")%></td>
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
					Response.Write " <a href='rm_invoice_new_spareparts.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_invoice_new_spareparts.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")

if request("orderby") <> "" then 
   orderby = request("orderby")
else
   orderby = "rpi_total_in desc, rpi_total_out desc"   
end if

if request("jobmonth") <> "" then
   jobmonth = request("jobmonth")
else
   jobmonth = month(date())
end if

if request("jobyear") <> "" then
   jobyear = request("jobyear")
else
   jobyear = year(date())
end if

sql2 = "SELECT rpi_id, rpi_month, rpi_year, rpi_item_code, rpi_item_name, rpi_tech01_in, rpi_tech01_out, rpi_tech02_in, rpi_tech02_out, " & _
	"rpi_tech03_in, rpi_tech03_out, rpi_tech04_in, rpi_tech04_out, rpi_tech05_in, rpi_tech05_out, rpi_tech06_in, rpi_tech06_out, rpi_tech07_in, rpi_tech07_out, rpi_tech08_in, rpi_tech08_out,  " & _
	"rpi_tech09_in, rpi_tech09_out, rpi_tech10_in, rpi_tech10_out, rpi_tech11_in, rpi_tech11_out, rpi_tech12_in, rpi_tech12_out, rpi_tech13_in, rpi_tech13_out,  " & _
	"rpi_tech14_in, rpi_tech14_out, rpi_tech15_in, rpi_tech15_out, rpi_tech16_in, rpi_tech16_out, rpi_tech17_in, rpi_tech17_out, rpi_tech18_in, rpi_tech18_out,  " & _
	"rpi_tech19_in, rpi_tech19_out, rpi_tech20_in, rpi_tech20_out, rpi_tech21_in, rpi_tech21_out, rpi_tech22_in, rpi_tech22_out, rpi_tech23_in, rpi_tech23_out,  " & _
	"rpi_tech24_in, rpi_tech24_out, rpi_tech25_in, rpi_tech25_out, rpi_tech26_in, rpi_tech26_out, rpi_tech27_in, rpi_tech27_out, rpi_tech28_in, rpi_tech28_out,  " & _
	"rpi_tech29_in, rpi_tech29_out, rpi_tech30_in, rpi_tech30_out, rpi_total_in, rpi_total_out " & _
	"FROM tblrpr_techinventory where rpi_month=" & jobmonth & " and rpi_year=" & jobyear & " "
	
if searchvalue <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if orderby <> "" then
sql2 = sql2 & " order by " & orderby & " " & ordertype
else
sql2 = sql2 & " order by rpi_total_in desc, rpi_total_out desc"
end if

response.write sql
	
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
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&orderby=" & orderby & "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&Searchor_date=" & Searchor_date & "&ordertype=" & ordertype

%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_inventory_shockin.asp";
	document.form1.submit();
}
</script> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Store Inventory Report</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1"> Stock-In and Stock-Out Report                          
                        <label for="select"></label></td>
                      <td width="20%" align="center" class="titlegrey1"><img src="images/excel.jpg" width="57" height="21" /></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="action_report.asp?type=inventorystock">
                <tr>
                  <td width="80%" valign="top" bgcolor="#FFFFFF"><strong>Filtered by</strong>
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
                    </select></td>
                  <td width="20%" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                
                <tr>
                  <td height="30" align="left" bgcolor="#FFFFFF"><select name="searchitem" id="searchitem">
                    <option value="rpi_item_code" <% if searchitem = "rpi_item_code" then response.write " selected" %>>Item Code</option>
                    <option value="rpi_item_name" <% if searchitem = "rpi_item_name" then response.write " selected" %>>Item Desc</option>
                </select>
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                    <select name="orderby" id="orderby">
                      <option value="rpi_item_code" <% if orderby = "rpi_item_code" then response.write " selected" %>>Item Code</option>
                      <option value="rpi_item_name" <% if orderby = "rpi_item_name" then response.write " selected" %>>Item Desc</option>
                      <option value="rpi_total_in" <%if orderby="rpi_total_in" then response.write " selected"%>>Total Stock-In</option>
                      <option value="rpi_total_out" <%if orderby="rpi_total_out" then response.write " selected"%>>Total Stock-Out</option>
                    </select>
                    <select name="ordertype" id="ordertype">
                      <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                      <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                  </select>
                    <span class="titlegrey1">
                    <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" />
                  </span></td>
                  <td height="30" align="left" bgcolor="#FFFFFF"><span class="titlegrey1">
                    <input type="submit" name="button" id="button3" value="Generate Report" />
                  </span></td>
                </tr>
                 </form>
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="22" align="right" class="style1"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_inventory_shockin.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_shockin.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="6%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="20%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item  Code.</span></strong></font></td>
                      <td width="29%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item  Name</span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> W1</span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>W2<span><br />
                      </span></strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>W3</strong></font></td>
                      <td width="6" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>W4</strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>W5</strong></font></td>
                      <td width="6" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>W6</strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>W7</strong></font></td>
                      <td width="5" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>W8</strong></font></td>
                      <td width="4" colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Balance</strong></font></td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>In</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Out</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1">&nbsp;</td>
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
                      <td height="40" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("rpi_item_code")%></font></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"> <%=rs("rpi_item_name")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W1&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech01_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W1&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech01_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W2&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech02_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W2&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech02_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W3&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech03_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W3&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech03_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W4&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech04_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W4&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech04_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W5&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech05_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W5&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech05_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W6&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech06_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W6&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech06_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W7&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech07_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W7&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech07_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W8&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech08_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <a href="javascript:popup('rm_rpt_inventory_shockin_detail.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&stk_reference=W8&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("rpi_tech08_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=rs("rpi_total_in")%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs("rpi_total_out")%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#B9B9FF"><strong><%=rs("rpi_total_in") + rs("rpi_total_out")%></strong></td>
                    </tr>
<%
 
rpi_tech01_in = rpi_tech01_in + rs("rpi_tech01_in")
rpi_tech02_in = rpi_tech02_in + rs("rpi_tech02_in")
rpi_tech03_in = rpi_tech03_in + rs("rpi_tech03_in")
rpi_tech04_in = rpi_tech04_in + rs("rpi_tech04_in")
rpi_tech05_in = rpi_tech05_in + rs("rpi_tech05_in")
rpi_tech06_in = rpi_tech06_in + rs("rpi_tech06_in")
rpi_tech07_in = rpi_tech07_in + rs("rpi_tech07_in")
rpi_tech08_in = rpi_tech08_in + rs("rpi_tech08_in")
rpi_total_in = rpi_total_in + rs("rpi_total_in")

rpi_tech01_out = rpi_tech01_out + rs("rpi_tech01_out")
rpi_tech02_out = rpi_tech02_out + rs("rpi_tech02_out")
rpi_tech03_out = rpi_tech03_out + rs("rpi_tech03_out")
rpi_tech04_out = rpi_tech04_out + rs("rpi_tech04_out")
rpi_tech05_out = rpi_tech05_out + rs("rpi_tech05_out")
rpi_tech06_out = rpi_tech06_out + rs("rpi_tech06_out")
rpi_tech07_out = rpi_tech07_out + rs("rpi_tech07_out")
rpi_tech08_out = rpi_tech08_out + rs("rpi_tech08_out")
rpi_total_out = rpi_total_out + rs("rpi_total_out")

rpi_balance = rpi_balance + (rs("rpi_total_in") + rs("rpi_total_out"))

 
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>                   
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="3" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=rpi_tech01_in%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=rpi_tech01_out%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=rpi_tech02_in%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=rpi_tech02_out%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=rpi_tech03_in%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=rpi_tech03_out%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech04_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech04_out%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech05_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech05_out%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech06_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech06_out%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech07_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech07_out%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech08_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_tech08_out%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_total_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_total_out%></strong></td>
                      <td align="center" bgcolor="#6F6FFF"><strong><%=rpi_balance%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="22" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_inventory_shockin.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_shockin.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
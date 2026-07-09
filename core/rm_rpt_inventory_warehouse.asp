<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
wh_code = request("wh_code")

if request("orderby") <> "" then 
   orderby = request("orderby")
else
   orderby = "rpi_total_in desc, rpi_total_out desc"   
end if

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if


sql2 = "SELECT sum(rpi_total_in) as total_rpi_total_in, " & _
       "sum(rpi_total_out) as total_rpi_total_out FROM tblrpr_techinventory_single " & _
       "where rpi_id is not null "
	   
if searchvalue <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue& "%' "
end if	  

 
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   total_rpi_total_in = rs("total_rpi_total_in")
   total_rpi_total_out = rs("total_rpi_total_out")
end if
rs.close


sql2 = "SELECT rpi_id, rpi_month, rpi_year, rpi_item_code, rpi_item_name, rpi_total_in, rpi_total_out, rpi_total_bal " & _
	"FROM tblrpr_techinventory_single where rpi_id is not null " 
	
if searchvalue <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if orderby <> "" then
sql2 = sql2 & " order by " & orderby & " " & ordertype
else
sql2 = sql2 & " order by rpi_total_bal desc"
end if


Response.Cookies("GAPS")("sqlexcel") = sql2	
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
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&orderby=" & orderby & "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&Searchor_date=" & Searchor_date & "&ordertype=" & ordertype & "&wh_code=" & wh_code & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to

%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_inventory_warehouse.asp";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Store Inventory Report (Single)</div></td>
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
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_inventory_warehouse_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="action_report.asp?type=inventorystock_single">
                <tr>
                  <td width="80%" valign="top" bgcolor="#FFFFFF"><strong>Filtered by</strong><strong><font color="#000000"><strong>
                    <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                    <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                    <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                  <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </td>
                  <td width="20%" valign="top" bgcolor="#FFFFFF"><span class="titlegrey1">
                    <input type="submit" name="Submit" id="button3" value="Generate Report" />
                  </span></td>
                </tr>
                
                
                <tr>
                  <td height="30" align="left" bgcolor="#FFFFFF"><select name="searchitem" id="searchitem">
                    <option value="rpi_item_code" <% if searchitem = "rpi_item_code" then response.write " selected" %>>Item Code</option>
                    <option value="rpi_item_name" <% if searchitem = "rpi_item_name" then response.write " selected" %>>Item Desc</option>
                </select>
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                    <select name="wh_code" id="wh_code">
                      <%			
				sql1 = "SELECT wh_id, wh_code, wh_name, wh_remark FROM tblwarehouse order by wh_code"	
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if (wh_code) = (rs1("wh_code")) then
					  response.write "<option value='" & rs1("wh_code") & "' selected>" & rs1("wh_code") & " - " & rs1("wh_name") & " - " & rs1("wh_remark") & "</option>"
					  else
					  response.write "<option value='" & rs1("wh_code") & "'>" & rs1("wh_code") & " - " & rs1("wh_name") & " - " & rs1("wh_remark") & "</option>"
					  end if 					  
				rs1.movenext
				wend
				rs1.close					
				%>
                    </select>
                    <br />
                    <select name="orderby" id="orderby">
                      <option value="rpi_item_code" <% if orderby = "rpi_item_code" then response.write " selected" %>>Item Code</option>
                      <option value="rpi_item_name" <% if orderby = "rpi_item_name" then response.write " selected" %>>Item Desc</option>
                      <option value="rpi_total_in" <%if orderby="rpi_total_in" then response.write " selected"%>>Total Stock-In</option>
                      <option value="rpi_total_out" <%if orderby="rpi_total_out" then response.write " selected"%>>Total Stock-Out</option>
                      <option value="rpi_total_bal" <%if orderby="rpi_total_bal" then response.write " selected"%>>Total Balance</option>                      
                    </select>
                    <select name="ordertype" id="ordertype">
                      <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                      <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                  </select>
                    <br /></td>
                  <td height="30" align="left" bgcolor="#FFFFFF"><span class="titlegrey1">
                    <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" />
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
					Response.Write " <a href='rm_rpt_inventory_warehouse.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_warehouse.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="50" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="188" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item  Code.</span></strong></font></td>
                      <td colspan="17" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item  Name</span></strong></font></td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                      <td width="191" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Balance</strong></font></td>
                    </tr>
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td colspan="17" align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td width="108" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock-In</strong></font></td>
                      <td width="73" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock-Out</strong></font></td>
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
                      <td height="40" align="center"><%=j%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("rpi_item_code")%></font></strong></td>
                      <td colspan="17" align="left" nowrap="nowrap" bgcolor="#FFFFFF"> <%=rs("rpi_item_name")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
                      <a href="javascript:popup('rm_rpt_inventory_warehouse_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&stk_reference=<%=wh_code%>&stk_itm_code=<%=rs("rpi_item_code")%>&stype=In&wh_code=<%=wh_code%>','cb18','scrollbars=yes,resizable=yes,width=500,height=500')">
                      <%=rs("rpi_total_in")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
                      <a href="javascript:popup('rm_rpt_inventory_warehouse_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&stk_reference=<%=wh_code%>&stk_itm_code=<%=rs("rpi_item_code")%>&stype=Out&wh_code=<%=wh_code%>','cb19','scrollbars=yes,resizable=yes,width=500,height=500')">
                      <%=rs("rpi_total_out")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#B9B9FF"><strong><%=rs("rpi_total_in") + rs("rpi_total_out")%></strong></td>
                    </tr>
<% 
rpi_total_in = rpi_total_in + rs("rpi_total_in")
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
                      <td height="40" colspan="19" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_total_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=rpi_total_out%></strong></td>
                      <td align="center" bgcolor="#6F6FFF"><strong><%=rpi_balance%></strong></td>
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="19" align="right" bgcolor="#CCCCCC"><strong>Grand Total</strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=total_rpi_total_in%></strong></td>
                      <td align="center" bgcolor="#CCCCCC"><strong><%=total_rpi_total_out%></strong></td>
                      <td align="center" bgcolor="#6F6FFF"><strong><%=total_rpi_total_in+total_rpi_total_out%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="22" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_inventory_warehouse.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_warehouse.asp?num=" & Showed+row & link & "'> Next >></a>"
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
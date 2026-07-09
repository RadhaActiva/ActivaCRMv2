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
   orderby = "tblmodel.md_category"   
end if

sql2 = "select tblmodel.md_category,  " & _
		"sum(tblstocktran.stk_qty) as totalqty, round(sum(tblstocktran.stk_qty*tblmodel.md_averageecost),2) as totalvalue " & _
		"from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
		"where tblstocktran.stk_itm_code is not null and tblstocktran.stk_reference <> '0' " 
				
if searchvalue <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

   sql2 = sql2 & " group by tblmodel.md_category "

if orderby <> "" then
sql2 = sql2 & " order by " & orderby & " " & ordertype
else
sql2 = sql2 & " order by tblmodel.md_category"
end if

'response.write sql2

sql3 = "SELECT  " & _
		"sum(tblwarehouse_stock.wst_itm_current_qty) as totalqty,  " & _
		"sum(tblwarehouse_stock.wst_itm_current_qty * tblmodel.md_costprice) as totalvalue " & _
		"FROM tblwarehouse_stock inner join tblmodel on tblwarehouse_stock.wst_itm_code=tblmodel.md_code " & _
		"where tblwarehouse_stock.wst_itm_code is not null "
		
if searchvalue <> "" then 
   sql3 = sql3 & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

set rs = server.CreateObject("adodb.recordset")

rs.ActiveConnection = strconnect
rs.Source = sql3
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   t_totalqty = rs("totalqty")
   t_totalvalue = rs("totalvalue")
end if
rs.close
	
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
	document.form1.action = "rm_rpt_stockcard_group.asp";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Summary Stock By Group</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> Closing Stock Report Group by Product Category (Group Type)</td>
                      <td width="20%" align="center" class="titlegrey1"><img src="images/excel.jpg" width="57" height="21" /></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="action_report.asp?type=warehouselocation">
                
                
                <tr>
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF"><select name="searchitem" id="searchitem">
                    <option value="tblmodel.md_category" <% if searchitem = "tblmodel.md_category" then response.write " selected" %>>Category Name</option>
                    <option value="tblwarehouse_stock.wst_itm_code" <% if searchitem = "tblwarehouse_stock.wst_itm_code" then response.write " selected" %>>Item Code</option>
                    <option value="tblmodel.md_desc" <% if searchitem = "tblmodel.md_desc" then response.write " selected" %>>Item Description</option>
                </select>
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                    <select name="orderby" id="orderby">
                      <option value="totalvalue" <%if orderby="totalvalue" then response.write " selected"%>>Total</option>
                      <option value="tblwarehouse_stock.wst_itm_current_qty" <%if orderby="tblwarehouse_stock.wst_itm_current_qty" then response.write " selected"%>>Qty</option>
                      <option value="tblwarehouse_stock.wst_itm_code" <% if orderby = "tblwarehouse_stock.wst_itm_code" then response.write " selected" %>>Item Code</option>
                      <option value="tblmodel.md_desc" <% if orderby = "tblmodel.md_desc" then response.write " selected" %>>Item Desc</option>
                    </select>
                    <select name="ordertype" id="ordertype">                     
                      <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                       <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                  </select>
                    <span class="titlegrey1">
                    <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" />
                  </span></td>
                  <td width="20%" height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                 </form>
                
              <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="6" align="right" class="style1"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_stockcard_group.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_group.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="30" colspan="3" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Category</span></strong></font></td>
                      <td width="151" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
                      <td width="139" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total&nbsp;</strong></font></td>
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
                      <td colspan="3" align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("md_category")%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=rs("totalqty")%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2(rs("totalvalue"))%>&nbsp;</strong></td>
                    </tr>
<%

totalqty = totalqty + cint(rs("totalqty")) 


if isnumeric(rs("totalvalue")) then 
totalvalue = totalvalue + rs("totalvalue")
end if
 
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
                   
                    <tr bgcolor="#F3F3F3">
                     <td height="40" colspan="4" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=totalqty%></strong></td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(totalvalue)%>&nbsp;</strong></td>
                   </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="6" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_stockcard_group.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_group.asp?num=" & Showed+row & link & "'> Next >></a>"
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
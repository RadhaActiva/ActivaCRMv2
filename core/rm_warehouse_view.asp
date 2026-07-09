<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")

if ordertype = "" then 
   ordertype = "desc"
end if

i = 1

'06/06/24 - show only active warehouse	  
'sql = "  select * from (Select tblstocktran.stk_reference as wh_code, " & _
'		"(select top 1 tblwarehouse.wh_name from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_name, " & _
'		"(select top 1 tblwarehouse.wh_contact_person from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_contact_person," & _
'		"(select top 1 tblwarehouse.wh_remark from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_remark, " & _
'		"(select top 1 tblwarehouse.wh_tel from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference) as wh_tel, " & _
'		"sum(tblstocktran.stk_qty) totaqty, round(sum(tblstocktran.stk_qty*tblmodel.md_averageecost),2) as totalvalue " & _
'		"from tblstocktran inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _
'		"where tblstocktran.stk_reference <> '0' " 

'<a href="header_jsmis.asp">header_jsmis.asp</a>
'if searchvalue <> "" then 
'   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
'end if

'sql = sql & " group by tblstocktran.stk_reference )t inner join tblwarehouse on t.wh_code = tblwarehouse.wh_code where tblwarehouse.wh_status = 'Y'"

sql="select wh_code, wh_name, wh_contact_person, wh_remark, wh_tel, sum(b.stk_qty) as totaqty "  & _
"from tblwarehouse a " &_
"left join tblstocktran b on a.wh_code=b.stk_reference " & _
"group by wh_code,wh_name, wh_contact_person, wh_remark, wh_tel"


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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby")

%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Store</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_warehouse_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td width="84%">
                         <select name="searchitem" id="searchitem">
                            <option value="tblstocktran.stk_reference" <% if searchitem = "tblstocktran.stk_reference" then response.write " selected" %>>Store Code </option>
                          </select>
                          
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                           <option value="tblstocktran.stk_reference" <% if orderby = "tblstocktran.stk_reference" then response.write " selected" %>>Store Code </option>
                          </select>
                         <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="button" id="button3" value="Submit" />
                         </td>
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
					Response.Write " <a href='rm_warehouse_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_warehouse_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="5%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="12%" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Store Code</span></strong></font></td>
                      <td width="26%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Store Name </strong></font></td>
                      <td width="15%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Person In Charge</span></strong></font></td>
                      <td width="18%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                      <td width="12%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Tel Number</span></strong></font></td>
                      <td width="12%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Current Stock Qty</strong></font></td>
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
                     <td height="40"><%=j%> </td>
                      <td align="left" nowrap="nowrap"><a href="rm_warehouse_new.asp?wh_code=<%=rs("wh_code")%>"><strong><%=rs("wh_code")%></strong></a></td>
                      <td align="left" nowrap="nowrap"><%=rs("wh_name")%></td>
                      <td align="left"><%=rs("wh_contact_person")%></td>
                      <td align="left"><%=rs("wh_remark")%></td>
                      <td align="left"><%=rs("wh_tel")%></td>
                      <td align="center"><%=chknumber(rs("totaqty"))%></td>
                    </tr>
                    
<%
totaqty = totaqty + chknumber(rs("totaqty"))
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
 <tr bgcolor="<%=nbgcolor%>">
                      <td height="30" colspan="6" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td height="30" align="center" bgcolor="#CCCCCC"><strong><%=totaqty%></strong></td>
                    </tr>

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
					Response.Write " <a href='rm_warehouse_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_warehouse_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
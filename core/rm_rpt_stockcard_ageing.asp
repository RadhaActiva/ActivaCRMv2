<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
md_status = request("md_status")

if request("stock_aging_date") <> "" then 
   stock_aging_date = request("stock_aging_date")
else   
   stock_aging_date = chkdate(date())
end if

if ordertype = "" then 
   ordertype = "desc"
end if

i = 1

'added to exclude service and labour condition - 270222
'also added a new column totalqty = qty * avg price in the screen and excel report plus sub/grand total at every page

sql = "SELECT     tblStock_Agieng.ag_id, tblStock_Agieng.ag_stock_code, tblStock_Agieng.ag_current_stock, tblStock_Agieng.ag_averagecost, tblStock_Agieng.Y0, tblStock_Agieng.Y1, tblStock_Agieng.Y2, " & _
      "tblStock_Agieng.Y3, tblStock_Agieng.Y4,  " & _
	  "tblStock_Agieng.Y5, tblStock_Agieng.Y6, tblStock_Agieng.status, tblmodel.md_category, tblmodel.md_model, tblmodel.md_type, tblmodel.md_brands, tblmodel.md_desc, tblmodel.md_rcpprice, tblmodel.md_status " & _
	  "FROM         tblStock_Agieng INNER JOIN " & _
	  "tblmodel ON tblStock_Agieng.ag_stock_code = tblmodel.md_code " & _
	  "WHERE     (tblStock_Agieng.ag_current_stock <> 0) and tblStock_Agieng.ag_stock_code NOT IN ('Service','Labour')" 

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tblStock_Agieng.ag_stock_code, tblStock_Agieng.ag_current_stock"
end if


response.Cookies("GAPS")("stockageingsql") = sql
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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&orderby=" & request("orderby") & "&md_status=" & request("md_status") & "&ordertype=" & request("ordertype") & "&stock_aging_date=" & request("stock_aging_date") 

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
               <table width="100%"><!--for UI Purpose-->
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Stock Ageing</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_rpt_stockcard_ageing.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      
                      <tr>
                        <td class="titlegrey1">Stock In Date</td>
                        <td width="69%"><strong><font color="#000000"><strong>
                          <input name="stock_aging_date" type="text" id="stock_aging_date" value="<%=stock_aging_date%>" size="15" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.stock_aging_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong>                          
			 <input type="button" name="button2" id="button" value="Generate Stock Ageing" onclick="javascript:popup('rm_rpt_stockcard_ageing_process.asp?stock_aging_date='+ document.form1.stock_aging_date.value,'cb18','scrollbars=yes,resizable=yes,width=500,height=500')" /></td>
                        <td width="15%" align="center"><span class="titlegrey1"><a href="rm_rpt_stockcard_ageing_excel.asp?stock_aging_date=<%=stock_aging_date%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></span></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"></div></td>
                        <td colspan="2"><select name="searchitem" id="searchitem">
                          <option value="tblStock_Agieng.ag_stock_code" <% if searchitem = "tblStock_Agieng.ag_stock_code" then response.write " selected" %>>Stock Code</option>
                          <option value="tblmodel.md_desc" <% if searchitem = "tblmodel.md_desc" then response.write " selected" %>>Stock Desc</option>                        
                          <option value="tblmodel.md_group_type" <% if searchitem = "tblmodel.md_group_type" then response.write " selected" %>>Group Type</option>
                          <option value="tblmodel.md_model" <% if searchitem = "tblmodel.md_model" then response.write " selected" %>>Model</option>
                          <option value="tblmodel.md_type" <% if searchitem = "tblmodel.md_type" then response.write " selected" %>>Type</option>
                          <option value="tblmodel.md_brands" <% if searchitem = "tblmodel.md_brands" then response.write " selected" %>>Brand</option>
                        </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                          <option value="tblStock_Agieng.ag_current_stock" <% if orderby = "tblStock_Agieng.ag_current_stock" then response.write " selected" %>>Stock Balance</option>
                          <option value="tblStock_Agieng.ag_stock_code" <% if orderby = "tblStock_Agieng.ag_stock_code" then response.write " selected" %>>Stock Code</option>
                          <option value="tblmodel.md_desc" <% if orderby = "tblmodel.md_desc" then response.write " selected" %>>Stock Desc</option>                        
                          <option value="tblmodel.md_model" <% if orderby = "tblmodel.md_model" then response.write " selected" %>>Model</option>
                          <option value="tblmodel.md_type" <% if orderby = "tblmodel.md_type" then response.write " selected" %>>Type</option>
                          <option value="tblmodel.md_brands" <% if orderby = "tblmodel.md_brands" then response.write " selected" %>>Brand</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="button" id="button3" value="Submit" /></td>
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
					Response.Write " <a href='rm_rpt_stockcard_ageing.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ageing.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Stock Code</span></strong></font></td>
                      <td bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock Name </strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>RCP</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>AVG Cost</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Brand</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Status</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 0 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 1 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 2 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 3 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 4 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 5 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>&gt; 6 Y</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Stock Bal</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total Value</strong></font></td>
                    </tr>
                   
                    
<% 
totalqty=0
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
                     <td height="40" align="center"><%=j%> </td>
                      <td align="left" nowrap="nowrap"><a href="rm_stock_new.asp?md_code=<%=rs("ag_stock_code")%>"><strong><%=rs("ag_stock_code")%></strong></a></td>
                      <td><%=rs("md_desc")%> </td>
                      <td align="center"><%=chknumber2(rs("md_rcpprice"))%></td>
                      <td align="center" nowrap="nowrap">
					   <%if request.Cookies("GAPS")("view_cost")="Y" then %>
					  <%=chknumber2(rs("ag_averagecost"))%>
                          <% subtotal=chknumber(rs("ag_current_stock")) * chknumber2(rs("ag_averagecost")) %> 
                      <%else%>
                      Restricted View
                      <%end if%>
                       </td>
                      <td align="center"><%=rs("md_brands")%></td>
                      <td align="center"><%=rs("md_status")%></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y0"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y1"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y2"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y3"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y4"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y5"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("Y6"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(rs("ag_current_stock"))%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=chknumber(subtotal)%></strong></td>
                 </tr>
<%
t_Y0 =t_Y0+rs("Y0")
t_Y1 =t_Y1+rs("Y1")
t_Y2 =t_Y2+rs("Y2")
t_Y3 =t_Y3+rs("Y3")
t_Y4 =t_Y4+rs("Y4")
t_Y5 =t_Y5+rs("Y5")
t_Y6 =t_Y6+rs("Y6")
grand_total = grand_total + subtotal
ag_current_stock = ag_current_stock + rs("ag_current_stock")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>  
                 
                  <tr>
                      <td height="40" colspan="7" align="right"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y0%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y1%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y2%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y3%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y4%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y5%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=t_Y6%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=ag_current_stock%></strong></td>
                      <td align="center" nowrap="nowrap"><strong><%=grand_total%></strong></td>
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
					Response.Write " <a href='rm_rpt_stockcard_ageing.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ageing.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
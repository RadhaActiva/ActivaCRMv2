<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(date())
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if request("orderby") <> "" then 
   orderby = request("orderby")
else
   orderby = "tblstocktran.stk_itm_code"   
end if

    'changed left join to inner join 25/06 

    'changed the code to use avg cost from tblmodel instead - 260722 and removed service and labour
    '29/09/2022 gets the avg cost based on the date selected from tblmodel_avgcost
sql2 = "select tblstocktran.stk_itm_code, tblmodel.md_desc, tblmodel.md_category,tblmodel.md_costprice, " & _
       "sum(tblstocktran.stk_qty) as totalqty, " & _ 
       "(select TOP 1 md_averagecost from tblmodel_avgcost where md_code=tblstocktran.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_from & "' " & _ 
       "order by md_date desc) AS avgcost " & _
       "from tblstocktran " & _ 
       "inner join tblmodel on tblstocktran.stk_itm_code=tblmodel.md_code " & _ 
       "where tblstocktran.stk_itm_code NOT IN ('Service','Labour')  and tblstocktran.stk_reference is not null " & _
       "and cast(tblstocktran.stk_date as date) < '" & ChkDateYYYYMMDD(DateAdd("d",1,job_date_from)) & "' "
    
if searchvalue <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

'sql2 = sql2 & " group by tblstocktran.stk_itm_code, tblmodel.md_desc, tblmodel.md_category"  'new statement below - 240722
sql2 = sql2 & " group by tblstocktran.stk_itm_code, tblmodel.md_desc, tblmodel.md_category,tblmodel.md_costprice"

if orderby <> "" then
sql2 = sql2 & " order by " & orderby & " " & ordertype
else
sql2 = sql2 & " order by tblstocktran.stk_itm_code"
end if

'response.write sql2
response.Cookies("GAPS")("sqlexcel") = sql2

t_totalqty = 0
t_totalvalue = 0

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
	document.form1.action = "rm_rpt_inventory_productgroup.asp";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Inventory By Product Group</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> <strong>Closing Stock Report Group by Product Category (Group Type), Product </strong></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_inventory_productgroup_excel.asp?job_date_from=<%=job_date_from%>"" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="action_report.asp?type=warehouselocation">
                
                
                <tr>
                  <td height="30" align="left" bgcolor="#FFFFFF"><strong><font color="#000000"><strong>Up to
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                  <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </td>
                  <td height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF"><select name="searchitem" id="searchitem">
                    <option value="tblmodel.md_category" <% if searchitem = "tblmodel.md_category" then response.write " selected" %>>Category Name</option>
                    <option value="tblstocktran.stk_itm_code" <% if searchitem = "tblstocktran.stk_itm_code" then response.write " selected" %>>Item Code</option>
                    <option value="tblmodel.md_desc" <% if searchitem = "tblmodel.md_desc" then response.write " selected" %>>Item Description</option>
                </select>
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                    <select name="orderby" id="orderby">
                      <option value="totalvalue" <%if orderby="totalvalue" then response.write " selected"%>>Total</option>
                      <option value="sum(tblstocktran.stk_qty)" <%if orderby="sum(tblstocktran.stk_qty)" then response.write " selected"%>>Qty</option>
                      <option value="tblstocktran.stk_itm_code" <% if orderby = "tblstocktran.stk_itm_code" then response.write " selected" %>>Item Code</option>
                      <option value="tblmodel.md_desc" <% if orderby = "tblmodel.md_desc" then response.write " selected" %>>Item Desc</option>
                    </select>
                    <select name="ordertype" id="ordertype">                     
                      <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                       <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                  </select>
                    <span class="titlegrey1"><input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" /></span></td>
                  <td width="20%" height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                 </form>
                
              <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="7" align="right" class="style1"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_inventory_productgroup.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_productgroup.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Category</span></strong></font></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Item  Code.</strong></font></td>
                      <td width="343" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item  Name</span></strong></font></td>
                      <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
                      <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Costing (Unit)</span></strong></font></td>
                      <td width="139" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total Value</strong></font></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Average Cost<span><br /></span></strong></font></td>
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

'total = rs("totalqty")*rs("avgcost")

if (IsNull(rs("avgcost"))) Then
	total = 0
else
	total = rs("totalqty") * rs("md_costprice")
end if 

totalavg = total

'total = 0
	
%>                   
                   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"><%=j%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("md_category")%></font></strong></td>
                      <td align="left" nowrap="nowrap"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_inventory_productgroup_detail.asp?stk_itm_code=<%=rs("stk_itm_code")%>','cb18','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("stk_itm_code")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"> <%=rs("md_desc")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=rs("totalqty")%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2(rs("md_costprice"))%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=chknumber2(total)%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2(rs("avgcost"))%></strong></td>
                    </tr>
<%

'if isnumeric(total) then 
'	totalvalue = totalvalue + total
'end if

t_totalqty = t_totalqty + cint(rs("totalqty")) 
t_totalvalue = t_totalvalue + total 

count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
                   
<!--                    <tr bgcolor="#F3F3F3">
                     <td height="40" colspan="4" align="right" bgcolor="#CCCCCC"><strong>Sub-Total</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=totalqty%></strong></td>
                     <td align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(totalvalue)%>&nbsp;</strong></td>
                   </tr>
-->
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=t_totalqty%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(t_totalvalue)%>&nbsp;</strong></td>
                      <td align="right" bgcolor="#999999"></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_inventory_productgroup.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_inventory_productgroup.asp?num=" & Showed+row & link & "'> Next >></a>"
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
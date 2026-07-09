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
   orderby = "tblstocktran.stk_reference"   
end if

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


' sql2 = "select s.stk_reference, s.wh_name,sum(qty) as Totalstockqty, sum(totalvalue)as totalvalue from " & _
'"(select t.stk_reference, t.wh_name, t.qty,t.qty * (select TOP 1 md_averagecost from tblmodel_avgcost where md_code=t.stk_itm_code and " & _
'"CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_from & "' order by md_date desc) as totalvalue from " & _
'"(SELECT a.stk_reference, b.wh_name,stk_itm_code,sum(a.stk_qty) as qty from tblstocktran A  inner join "  & _
'"tblwarehouse b on a.stk_reference = b.wh_code where " & _
'"CAST(a.stk_Date as date) <= '" & job_date_from & "' and stk_itm_code not in ('Service', 'Labour') and b.wh_status = 'Y' " & _
'"group by a.stk_reference,a.stk_itm_code,b.wh_name) t " & _
'"where t.qty >0 )s " & _
'"group by s.stk_reference,s.wh_name order by s.stk_reference" 


 '"where CAST(a.stk_Date as date) <= '" & job_date_from & "' and stk_itm_code not in ('Service', 'Labour') " & _
'"and b.wh_status = 'Y' group by a.stk_reference,a.stk_itm_code,b.wh_name)t)s " & _

sql2="select s.stk_reference, s.wh_name,sum(s.qty) as Totalstockqty, sum(s.totalcost) as totalvalue from " & _
"(select t.stk_reference,t.wh_name, t.qty, (t.qty * t.avg_cost) as totalcost from ( " & _
"SELECT a.stk_reference, b.wh_name,stk_itm_code,sum(a.stk_qty) as qty,(select TOP 1 md_averagecost from tblmodel_avgcost where md_code=a.stk_itm_code and CAST(tblmodel_avgcost.md_date as date) <= '" & job_date_from & "' " & _
"order by md_date desc) as avg_cost from tblstocktran A " & _
"inner join tblwarehouse b on a.stk_reference = b.wh_code " & _
"where CAST(a.stk_Date as date) < '" & ChkDateYYYYMMDD(DateAdd("d",1,job_date_from)) & "' and stk_itm_code not in ('Service', 'Labour') " & _
"group by a.stk_reference,a.stk_itm_code,b.wh_name)t)s " & _
"group by s.stk_reference,s.wh_name order by s.stk_reference"

'rm_rpt_stockcard_warehouyse.asp 

response.Cookies("GAPS")("sqlexcel") = sql2

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
	document.form1.action = "rm_rpt_stockcard_warehouse.asp";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Summary By Store Location</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> <strong>Closing Stock Report Group by Store</strong></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_stockcard_warehouse_excel.asp?job_date_from=<%=job_date_from%>"" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="rm_rpt_stockcard_warehouse.asp">
                <tr>
                  <td height="30" align="left" bgcolor="#FFFFFF"><strong><font color="#000000"><strong>
                    Up to 
                    <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                    <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong>Date must be (dd-MMM-yyyy) eg: 21-May-2015
                  <td height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                    
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF">
                <!--<select name="searchitem" id="searchitem">
                    <option value="tblwarehouse_stock.wst_wh_code" <% if searchitem = "tblwarehouse_stock.wst_wh_code" then response.write " selected" %>></option>
                </select>-->
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" hidden />
                    <!--<select name="orderby" id="orderby">
                      <option value="totalvalue" <%if orderby="totalvalue" then response.write " selected"%>>Total</option>
                      <option value="tblwarehouse_stock.wst_itm_current_qty" <%if orderby="tblwarehouse_stock.wst_itm_current_qty" then response.write " selected"%>>Qty</option>
                      <option value="tblwarehouse_stock.wst_wh_code" <% if orderby = "tblwarehouse_stock.wst_wh_code" then response.write " selected" %>>Warehouse</option>
                    </select>-->
                   <!-- <select name="ordertype" id="ordertype">                     
                      <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                       <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                  </select>-->
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
					Response.Write " <a href='rm_rpt_stockcard_warehouse.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_warehouse.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="30" colspan="3" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Store</span></strong></font></td>
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
                      <td colspan="3" align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_warehouse_new.asp?wh_code=<%=rs("stk_reference")%>" target="_blank"><%=rs("stk_reference")%> - <%=rs("wh_name")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=rs("Totalstockqty")%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2(rs("totalvalue"))%>&nbsp;</strong></td>
                    </tr>
<%

totalqty = totalqty + cint(rs("Totalstockqty")) 
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
					Response.Write " <a href='rm_rpt_stockcard_warehouse.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_warehouse.asp?num=" & Showed+row & link & "'> Next >></a>"
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
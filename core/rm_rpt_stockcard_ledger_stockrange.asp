<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
searchvalue2 = request("searchvalue2")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
wh_code = request("wh_code")
whchk = request("whchk")
reccount=request("reccount")
    
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

    totalqty_in = 0
    totalqty_out = 0 
    total_balance = 0

    'state some condition so that the report doesn't start with default data
    'sql4= "select stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity " & _
    '"from tblstock_ledger_stock_range a where stk_code_id >= '" & searchvalue & "' and stk_code_id <= '" & searchvalue2 & "' and stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "' order by stk_date"
    if whchk = "No" then 
        sql4 = "select stk_itm_code, sum(stk_qty) AS 'stk_qty' from tblstocktran " & _
	    "where stk_type in ('Job', 'Spareparts-Request','Stock-Transfer-Out','Stock-Transfer-In','Stock-In','Stock-Out','Stock-Adj','CN-Cancel','DO') " & _
	    "and stk_itm_code  >= '" & searchvalue & "' and stk_itm_code<= '" & searchvalue2 & "' and CAST(tblstocktran.stk_date as date) >= '" & job_date_from & "' AND "  & _
	    "CAST(tblstocktran.stk_date as date) <= '" & job_date_to & "' and stk_reference='" & wh_code & "' " & _
	    "group by stk_itm_code"
    else
        sql4 = "select stk_itm_code, sum(stk_qty) AS 'stk_qty' from tblstocktran " & _
	    "where stk_type in ('Job', 'Spareparts-Request','Stock-Transfer-Out','Stock-Transfer-In','Stock-In','Stock-Out','Stock-Adj','CN-Cancel','DO') " & _
	    "and stk_itm_code  >= '" & searchvalue & "' and stk_itm_code<= '" & searchvalue2 & "' and CAST(tblstocktran.stk_date as date) >= '" & job_date_from & "' AND "  & _
	    "CAST(tblstocktran.stk_date as date) <= '" & job_date_to & "'" & _
	    "group by stk_itm_code"
    end if

    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open

response.Cookies("GAPS")("sqlexcel") = sql4

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs4.RecordCount
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
'end if

count = count + Showed
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&orderby=" & orderby & "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&searchvalue2=" & searchvalue2 & "&Searchor_date=" & Searchor_date & "&ordertype=" & ordertype & "&wh_code=" & wh_code &  "&whchk=" & whchk & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to

%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_stockcard_ledger_stockrange.asp?post=yes";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Stock Ledger By Stock Range</div>
                          </td>
                      </tr>
                    </table>
                    </td>
                </tr>
           
                <tr>
                    <br />
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                       <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                    <tr>
                      <td width="80%"> <strong>Stock Ledger by Store and Product</strong></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_stockcard_ledger_stockrange_excel.asp?searchvalue=<%=searchvalue%>&searchvalue2=<%=searchvalue2%>&wh_code=<%=wh_code%>&whchk=<%=whchk%>&job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="rm_rpt_stockcard_ledger_stockrange.asp?post=yes">
                     <tr>
                        <td height="30" align="left" bgcolor="#FFFFFF" width="45%"><strong><font color="#000000"><strong>Date from<font color="#000000">
                        <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>
                         <font color="#000000"><strong>to 
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>" size="15" />
                  <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong><td> Date must be (dd-MMM-yyyy) eg: 21-May-2022 </>
                  <td height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                  <tr>
                   <td height="30" align="left" bgcolor="#FFFFFF"> <Strong>Store  </Strong>&nbsp;
                   <select name="wh_code" id="wh_code">
                      <%			
				sql2 = "SELECT wh_id, wh_code, wh_name, wh_remark FROM tblwarehouse order by wh_code"	
                set rs2 = server.CreateObject("adodb.recordset")
				rs2.Open sql2,strconnect,3,3,&H0001
                while Not rs2.EOF
					  if (wh_code) = (rs2("wh_code")) then
					  response.write "<option value='" & rs2("wh_code") & "' selected>" & rs2("wh_code") & " - " & rs2("wh_name") & " - " & rs2("wh_remark") & "</option>"
                      else
					  response.write "<option value='" & rs2("wh_code") & "'>" & rs2("wh_code") & " - " & rs2("wh_name") & " - " & rs2("wh_remark") & "</option>"
					  end if 					  
                     
				rs2.movenext
				wend
				rs2.close					
				%>
                    </select>
               
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF">Include All Store &nbsp;&nbsp
                     <select name="whchk" id="whchk">                     
                       <option value="No" <% if whchk = "No" then response.write " selected"%>>No</option>
                      <option value="Yes" <% if whchk = "Yes" then response.write " selected"%>>Yes</option>
                  </select></td>
                  
                <tr><td height="30" bgcolor="#FFFFFF"> <Strong>Start Stock Code  </Strong>
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />                    
                   [<a href="javascript:popup('rm_spareparts_list.asp?searchitem=md_type&amp;searchvalue=<%=cust_code%>&amp;formname=form1&fieldname=searchvalue&fieldname1=desc1','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                    <td height="30" bgcolor="#FFFFFF"> <Strong>End Stock Code  </Strong>
                    <input name="searchvalue2" type="text" id="searchvalue2" value="<%=searchvalue2%>" /> 
                   [<a href="javascript:popup('rm_spareparts_list.asp?searchitem=md_type&amp;searchvalue=<%=cust_code%>&amp;formname=form1&fieldname=searchvalue2&fieldname1=desc2','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] 
                  <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" /></td></tr>
                      <tr>
                          <td><input name="desc1" type="hidden" id="desc1" value="<%=desc1%>"/></td>
                          <td><input name="desc2" type="hidden" id="desc2" value="<%=desc2%>"/></td>
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
					Response.Write " <a href='rm_rpt_stockcard_ledger_stockrange.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ledger_stockrange.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="6%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">No</font></strong></td>
                      <td width="30%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Stock Code</font></strong></td>
                      <td width="30%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Stock Desc</font></strong></td>
                      <td width="30%" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Quantity</strong></td>          
                    </tr>
 <%
        if not rs4.eof then
        rs4.Move Showed
        count = Showed + 1
        end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if	

     partdesc = ""
     sql = "select md_desc from tblmodel where md_code = '" & rs4("stk_itm_code") & "'  "
     partdesc = selectid(sql)
%>                   
                     
                      
                     <tr bgcolor="<%=nbgcolor%>">
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=count%></strong></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF">
                     <a href="javascript:popup('rm_rpt_stockcard_ledger_list.asp?searchvalue=<%=rs4("stk_itm_code")%>&whchk=<%=whchk%>&wh_code=<%=wh_code%>&job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&post=yes','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs4("stk_itm_code")%></a></td>                       
                  <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=partdesc%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs4("stk_qty")%></strong></td>
                    </tr>
<%

    totalqty = totalqty + ChkNumberInt(rs4("stk_qty"))

    count = count + 1 
    i = i + 1
    rs4.MoveNext
    Next
    'rs4.Close
    'Set rs4 = Nothing
    'Set rs1 = Nothing
    'Set rs2 = Nothing
    'Set rs3 = Nothing
'end if 
%>         
                    <tr bgcolor="#F3F3F3">
                      <td align="right" nowrap="nowrap" bgcolor="#999999"></td>
                      <td height="40"  align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_stockcard_ledger_stockrange.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ledger_stockrange.asp?num=" & Showed+row & link & "'> Next >></a>"
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
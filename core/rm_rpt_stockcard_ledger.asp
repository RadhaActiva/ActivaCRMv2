<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
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

    sql4= "select stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_balance " & _
    "from tblstock_ledger a where stk_code_id = '" & searchvalue & "' and stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "' order by stk_date"
    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open
   
If Request.QueryString("post") = "yes" Then
    sql6= "delete FROM tblstock_ledger where stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "'" 'delete at user level 
    set rs6 = server.CreateObject("adodb.recordset")
    rs6.ActiveConnection = strconnect
    rs6.Source = sql6
    rs6.Open    
End if 

if searchvalue <> "" and Request.QueryString("post") = "yes" then 'this section loads the fist time and each time the page is rebsubmitted (after value changes)

    'Delete all data by user
    sql8 = "delete FROM tblstock_ledger where stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "'"
    set rs8 = server.CreateObject("adodb.recordset")
    rs8.ActiveConnection = strconnect
    rs8.Source = sql8
    rs8.Open

    if whchk = "No" then  'this statement only updated the first value ir balance b/f
       sql1="insert into tblstock_ledger (stk_code_id, stk_date, stk_doc_type, stk_balance, stk_logby) " & _
             "select '" & searchvalue & "', '" & job_date_from & "', 'Bal B/F' , sum(tblstocktran.stk_qty) as totalqty, '" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran where tblstocktran.stk_id is not null " & _
             "and CAST(tblstocktran.stk_date as date) <=  '" & job_date_from & "' and stk_itm_code = '" & searchvalue & "' and stk_reference='" & wh_code & "'" 
    else
       sql1="insert into tblstock_ledger (stk_code_id, stk_date, stk_doc_type, stk_balance, stk_logby) " & _
             "select '" & searchvalue & "', '" & job_date_from & "', 'Bal B/F' , sum(tblstocktran.stk_qty) as totalqty,'" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran where tblstocktran.stk_id is not null " & _
             "and CAST(tblstocktran.stk_date as date) <=  '" & job_date_from & "' and stk_itm_code = '" & searchvalue & "'"
    end if 

    'response.write sql1

    set rs1 = server.CreateObject("adodb.recordset")
    rs1.ActiveConnection = strconnect
    rs1.Source = sql1
    rs1.CursorLocation  = 3
    rs1.Open

    balance_bf = 0.00

    'grab the total qty (bal/cf) from the table 

    sql3 = "select stk_balance from tblstock_ledger where stk_code_id = '" & searchvalue & "' and stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "'"
    set rs3 = server.CreateObject("adodb.recordset")
    rs3.ActiveConnection = strconnect
    rs3.Source = sql3
    rs3.CursorLocation  = 3
    rs3.Open
    if not rs3.eof then
        balance_bf = rs3("stk_balance")
    else 
        balance_bf = 0
    end if
 
    if IsNull(rs3("stk_balance")) then
        balance_bf = 0
    end if 

    'start a loop to insert each records into tblstock_ledger

    if whchk = "No" then 
        sql9 = "SELECT * FROM tblstocktran WHERE CAST(tblstocktran.stk_date as date) > '" & job_date_from & "' AND CAST(tblstocktran.stk_date as date) <= '" & job_date_to & "'" & _
                "AND stk_itm_code = '" & searchvalue & "' and stk_reference='" & wh_code & "' ORDER BY stk_date"
    else
        sql9 = "SELECT * FROM tblstocktran WHERE CAST(tblstocktran.stk_date as date) > '" & job_date_from & "' AND CAST(tblstocktran.stk_date as date) <= '" & job_date_to & "'" & _
                "AND stk_itm_code = '" & searchvalue & "' ORDER BY stk_date"
    end if 

    set rs9 = server.CreateObject("adodb.recordset")
    rs9.ActiveConnection = strconnect
    rs9.Source = sql9
    rs9.CursorLocation = 3
    rs9.Open
    
    while Not rs9.EOF

    'if isNumeric(balance_bf) = False then 
    'prev_total_cost = 0.00
    'end if 

    balance_bf = balance_bf + rs9("stk_qty")
            
    if rs9("stk_type") = "Job" then
       sql10= "insert into tblstock_ledger (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_balance,stk_logby) "  & _
            "select stk_itm_code, stk_date, stk_type , stk_voucherno, b.job_cust_code, b.job_cust_name, a.stk_qty,'" & balance_bf & "','" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a inner join tbljob b on " & _
            "a.stk_voucherno = b.job_code where a.stk_id is not null and stk_itm_code = '" & rs9("stk_itm_code") & "'  and a.stk_voucherno = '" & rs9("stk_voucherno") & "'"
       set rs10 = server.CreateObject("adodb.recordset")
       rs10.ActiveConnection = strconnect
       rs10.Source = sql10
       rs10.Open        
    elseif rs9("stk_type") = "Spareparts-Request" or rs9("stk_type") = "Stock-Transfer-Out" or rs9("stk_type") = "Stock-Transfer-In"  or rs9("stk_type") = "Stock-In" or rs9("stk_type") = "Stock-Out" or rs9("stk_type") = "Stock-Adj" or rs9("stk_type") = "CN-Cancel" then
        longdesc =  rs9("stk_fromwarehouse") + " to " + rs9("stk_towarehouse")
    
        if trim(rs9("stk_towarehouse")) = "" then 'if either para is missing then dont show the 'to'
            longdesc =  rs9("stk_fromwarehouse")
        elseif trim(rs9("stk_fromwarehouse")) = "" then
            longdesc =  rs9("stk_towarehouse")
        end if

       sql10= "INSERT INTO tblstock_ledger (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_balance, stk_logby) "  & _
              "VALUES ('" & rs9("stk_itm_code") & "', '" & rs9("stk_date") & "', '" & rs9("stk_type") & "', '" & rs9("stk_voucherno") & "', '','" & longdesc & "', '" & rs9("stk_qty") & "','" & balance_bf & "','" &  Request.Cookies("GAPS")("sloginid") & "' )"
       set rs10 = server.CreateObject("adodb.recordset")
       rs10.ActiveConnection = strconnect
       rs10.Source = sql10
       rs10.Open        
       longdesc=""
    elseif rs9("stk_type") = "DO" then
        sql10= "insert into tblstock_ledger (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_balance, stk_logby) "  & _
            "select stk_itm_code, stk_date, stk_type , stk_voucherno,b.inv_cust_code, b.inv_cust_name, a.stk_qty,'" & balance_bf & "','" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a inner join  tblinvoice b on " & _
            "a.stk_voucherno =  b.inv_dono where a.stk_id is not null and stk_itm_code = '" & rs9("stk_itm_code") & "'  and a.stk_voucherno = '" & rs9("stk_voucherno") & "'"
       set rs10 = server.CreateObject("adodb.recordset")
       rs10.ActiveConnection = strconnect
       rs10.Source = sql10
       rs10.Open       
    End if
  
    rs9.movenext
	wend

   ' rs4.close
    'Set rs4 = Nothing
  
    sql4= "select stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_balance " & _
    "from tblstock_ledger a where stk_code_id = '" & searchvalue & "' and stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "' order by stk_date"
    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open
  end if   

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
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&orderby=" & orderby & "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&Searchor_date=" & Searchor_date & "&ordertype=" & ordertype & "&wh_code=" & wh_code &  "&whchk=" & whchk & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to

%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_stockcard_ledger.asp?post=yes";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Stock Ledger</div>
                          </td>
                      </tr>
                    </table>
                    </td>
                </tr>
           
                <tr>
                    <br />
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> <strong>Stock Ledger by Store and Product</strong></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_stockcard_ledger_excel.asp?searchvalue=<%=searchvalue%>&wh_code=<%=wh_code%>&whchk=<%=whchk%>&job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="rm_rpt_stockcard_ledger.asp?post=yes">
                     <tr>
                        <td height="30" align="left" bgcolor="#FFFFFF" width="40%"><strong><font color="#000000"><strong>Date from<font color="#000000">
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
                   <td height="30" align="left" bgcolor="#FFFFFF"> <Strong>Store </Strong>&nbsp;
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
                  
                <tr><td height="30" bgcolor="#FFFFFF"> <Strong>Stock Code  </Strong>
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" /></td>                    
                  <td><input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" /></td></tr>
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
					Response.Write " <a href='rm_rpt_stockcard_ledger.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ledger.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="40" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">No</font></strong></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Date</strong></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Doc Type</strong></td>
                      <td width="50" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Doc No</strong></td>
                      <td width="109" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Code</strong></td>
                      <td width="104" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Description</strong></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Qty In</strong></td>          
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Qty Out</strong></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Balance</strong></tr>                    
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
%>                   
                     
                      <%
                      stk_qty_in=0
                      stk_qty_out=0
                      if rs4("stk_quantity") >= 0 then
                          stk_qty_in = rs4("stk_quantity")
                      elseif  rs4("stk_quantity") < 0 then 
                          stk_qty_out = rs4("stk_quantity")
                      elseif isnull(rs4("stk_quantity")) then
                        stk_qty_in=0
                        stk_qty_out=0
                      end if  
                    
                    totalqty_in = totalqty_in + cint(stk_qty_in) 
                    totalqty_out = totalqty_out + cint(stk_qty_out) 
                    total_balance = rs4("stk_balance")
                    %>

                     <tr bgcolor="<%=nbgcolor%>">
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=count%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=chkdate(rs4("stk_date"))%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_doc_type")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_doc_no")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_code")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_description")%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=stk_qty_in%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=stk_qty_out%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_balance")%></td>
                    </tr>
<%


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
                      <td height="40" colspan="6" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty_in%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty_out%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=total_balance%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_stockcard_ledger.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ledger.asp?num=" & Showed+row & link & "'> Next >></a>"
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
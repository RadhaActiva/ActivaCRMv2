<!-- #include file="database/datastore.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
searchvalue2 = request("searchvalue2")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
wh_code = request("wh_code")
whchk = request("whchk")
reccount=request("reccount")%>

<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>
<%    
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
    sql4= "select stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity " & _
    "from tblstock_ledger_stock_range a where stk_code_id >= '" & searchvalue & "' and stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "' order by stk_date"
    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open
   
if searchvalue <> "" and Request.QueryString("post") = "yes" then 'this section loads the fist time and each time the page is rebsubmitted (after value changes)
    'Delete all data by user
    sql8 = "delete FROM tblstock_ledger_stock_range where stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "'"
    set rs8 = server.CreateObject("adodb.recordset")
    rs8.ActiveConnection = strconnect
    rs8.Source = sql8
    rs8.Open

    'start a loop to insert each records into tblstock_ledger_stock_range

    if whchk = "No" then 
        sql9 = "SELECT * FROM tblstocktran WHERE CAST(tblstocktran.stk_date as date) >= '" & job_date_from & "' AND CAST(tblstocktran.stk_date as date) <= '" & job_date_to & "'" & _
                "AND stk_itm_code = '" & searchvalue & "' AND stk_reference='" & wh_code & "' ORDER BY stk_date"
    else
        sql9 = "SELECT * FROM tblstocktran WHERE CAST(tblstocktran.stk_date as date) >= '" & job_date_from & "' AND CAST(tblstocktran.stk_date as date) <= '" & job_date_to & "'" & _
                "AND stk_itm_code = '" & searchvalue & "' ORDER BY stk_date"
    end if 

    set rs9 = server.CreateObject("adodb.recordset")
    rs9.ActiveConnection = strconnect
    rs9.Source = sql9
    rs9.CursorLocation = 3
    rs9.Open
    
    while Not rs9.EOF
            
    if rs9("stk_type") = "Job" then
       sql10= "insert into tblstock_ledger_stock_range (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_logby) "  & _
            "select stk_itm_code, stk_date, stk_type , stk_voucherno, b.job_cust_code, b.job_cust_name, a.stk_qty,'" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a inner join tbljob b on " & _
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

       sql10= "INSERT INTO tblstock_ledger_stock_range (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_logby) "  & _
              "VALUES ('" & rs9("stk_itm_code") & "', '" & rs9("stk_date") & "', '" & rs9("stk_type") & "', '" & rs9("stk_voucherno") & "', '','" & longdesc & "', '" & rs9("stk_qty") & "','" &  Request.Cookies("GAPS")("sloginid") & "' )"
       set rs10 = server.CreateObject("adodb.recordset")
       rs10.ActiveConnection = strconnect
       rs10.Source = sql10
       rs10.Open        
       longdesc=""
    elseif rs9("stk_type") = "DO" then
        sql10= "insert into tblstock_ledger_stock_range (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity, stk_logby) "  & _
            "select stk_itm_code, stk_date, stk_type , stk_voucherno,b.inv_cust_code, b.inv_cust_name, a.stk_qty,'" &  Request.Cookies("GAPS")("sloginid") & "' from tblstocktran a inner join  tblinvoice b on " & _
            "a.stk_voucherno =  b.inv_dono where a.stk_id is not null and stk_itm_code = '" & rs9("stk_itm_code") & "'  and a.stk_voucherno = '" & rs9("stk_voucherno") & "'"
       set rs10 = server.CreateObject("adodb.recordset")
       rs10.ActiveConnection = strconnect
       rs10.Source = sql10
       rs10.Open       
    End if
  
    rs9.movenext
	wend
  
    sql4= "select stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_code, stk_description, stk_quantity " & _
    "from tblstock_ledger_stock_range a where stk_logby = '" &  Request.Cookies("GAPS")("sloginid") & "' order by stk_date" 'can show without condition as its already extracted bsaed on condition
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
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&orderby=" & orderby & "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&searchvalue2=" & searchvalue2 & "&Searchor_date=" & Searchor_date & "&ordertype=" & ordertype & "&wh_code=" & wh_code &  "&whchk=" & whchk & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to

%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_stockcard_ledger_list.asp?post=yes";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Stock Ledger By Stock ID - <%=searchvalue%></div>
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
                  </table></td>
                </tr>                    
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
					Response.Write " <a href='rm_rpt_stockcard_ledger_list.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ledger_list.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="6%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">No</font></strong></td>
                      <td width="15%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Stock Code</font></strong></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Date</strong></td>
                      <td width="20%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Doc Type</strong></td>
                      <td width="13%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Doc No</strong></td>
                      <td width="109" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Cust Code</strong></td>
                      <td width="32%" height="30" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Description</strong></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Qty In</strong></td>          
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Qty Out</strong></td>
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
                   ' total_balance = rs4("stk_balance")
                    %>

                     <tr bgcolor="<%=nbgcolor%>">
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=count%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs4("stk_code_id")%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=chkdate(rs4("stk_date"))%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_doc_type")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_doc_no")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_code")%></td>
                      <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_description")%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=stk_qty_in%></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=stk_qty_out%></td>
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
                      <td align="right" nowrap="nowrap" bgcolor="#999999"></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty_in%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=totalqty_out%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_stockcard_ledger_list.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard_ledger_list.asp?num=" & Showed+row & link & "'> Next >></a>"
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
    </table>
</body>
</html>

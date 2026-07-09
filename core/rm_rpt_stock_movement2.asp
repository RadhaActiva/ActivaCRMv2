<!-- #include file="header.asp" -->'
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

    sql5= "SELECT COUNT(*) AS Total FROM tblstock_movement"
    set rs5 = server.CreateObject("adodb.recordset")
    rs5.ActiveConnection = strconnect
    rs5.Source = sql5
    rs5.Open
   
    sql4= "select stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_purchase_cost, stk_avg_cost,stk_purchase_totalcost, " & _
    "stk_total_cost from tblstock_movement a where stk_code_id = '" & searchvalue & "' order by stk_date"
    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open
   
if searchvalue = "" then
    sql6= "delete FROM tblstock_movement"
    set rs6 = server.CreateObject("adodb.recordset")
    rs6.ActiveConnection = strconnect
    rs6.Source = sql6
    rs6.Open    
End if 

if searchvalue <> "" and rs5("total") < "1" then ''this code should be executed on the 1st time loading data only

    'Delete all data
    sql8 = "delete from tblstock_movement"
    set rs8 = server.CreateObject("adodb.recordset")
    rs8.ActiveConnection = strconnect
    rs8.Source = sql8
    rs8.Open

    if whchk = "No" then 
        sql1="insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_avg_cost, stk_purchase_cost) " & _
            "select top 1 a.stk_itm_code,'2022-01-01','Balance','Balance C/F','0.00',(select sum (a.stk_qty) from tblstocktran a " & _
            "where a.stk_itm_code ='" & searchvalue & "' and a.stk_date <='01/01/2022' and a.stk_reference='" & wh_code & "') as 'bal_cf', (select md_averageecost from tblmodel where tblmodel.md_code='" & searchvalue & "') as avgcost, " & _
            "(select md_costprice from tblmodel where tblmodel.md_code='" & searchvalue & "') as purchase_price from tblstocktran a where a.stk_itm_code ='" & searchvalue & "' " & _
            " and a.stk_date < '" & job_date_to & "' and a.stk_reference='" & wh_code & "'"
    else
        sql1="insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_avg_cost, stk_purchase_cost) " & _
            "select top 1 a.stk_itm_code,'2022-01-01','Balance','Balance C/F','0.00',(select sum (a.stk_qty) from tblstocktran a " & _
            "where a.stk_itm_code ='" & searchvalue & "' and a.stk_date <='01/01/2022') as 'bal_cf', (select md_averageecost from tblmodel where tblmodel.md_code='" & searchvalue & "') as avgcost, " & _
            "(select md_costprice from tblmodel where tblmodel.md_code='" & searchvalue & "') as purchase_price from tblstocktran a where a.stk_itm_code ='" & searchvalue & "' " & _
            " and a.stk_date < '" & job_date_to & "'"
   end if 

        set rs1 = server.CreateObject("adodb.recordset")
        rs1.ActiveConnection = strconnect
        rs1.Source = sql1
        rs1.CursorLocation  = 3
        rs1.Open

    if whchk = "No" then 
        sql3= "insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty,stk_bf_qty,stk_avg_cost, stk_purchase_cost)" & _
        "select a.stk_itm_code, a.stk_date,a.stk_type, stk_voucherno,stk_qty,'0.00','0.00',stk_cost_price from tblstocktran a " & _
        "where  stk_date >= '01/01/2022' and stk_itm_code = '" & searchvalue & "' and a.stk_reference='" & wh_code & "' order by stk_date "
    else
        sql3= "insert into tblstock_movement (stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty,stk_bf_qty,stk_avg_cost, stk_purchase_cost)" & _
        "select a.stk_itm_code, a.stk_date,a.stk_type, stk_voucherno,stk_qty,'0.00','0.00',stk_cost_price from tblstocktran a " & _
        "where  stk_date >= '01/01/2022' and stk_itm_code = '" & searchvalue & "' order by stk_date "
    end if 

    set rs3 = server.CreateObject("adodb.recordset")
    rs3.ActiveConnection = strconnect
    rs3.Source = sql3
    rs3.CursorLocation  = 3
    rs3.Open
 
    sql4= "select stk_move_id, stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_purchase_cost, stk_avg_cost,stk_total_cost " & _
    "from tblstock_movement a where stk_code_id = '" & searchvalue & "' order by stk_date"
    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open
    
    while Not rs4.EOF
    totalpurchase_cost = "0.00"
    total_cost= "0.00"
    prev_total_cost= "0.00"    

    if rs4("stk_doc_no") = "Balance C/F" then
         total_cost = rs4("stk_bf_qty") * (rs4("stk_avg_cost"))
         balanceBFqty = rs4("stk_bf_qty") 
         weighted_avg_cost = rs4("stk_avg_cost")
         prev_total_cost = rs4("stk_total_cost")
    elseif rs4("stk_doc_type") = "Stock-In" then
           totalpurchase_cost = rs4("stk_qty") * rs4("stk_purchase_cost")
    else 
           totalpurchase_cost = rs4("stk_qty") * prev_stk_avg_cost
    end if 
     
      total_cost = totalpurchase_cost + prev_total_cost
      balanceBFqty = balanceBFqty + rs4("stk_qty")
        
        response.write rs4("stk_qty")
        response.write "/"
    
      if total_cost > 0 and balanceBFqty > 0 then 'this is to prevent sys crash if divided by 0
        weighted_avg_cost = total_cost / balanceBFqty        
      end if 

      if rs4("stk_doc_type") = "Spareparts-Request" or rs4("stk_doc_type") = "Stock-Transfer-Out" or rs4("stk_doc_type") = "Stock-Transfer-In" then
        weighted_avg_cost = prev_stk_avg_cost
      End if

        sql7= "update tblstock_movement set stk_bf_qty =  '" & balanceBFqty & "', stk_purchase_totalcost = '" & totalpurchase_cost & "', stk_total_cost =  '" & total_cost & "', " & _ 
        "stk_avg_cost = '" & weighted_avg_cost & "' where stk_move_id ='" & rs4("stk_move_id") & "' "
        set rs7 = server.CreateObject("adodb.recordset")
        rs7.ActiveConnection = strconnect
        rs7.Source = sql7
        rs7.CursorLocation  = 3
        rs7.Open    
 
    prev_stk_avg_cost = weighted_avg_cost
    prev_total_cost = total_cost
    
   if isNumeric(prev_total_cost) = False then 
    prev_total_cost = 0.00
   end if 

    if isNumeric(balanceBFqty) = False then 
        balanceBFqty = 0.00
    end if

    rs4.movenext
	wend

  end if

    sql4= "select stk_move_id, stk_code_id, stk_date, stk_doc_type, stk_doc_no, stk_qty, stk_bf_qty, stk_purchase_cost, stk_avg_cost,stk_purchase_totalcost, stk_total_cost " & _
    "from tblstock_movement a where stk_code_id = '" & searchvalue & "' order by stk_date"
    set rs4 = server.CreateObject("adodb.recordset")
    rs4.ActiveConnection = strconnect
    rs4.Source = sql4
    rs4.CursorLocation  = 3
    rs4.Open
'end if   

if rs4.eof then
   norecord = "There is no record found."
end if

If Not rs4.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if

'response.Cookies("GAPS")("sqlexcel") = sql4

   			
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
end if


count = count + Showed
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&orderby=" & orderby & "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&Searchor_date=" & Searchor_date & "&ordertype=" & ordertype & "&wh_code=" & wh_code & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to

%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_stock_movement.asp";
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Stock Movement</div>
                          </td>
                      </tr>
                    </table>
                    </td>
                </tr>
           
                <tr>
                    <br />
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> <strong>Stock Movement Report by Product , Store Code </strong></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_stock_movement_excel.asp?searchvalue=<%=searchvalue%>&wh_code=<%=wh_code%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="action_report.asp?type=warehouselocation">
                     <tr>
                  <td height="30" align="left" bgcolor="#FFFFFF"><strong><font color="#000000"><strong>Up to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>" size="15" />
                  <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </td>
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
               
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF">&nbsp;&nbsp;
                   Include All Store &nbsp;&nbsp
                     <select name="whchk" id="whchk">                     
                       <option value="No" <% if whchk = "No" then response.write " selected"%>>No</option>
                      <option value="Yes" <% if whchk = "Yes" then response.write " selected"%>>Yes</option>
                  </select></td>
                  
                <tr><td height="30" bgcolor="#FFFFFF"> <Strong>Stock Code  </Strong>
                    <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                    </td>                    
                   <td width="80%" height="30" align="left" bgcolor="#FFFFFF"><input type="button" name="button2" id="button" value="Clear/Display Report" onclick="javascript:DisplayReport();" /></td></tr>
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
					Response.Write " <a href='rm_rpt_stock_movement.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stock_movement.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                       <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Doc Date</span></strong></font></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Doc Type</span></strong></font></td>
                      <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Doc No</strong></font></td>
                      <td width="50" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><span><strong>Qty</strong></span></font></td>
                      <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty C/F</span></strong></font></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Purchase Unit Cost</span></strong></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total Purchase Cost</span></strong></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Weighted Avg unit Cost</span></strong></td>
                      <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total Cost</span></strong><br/>                          
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


totalvalue = total
'calculation logic

totalcost = 0
	
%>                   
                     <tr bgcolor="<%=nbgcolor%>">
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=chkdate(rs4("stk_date"))%></font></strong></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs4("stk_doc_type")%></font></strong></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs4("stk_doc_no")%></font></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs4("stk_qty")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs4("stk_bf_qty")%> </strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=chknumber2(rs4("stk_purchase_cost"))%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=chknumber2(rs4("stk_purchase_totalcost"))%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2(rs4("stk_avg_cost"))%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=chknumber2(rs4("stk_total_cost"))%></strong></td>
                    </tr>
<%

'if isnumeric(total) then 
'	totalvalue = totalvalue + total
'end if

't_totalqty = t_totalqty + cint(rs("qty")) 
't_totalvalue = t_totalvalue + total 
 
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
              <!--      <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=t_totalqty%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(t_totalvalue)%>&nbsp;</strong></td>
                    </tr>-->
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_stock_movement.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stock_movement.asp?num=" & Showed+row & link & "'> Next >></a>"
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
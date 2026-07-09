<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style3 {
            width: 220px;
        }
        .auto-style4 {
            width: 215px;
        }
    </style>
</head>
<%
searchitem = request("searchitem")
searchitem2 = request("searchitem2")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
inv_status = request("inv_status")

if ordertype = "" then 
   ordertype = "desc"
end if

if orderby = "" then 
   orderby = "tblinvoice.inv_id"
end if

if request("inv_createddate_from") <> "" then
   inv_createddate_from = request("inv_createddate_from")
else
   inv_createddate_from = chkdate(DateAdd("d",-90,date()))
end if

if request("inv_createddate_to") <> "" then
   inv_createddate_to = request("inv_createddate_to")
else
   inv_createddate_to = chkdate(date())
end if

if request("inv_status") <> "" then
   inv_status = request("inv_status")
else

   if request.Cookies("GAPS")("slevel") = "technician" then 
   inv_status = "Submitted"
   else
   inv_status = "Open"
   end if
end if


sql = "SELECT sum(tblinvoice.inv_totalqty) as total_qty, sum(tblinvoice.inv_totalPartsAmt) as total_partsAmt, sum(tblinvoice.inv_labourAmt) as total_labourAmt, " & _
"sum(tblinvoice.inv_transportAmt) as total_transportAmt, sum(tblinvoice.inv_gstAmt) as total_gstAmt, sum(tblinvoice.inv_totalAmt) as total_totalAmt, " & _
"sum(tblinvoice.inv_payment) as total_payment, sum(tblinvoice.inv_cnamount) as total_cnamt, sum(tblinvoice.inv_balance) as total_balance  " & _ 
"FROM tblinvoice left join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code where tblinvoice.inv_id is not null " & _
"and tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(inv_createddate_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(inv_createddate_to) & "' "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if request.Cookies("GAPS")("slevel") = "technician" then 
   sql = sql & " and tblinvoice.inv_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
end if

if inv_status <> "All" and inv_status <> "" then
   sql = sql & " and  tblinvoice.inv_status = '" & inv_status & "' "
end if

if searchitem2 = "ExRecep" and inv_status = "Posted" then 'this option is to sum when excluded already receipted / inv > 0
     sql = sql & " and inv_balance > '0'" 
end if
    
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open

if not rs.eof then 
   total_qty = rs("total_qty")
   total_partsAmt = rs("total_partsAmt")
   total_labourAmt = rs("total_labourAmt")
   total_transportAmt = rs("total_transportAmt")
   total_gstAmt = rs("total_gstAmt")
   total_totalAmt = rs("total_totalAmt")
   total_payment = rs("total_payment")
   total_cnamt = rs("total_cnamt")   
   total_balance = rs("total_balance")   
end if
rs.close

i = 1
sql2 = "SELECT tblinvoice.inv_id, tblinvoice.inv_no, tblreceipt.receipt_no,tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address, tblinvoice.inv_cust_postcode, " & _
"tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id, tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, " & _
"tblinvoice.inv_createddate, tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt, " & _
"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate, tblinvoice.inv_status, inv_approvedby, inv_approveddate, " & _
"tblinvoice.inv_payment, tblinvoice.inv_cnamount, tblinvoice.inv_balance, tblinvoice.inv_payment_type, tblinvoice.inv_chequeno, tblinvoice.inv_payment_remark, tblinvoice.inv_dono, tblinvoice.inv_dodate, " & _
"tbltechnician.tech_name, tbltechnician.tech_tel1 FROM tblinvoice left join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code " & _
"left join tblreceipt on tblinvoice.inv_no = tblreceipt.receipt_inv_no where tblinvoice.inv_id is not null " & _
"and tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(inv_createddate_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(inv_createddate_to) & "' "


if searchvalue <> "" then 
   sql2 = sql2 & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if request.Cookies("GAPS")("slevel") = "technician" then 
   sql2 = sql2 & " and tblinvoice.inv_tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
end if

if inv_status <> "All" and inv_status <> "" then
   sql2 = sql2 & " and  tblinvoice.inv_status = '" & inv_status & "' "
end if

if inv_status = "Posted" and searchitem2 <> "ExRecep" then '100724 this is needed as sometimes posted inv can have cancelled receipt
    sql2 = sql2 & " and tblreceipt.receipt_status = 'Posted'"
end if

if searchitem2 = "ExRecep" then 'this option is to exclude already receipted and FOC items for easier ticking
     sql2 = sql2 & " and inv_balance > '0'"  & " and receipt_no is NULL" 
end if

if ordertype <> "" then
sql2 = sql2 & " order by " & orderby & " " & ordertype
else
sql2 = sql2 & " order by tblinvoice.inv_no desc"
end if
 
'response.write request.Cookies("GAPS")("slevel") & "<br>"

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
link = "&searchitem=" & request("searchitem") & "&searchitem2=" & request("searchitem2") & "&searchvalue=" & request("searchvalue") & "&inv_createddate_from=" & inv_createddate_from & "&inv_createddate_to=" & inv_createddate_to & "&ordertype=" & ordertype & "&orderby=" & orderby   

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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Invoice</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_invoice_view.asp?type=reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td nowrap="nowrap" class="titlegrey1"><strong> Invoice Date <br />
                        </strong></td>
                        <td width="84%"><div align="left"><strong><font color="#000000"><strong>
                          <input name="inv_createddate_from" type="text" id="inv_createddate_from" value="<%=inv_createddate_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.inv_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="inv_createddate_to" type="text" id="inv_createddate_to" value="<%=inv_createddate_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.inv_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015</div></td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td>
                          <select name="searchitem" id="searchitem">
                            <option value="tblinvoice.inv_no"  <% if searchitem = "tblinvoice.inv_no" then response.write " selected" %>>Invoice No</option>
                            <option value="tblinvoice.inv_cust_name" <% if searchitem = "tblinvoice.inv_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblinvoice.inv_cust_tel1" <% if searchitem = "tblinvoice.inv_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tblinvoice.inv_cust_email" <% if searchitem = "tblinvoice.inv_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tblinvoice.inv_tech_code" <% if searchitem = "tblinvoice.inv_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tbltechnician.tech_name" <% if searchitem = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tblinvoice.inv_cust_state" <% if searchitem = "tblinvoice.inv_cust_state" then response.write " selected" %>>Customer State</option>
                           <!-- <option value="tblreceipt.receipt_no" <% if searchitem = "tblreceipt.receipt_no" then response.write " selected" %>>Without Receipt</option>-->
                          </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                            <option value="tblinvoice.inv_no"  <% if orderby = "tblinvoice.inv_no" then response.write " selected" %>>Invoice No</option>
                            <option value="tblinvoice.inv_cust_name" <% if orderby = "tblinvoice.inv_cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblinvoice.inv_cust_tel1" <% if orderby = "tblinvoice.inv_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
                            <option value="tblinvoice.inv_cust_email" <% if orderby = "tblinvoice.inv_cust_email" then response.write " selected" %>>Customer Email</option>
                            <option value="tblinvoice.inv_tech_code" <% if orderby = "tblinvoice.inv_tech_code" then response.write " selected" %>>Technician Code</option>
                            <option value="tbltechnician.tech_name" <% if orderby = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tblinvoice.inv_cust_state" <% if orderby = "tblinvoice.inv_cust_state" then response.write " selected" %>>Customer State</option>
                            <option value="tblinvoice.inv_payment" <% if orderby = "tblinvoice.inv_payment" then response.write " selected" %>>Payment Amount</option>
                          </select>
                          <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                        <input type="submit" name="Submit43" value="Display" />
                        <input type="hidden" name="inv_status" id="inv_status" value="<%=inv_status%>" /></td>
                      </tr>
                        <%if inv_status = "Posted" then %>
                      <tr>
                      <td valign="top" bgcolor="#FFFFFF">Display Receipt</td>
                         <td>
                          <select name="searchitem2" id="searchitem2">
                            <option value="All"  <% if searchitem2 = "All" then response.write " selected" %>>All</option>
                            <option value="ExRecep" <% if searchitem2 = "ExRecep" then response.write " selected" %>>Exclude Receipt</option>
                          </select>
                         </td>
                      </tr>
                        <%end if %>
                    </table>
                  </form></td>
                </tr>
       
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">
                                    <strong>Page</strong> <font color="3366ff"> 
                        <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
                        <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					'Response.Write " <a href='cs_invoice_view.asp?num=" & (j-1) * row & link & "&inv_status=" & inv_status & "' >"& j &"</a>"
                     Response.Write " <a href='rm_invoice_view.asp?num=" & (j-1) * row & link & "&inv_status=" & inv_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_invoice_view.asp?num=" & Showed+row & link & "&inv_status=" & inv_status & "' > Next >></a>"
	End If
	
                    %> 
                  </td>
                </tr>
                <tr>
                <td align="right" valign="top" bgcolor="#FFFFFF"><table border="0" align="right" cellpadding="4" cellspacing="1">
                    <tr>
                     <td nowrap="nowrap" class="auto-style1"><div align="center" class="titlegrey1"> Status :</div> </td>
                       <td <%if inv_status="All" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_invoice_view.asp?inv_status=All<%=link%>"><font color="#FFFFFF"><strong>All</strong></font></a></td>
                      <td <%if inv_status="Open" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_invoice_view.asp?inv_status=Open<%=link%>"><font color="#FFFFFF"><strong>Open</strong></font></a></td>
                      <td <%if inv_status="Submitted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_invoice_view.asp?inv_status=Submitted<%=link%>"><font color="#FFFFFF"> <strong>Submitted</strong></font></a></td>
                       <td <%if inv_status="Posted" then response.write "bgcolor='#475387'" else response.write "bgcolor='#333333'" end if%>><a href="rm_invoice_view.asp?inv_status=Posted<%=link%>"><font color="#FFFFFF"> <strong>Posted</strong></font></a></td>
                    </tr>
                  </table>
                    <div align="right"></div>
                  <div align="right"></div></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Invoice  No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Invoice  Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Customer</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>DO No.</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Tel 1</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Location</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Technician </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Approved By / Date</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice Amt <span>(RM)</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Payment Amt <span>(RM)</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>CN Amt <span>(RM)</span></strong></font></td>
                      <td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Balance Amt <span>(RM)</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Issue. Receipt</strong></font></td>
                    </tr>
<% 
job_totalAmt = 0
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

                   <form name="formorder" id="formorder" method="post" action="action.asp?type=GenReceipt" >
                   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40"><div align="center"><%=j%> </div></td>
                      <td><strong><a href="rm_invoice_new.asp?inv_no=<%=rs("inv_no")%>"> <font color="#0000FF"><%=rs("inv_no")%></font></a></strong><br /><%=rs("inv_status")%>
                      </td>
                      <td nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
                      <td><%=rs("inv_cust_name")%></td>
                      <td align="center"><%=rs("inv_dono")%></td>
                      <td align="left"><%=rs("inv_cust_tel1")%></td>
                      <td align="left"><%=rs("inv_cust_city")%></td>
                      <td><%=rs("tech_name")%></td>
                      <td><%=rs("inv_approvedby")%><br />
                      <%=chkdate(rs("inv_approveddate"))%></td>
                      <td align="right"><%=chknumber2(rs("inv_totalAmt"))%> </td>
                      <td align="right">-<%=chknumber2(rs("inv_payment"))%></td>
                      <td align="right">-<%=chknumber2(rs("inv_cnamount"))%></td>
                      <td align="right"><%=chknumber2(rs("inv_balance"))%></td>
                      <td align="center">
                      <%inv_balance = chknumber2(rs("inv_balance"))
                         receipt_no =  rs("receipt_no")  %>

                       <%if inv_status="Posted" and chknumber2(inv_balance) > 0 then %>
                           <input type="checkbox" name="Genreceipt" id="Genreceipt" value="<%=rs("inv_no")%>"/>
                      <%elseif inv_status="Posted" and chknumber2(inv_balance) = 0 and not isnull(receipt_no) then%>
                           <%=rs("receipt_no")%>
                      <%elseif inv_status="Posted" and chknumber2(inv_balance) = 0 and isnull(receipt_no) then%> 
                           <%response.write "Parts FOC" %>                       
                      <%end if%>
                      </td>
                    </tr>
<%
inv_totalAmt = inv_totalAmt + rs("inv_totalAmt")
inv_cnamount = inv_cnamount + rs("inv_cnamount")
inv_payment = inv_payment + rs("inv_payment")
inv_balance_totalamt = inv_balance_totalamt + rs("inv_balance")
    
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
                    
                    <tr>
                      <td height="30" colspan="9" align="right" bgcolor="#CCCCCC"><strong> Total</strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong> <%=chknumber2(inv_totalAmt)%> </strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong>-<%=chknumber2(inv_payment)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong>-<%=chknumber2(inv_cnamount)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_balance_totalamt)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong>

                      <%if inv_status="Posted" then %>
                        <input type="submit" name="button2" value="Generate Single Receipt" formaction="action.asp?type=GenReceipt" class="auto-style3"  />
                      <%end if%> </strong>
                        
                        <%if inv_status="Posted" then %>
                          <% if searchitem = "tblinvoice.inv_tech_code" then '060724 - radha use only code as name may not be unique%> 
                                <% if searchvalue <> "" then %>
                                    <input type="submit" name="button2" value="Generate Combined Receipt" formaction="action.asp?type=GenReceiptAll" class="auto-style3"/>
                                <%end if%>
                          <%end if%>
                        <%end if%> </td>
                    </tr>
                    <tr>
                      <td height="30" colspan="9" align="right" bgcolor="#CCCCCC"><strong>Grand Total</strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong> <%=chknumber2(total_totalAmt)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong>-<%=chknumber2(total_payment)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong>-<%=chknumber2(total_cnamt)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(total_balance)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong></strong></td>
                    </tr>
                      </form>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">
                                    <strong>Page</strong> <font color="3366ff"> 
                        <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
                        <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_invoice_view.asp?num=" & (j-1) * row & link & "&inv_status=" & inv_status & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_invoice_view.asp?num=" & Showed+row & link & "&inv_status=" & inv_status & "'> Next >></a>"
	End If
	
                    %> </td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
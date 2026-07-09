<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = "Over"
inv_status = request("inv_status")
inv_no = request("inv_no")

if ordertype = "" then 
   ordertype = "desc"
end if

if inv_status = "" then 
   inv_status = "Posted"
end if


if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-30,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if request("job_tech_code") <> "" then
   job_tech_code = replace(request("job_tech_code"), " ", "")
   arrjob_tech_code = split(job_tech_code,",")
   job_tech_code = replace(job_tech_code, ",", "','")
   listjob_tech_code = listjob_tech_code & job_tech_code
else
   listjob_tech_code = ""
   arrjob_tech_code = split("0,0",",")
end if

function checkTechlList(strv)
for k = 0 to ubound(arrjob_tech_code)
    if arrjob_tech_code(k) = strv then 
	   checkTechlList = true
	   exit for
	else
	   checkTechlList = false
	end if
next
end function

i = 1

sql2 = "SELECT sum(tblinvoice.inv_totalqty) as totalqty, sum(tblinvoice.inv_totalPartsAmt*0.9433962264) totalPartsAmt, sum(tblinvoice.inv_labourAmt*0.9433962264) as totallabourAmt,  " & _
		"sum(tblinvoice.inv_transportAmt*0.9433962264) as totaltransportAmt, sum(tblinvoice.inv_gstAmt) as totalgstAmt,  sum(tblinvoice.inv_totalAmt) as totaltotalAmt, " & _
		"sum(tblinvoice.inv_payment) as totalpayment, sum(tblinvoice.inv_balance) as totalbalance " & _
		"FROM tblinvoice left join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code " & _
		"where tblinvoice.inv_id is not null " & _
		"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code in ( '" & job_tech_code & "') "
	end if
	
	if inv_no <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_no like '%" & inv_no & "%' "
	end if
	
	if inv_status <> "All" and inv_status<>"" then 
	   sql2 = sql2 & " and tblinvoice.inv_status = '" & inv_status & "' "
	end if
	
'response.write sql2
'response.End()		
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   totalqty = rs("totalqty")
   totalPartsAmt = rs("totalPartsAmt")
   totallabourAmt = rs("totallabourAmt")
   totaltransportAmt = rs("totaltransportAmt")
   totalgstAmt = rs("totalgstAmt")
   totaltotalAmt = rs("totaltotalAmt")
   totalpayment = rs("totalpayment")
   totalbalance = rs("totalbalance")
end if
rs.close


sql2 = "select sum(tblcn.cn_totalAmt) as totalcnamount from tblcn inner join tblinvoice on tblcn.cn_inv_no=tblinvoice.inv_no " & _
       "left join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code " & _
       "where tblcn.cn_status='Posted' and " & _
	   "tblcn.cn_date > '" & ChkDateYYYYMMDD(job_date_from) & "' and tblcn.cn_date < '" & ChkDateYYYYMMDD(job_date_to) & "' "
		
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code in ( '" & job_tech_code & "') "
	end if
	
	if inv_no <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_no like '%" & inv_no & "%' "
	end if
	
	if inv_status <> "All" and inv_status<>"" then 
	   sql2 = sql2 & " and tblinvoice.inv_status = '" & inv_status & "' "
	end if
	
'response.write sql2
'response.End()		
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   totalcnamount = rs("totalcnamount")
end if
rs.close


	   
sql2 = "SELECT tblinvoice.inv_id, tblinvoice.inv_no, tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address,  " & _
		"tblinvoice.inv_cust_postcode, tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id,  " & _
		"tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, tblinvoice.inv_createddate,  " & _
		"tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt,  " & _
		"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate,  " & _
		"tblinvoice.inv_status, tblinvoice.inv_approvedby, tblinvoice.inv_approveddate, tblinvoice.inv_remark, tblinvoice.inv_postedby, tblinvoice.inv_posteddate, " & _
		"tblinvoice.inv_payment, " & _
		"(select sum(tblcn.cn_totalAmt) as cn_totalAmt from tblcn where tblcn.cn_inv_no=tblinvoice.inv_no and tblcn.cn_status='Posted' and tblcn.cn_date > '" & ChkDateYYYYMMDD(job_date_from) & "' and tblcn.cn_date < '" & ChkDateYYYYMMDD(job_date_to) & "') " & _
		"as inv_cnamount, tblinvoice.inv_balance, tblinvoice.inv_payment_type, tblinvoice.inv_chequeno, tblinvoice.inv_payment_remark,  " & _
		"tbltechnician.tech_name " & _
		"FROM tblinvoice left join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code " & _
		"where tblinvoice.inv_id is not null " & _
		"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code in ( '" & job_tech_code & "') "
	end if
	
	if inv_no <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_no like '%" & inv_no & "%' "
	end if
	
	if inv_status <> "All" and inv_status<>"" then 
	   sql2 = sql2 & " and tblinvoice.inv_status = '" & inv_status & "' "
	end if
	
       sql2 = sql2 & " order by tblinvoice.inv_no  "

response.Cookies("GAPS")("sqlexcel") = sql2
'response.write sql2
'response.End()
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

link = "&job_tech_type=" & job_tech_type & "&Searchor_date=" & Searchor_date & "&orderby=" & orderby & "&ordertype=" & ordertype & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & "&inv_status=" & inv_status   
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Sales Invoice Summary (By Month)</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1">&nbsp;</td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_salesanalysis_bymonth_summary_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_salesanalysis_bymonth_summary.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Invoice Date<br />
                        </strong></td>
                        <td colspan="4"><div align="left"><strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </div></td>
                      </tr>
                      <tr>
                        <td class="titlegrey1">Type </td>
                        <td width="14%"><span class="titlegrey1">Technician</span></td>
                        <td width="23%" align="center"><span class="titlegrey1">Invoice Status</span></td>
                        <td width="24%" align="center"><span class="titlegrey1">Invoice No.</span></td>
                        <td width="23%" rowspan="2"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="job_tech_type" id="job_tech_type">
                          <option value="">All</option>
                          <option value="CF" <%if job_tech_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                          <option value="WH" <%if job_tech_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                        </select></td>
                        <td width="14%"><span class="titlegrey1">
                          <select name="job_tech_code" size="6" multiple="multiple" id="job_tech_code">
                            <option value="" <%if job_tech_code="" then response.write " selected"%>>All Technicians</option>
                            <%			
				sql1 = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC' order by tech_code "	
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if checkTechlList(rs1("tech_code")) then
					  response.write "<option value='" & rs1("tech_code") & "' selected>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs1("tech_code") & "'>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  end if 					  
				rs1.movenext
				wend
				rs1.close					
				%>
                          </select>
                        </span></td>
                        <td width="23%" align="center" valign="top">
                          <select name="inv_status" id="inv_status">
                            <option value="All">All</option>
                            <option value="Open" <%if inv_status="Open" then response.write " selected"%>>Open</option>
                            <option value="Submitted" <%if inv_status="Submitted" then response.write " selected"%>>Submitted</option>
                            <option value="Posted" <%if inv_status="Posted" then response.write " selected"%>>Posted</option>                           
                          </select>
                       </td>
                        <td width="24%" align="center" valign="top"><label for="inv_no"></label>
                        <input name="inv_no" type="text" id="inv_no" value="<%=inv_no%>" size="15" maxlength="20" /></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_salesanalysis_bymonth_summary.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_salesanalysis_bymonth_summary.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="5" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Invoice No</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Date</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td width="80" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer <br />
                      </span></strong></font></td>
                      <td width="80" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Technician<span><br />
                      </span></strong></font></td>
                      <td width="20" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Qty</strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Amt</strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GST Amt</span></strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Parts Amt</span></strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Labour Amt</span></strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Transport Amt</span></strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Payment Amt</strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CN Amt</strong></font></td>
                      <td width="40" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Balance Amt</strong></font></td>
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
                      <td width="5" height="40" align="center"><%=j%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_invoice_new.asp?inv_no=<%=rs("inv_no")%>" target="_blank"><%=rs("inv_no")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("inv_status"))%></td>
                      <td width="80" align="left"><%=rs("inv_cust_code") & " " & left(rs("inv_cust_name"),15) %></td>
                      <td width="80" align="left"><%=rs("inv_tech_code") &  "<br>" & left(rs("tech_name"),15)%></td>
                      <td align="center" nowrap="nowrap"><strong> <%=rs("inv_totalqty")%></strong></td>
                      <td width="40" align="right" nowrap="nowrap"><strong> <%=chknumber2(rs("inv_totalAmt"))%></strong></td>
                      <td width="40" align="right"><strong><%=chknumber2(rs("inv_gstAmt"))%></strong></td>
                      <td width="40" align="right"><strong><%=chknumber2(rs("inv_totalPartsAmt")*0.9433962264)%></strong></td>
                      <td width="40" align="right"><strong><%=chknumber2(rs("inv_labourAmt")*0.9433962264)%></strong></td>
                      <td width="40" align="right"><strong><%=chknumber2(rs("inv_transportAmt")*0.9433962264)%></strong></td>
                      <td width="40" align="right"><strong><%=chknumber2(rs("inv_payment"))%></strong></td>
                      <td width="40" align="right"><strong><%=chknumber2(rs("inv_cnamount"))%></strong></td>
                      <td width="40" align="right"><strong><%=chknumber2(rs("inv_balance"))%></strong></td>
                    </tr>
<%
inv_totalAmt = inv_totalAmt + ccur(rs("inv_totalAmt"))
inv_gstAmt = inv_gstAmt + ccur(rs("inv_gstAmt"))
inv_totalPartsAmt = inv_totalPartsAmt + ccur(rs("inv_totalPartsAmt")*0.9433962264)
inv_labourAmt = inv_labourAmt + ccur(rs("inv_labourAmt")*0.9433962264)
inv_transportAmt = inv_transportAmt + ccur(rs("inv_transportAmt")*0.9433962264)

inv_payment = inv_payment + ccur(rs("inv_payment"))

if not isnull(rs("inv_cnamount")) then 
inv_cnamount = inv_cnamount + ccur(rs("inv_cnamount"))
end if

inv_balance = inv_balance + ccur(rs("inv_balance"))

count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
                    
                   
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <!--Open-->
                      <td width="40" align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_totalAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_gstAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_totalPartsAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_labourAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_transportAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_payment)%></strong></td>
                      <td width="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_cnamount)%></strong></td>
                      <td width="40" align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_balance)%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="7" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td width="40" align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(totaltotalAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalgstAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalPartsAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totallabourAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totaltransportAmt)%></strong></td>
                      <td width="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalpayment)%></strong></td>
                      <td width="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalcnamount)%></strong></td>
                      <td width="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalbalance)%></strong></td>
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
					Response.Write " <a href='rm_rpt_salesanalysis_bymonth_summary.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_salesanalysis_bymonth_summary.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
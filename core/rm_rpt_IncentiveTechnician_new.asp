<!-- #include file="header.asp" -->

<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
tech_type = request("tech_type")
job_tech_code = request("job_tech_code")

'response.write request("jobmonth")
'response.write request("jobyear")

if ordertype = "" then 
   ordertype = "desc"
end if

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

if request("jobmonth") <> "" then
   jobmonth = trim(request("jobmonth"))
else
   jobmonth = month(date())
end if

if request("jobyear") <> "" then
   jobyear = trim(request("jobyear"))
else
   jobyear = year(date())
end if

i = 1	

sql2 = "SELECT A.inc_tech_code,A.inc_month,A.inc_year,A.inc_invoice_no,A.inc_jobno,A.inc_modelno,B.md_desc,A.inc_received_amt,A.inc_labor_charge_amt,A.inc_labor_overwriting, " & _
"A.inc_labor_payout,A.inc_part_charge,A.inc_part_overwriting,A.inc_part_payout,A.inc_total_payout,A.inc_net_received FROM tbltech_incentive A " & _
"INNER JOIN tblmodel B " & _
"ON  A.inc_modelno = B.md_code " & _
"where A.inc_tech_code= '" & job_tech_code & "' and A.inc_month = '" & jobmonth & "'and A.inc_year='" & jobyear & "' order by A.inc_invoice_no"

response.Cookies("GAPS")("sqlexcel") = sql2

set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open
if rs1.eof then
   norecord = "There is no record found."
end if

'response.write sql2
If Not rs1.EOF Then
    'response.write rs1("inc_invoice_no")
if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs1.RecordCount
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

link = "&job_tech_type=" & job_tech_type & " &job_tech_code=" & job_tech_code & "&tech_type=" & tech_type &"&job_date_from=" & job_date_from & "&job_date_to =" & job_date_to & "&jobyear=" & jobyear & "&jobmonth=" & jobmonth

%>  
        <tr>
          <td><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Over-Warranty Incentive Technician Report</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td class="titleblue1">&nbsp;</td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_IncentiveTechnician_new_excel.asp?sparepartPer=<%=sparepartPer%>&labourPer=<%=labourPer%>&job_tech_code=<%=job_tech_code%>&job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_type=<%=tech_type%>&jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="action_report.asp?type=monthtechIncentive">
                    <table width="90%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="titlegrey1">Invoice Date</td>
                        <td colspan="3"><strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong>  eg: 21-May-2015 </td>
                      </tr>
                        <tr>
                            <td class="titlegrey1">Incentive for Month</td>
                            <td valign="top" bgcolor="#FFFFFF">
                    <select name="jobyear" id="jobyear">
                        <option value="2021"<%if jobyear="2021" then response.write " selected"%>>2021</option>
                        <option value="2022"<%if jobyear="2022" then response.write " selected"%>>2022</option>
                        <option value="2023"<%if jobyear="2023" then response.write " selected"%>>2023</option>
                        <option value="2024"<%if jobyear="2024" then response.write " selected"%>>2024</option>
                        <option value="2025"<%if jobyear="2025" then response.write " selected"%>>2025</option>
                        <option value="2026"<%if jobyear="2026" then response.write " selected"%>>2026</option>
                    </select>
                    <select name="jobmonth" id="jobmonth">
                      <option value="1" <%if jobmonth="1" then response.write " selected"%>>Jan</option>
                      <option value="2" <%if jobmonth="2" then response.write " selected"%>>Feb</option>
                      <option value="3" <%if jobmonth="3" then response.write " selected"%>>Mar</option>
                      <option value="4" <%if jobmonth="4" then response.write " selected"%>>Apr</option>
                      <option value="5" <%if jobmonth="5" then response.write " selected"%>>May</option>
                      <option value="6" <%if jobmonth="6" then response.write " selected"%>>Jun</option>
                      <option value="7" <%if jobmonth="7" then response.write " selected"%>>Jul</option>
                      <option value="8" <%if jobmonth="8" then response.write " selected"%>>Aug</option>
                      <option value="9" <%if jobmonth="9" then response.write " selected"%>>Sep</option>
                      <option value="10" <%if jobmonth="10" then response.write " selected"%>>Oct</option>
                      <option value="11" <%if jobmonth="11" then response.write " selected"%>>Nov</option>
                      <option value="12" <%if jobmonth="12" then response.write " selected"%>>Dec</option>
                    </select></td></tr>
                      <tr>
                        <td width="16%" class="titlegrey1">Technician</td>
                        <td class="titlegrey1"><span class="titlegrey1">
                          <select name="job_tech_code" id="job_tech_code">
                           <option value="0"></option>
                            <%			
				'sql3 = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC'"
                sql3 = "select distinct tbljob.job_tech_code, tbltechnician.tech_name FROM tbljob INNER JOIN tblinvoice ON tbljob.job_inv_no = tblinvoice.inv_no INNER JOIN " & _
                "tbltechnician ON tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null and tbljob.job_status='Posted' and tbljob.job_actual_wrty_status = 'Over' " & _
                "and tbljob.job_submitforclaims='Yes' and tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "'"
                set rs3 = server.CreateObject("adodb.recordset")
				rs3.Open sql3,strconnect,3,3,&H0001
                while Not rs3.EOF
					  if rs3("job_tech_code") = job_tech_code then
					  response.write "<option value='" & rs3("job_tech_code") & "' selected>" & rs3("job_tech_code") & " - " & rs3("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs3("job_tech_code") & "'>" & rs3("job_tech_code") & " - " & rs3("tech_name")  & "</option>"
					  end if 					  
				rs3.movenext
				wend
				rs3.close					
				%>      </select>
                        </span></td>
                        <td class="titlegrey1"></td>
                        <td><span class="titlegrey1">
                          <input type="submit" name="button2" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>                    
                    </table>
                  </form></td>
                </tr>               
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_IncentiveTechnician_new.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_IncentiveTechnician_new.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If      
 %>

                </td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="97" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap"    bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span> Invoice No</span></strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>JS No</span></strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong><span>Model</span></strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Total (A) <br/>Contractor Amt <br />Received <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Labor <br/> Charge <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Overwriting <br/> % (Labour)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Labour <br/> Payout <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Spare Part <br/> Charge <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Overwriting <br/>  % (Spare <br/>Part)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Sparepart <br/> Payout <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Total (B) <br/> Payout <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap"  bgcolor="#CCCCCC" class="style1"><font color="#030303"><strong>Net Riegen Marketing <br/>Received (RM)<br/> Total (A) - (B)</strong></font></td>
                    </tr>   
<%
gtotal_received = 0
gtotal_labourCharge = 0
gtotal_labourpayout = 0
gtotal_sparepartscharge = 0
gtotal_partsamount = 0
gtotal_Bpayout = 0
gtotal_netReceived = 0

if not rs1.eof then
rs1.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

'for grand total figures
gtotal_received = gtotal_received + rs1("inc_received_amt")
gtotal_labourCharge = gtotal_labourCharge + rs1("inc_labor_charge_amt")
gtotal_labourpayout = gtotal_labourpayout + rs1("inc_labor_payout")
gtotal_partsamount = gtotal_partsamount + rs1("inc_part_charge")
gtotal_sparepartscharge = gtotal_sparepartscharge + rs1("inc_part_payout")
gtotal_Bpayout =  gtotal_Bpayout + rs1("inc_total_payout")
gtotal_netReceived = gtotal_netReceived + rs1("inc_net_received")

%>                    
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center" valign="top" nowrap="nowrap"><%=j%></td>
                      <td align="left" valign="top" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs1("inc_invoice_no")%></font></strong></td>
                      <td align="left" valign="top" nowrap="nowrap"><strong><font color="#0000FF"><%=rs1("inc_jobno")%></font></strong></td>
                      <td align="left" valign="top" nowrap="nowrap"><%=rs1("md_desc")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_received_amt"))%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_labor_charge_amt"))%></td>
					  <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("inc_labor_overwriting")%>%</td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="#F5B041"><%=chknumber2(rs1("inc_labor_payout"))%></td> 
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_part_charge"))%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("inc_part_overwriting")%>%</td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="#F5B041"><%=chknumber2(rs1("inc_part_payout"))%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_total_payout"))%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_net_received"))%></td>
                    </tr>      
 <%
count = count + 1 
i = i + 1
rs1.MoveNext
Next
rs1.Close

%>
    
<tr bgcolor="<%=nbgcolor%>">
  <td height="40" colspan="4" align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Grand Total</strong></td>
  <td height="40" align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_received)%></strong></td>
  <td height="40" align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_labourCharge)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"></td>
  <td align="right" nowrap="nowrap" bgcolor="#F5B041"><strong><%=chknumber2(gtotal_labourpayout)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_partsamount)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"></td>
  <td align="right" nowrap="nowrap" bgcolor="#F5B041"><strong><%=chknumber2(gtotal_sparepartscharge)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#2ECC71"><strong><%=chknumber2(gtotal_Bpayout)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_netReceived)%></strong></td>
</tr>    
</table></td> 
<tr>
<td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
</tr>
<tr>
<td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
</tr>
<tr><td align="right"><strong>G/Total <%=chknumber2(gtotal_Bpayout)%></strong>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
<tr><td align="right">(Riegen to pay out)&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
<tr>
<td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
</tr>
</tr>
<tr>
<td>&nbsp;&nbsp;&nbsp;&nbsp;<strong><u>OVER WARRANTY INCENTIVE FEES</u></strong></td>
</tr>
<tr>
<td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
</tr>
<tr>
<td width="15%">&nbsp;&nbsp;&nbsp;&nbsp;a) Riegen Marketing Received  :&nbsp;<%=chknumber2(gtotal_received)%> </td>
</tr>
<tr>
<td width="15%">&nbsp;&nbsp;&nbsp;&nbsp;b) Contractor Overiding  : &nbsp;<%=chknumber2(gtotal_Bpayout)%></td>
</tr>
<tr>
<td width="15%">&nbsp;&nbsp;&nbsp;&nbsp;c) Balance (A-B)  :&nbsp;<%=chknumber2(gtotal_netReceived)%></td>
</tr>
 <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_IncentiveTechnician_new.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_IncentiveTechnician_new.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->

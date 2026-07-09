<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = request("job_actual_wrty_status")
tech_type = request("tech_type")
sparepartPer = request("sparepartPer")
labourPer = request("labourPer")
job_tech_code = request("job_tech_code")

if sparepartPer = "" then 
   sparepartPer = "25"
end if

if labourPer = "" then 
   labourPer = "75"
end if


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

sql2 = "SELECT  sum(tbljob.job_totalPartsAmt) as job_totalPartsAmt, sum(tbljob.job_totallabourAmt) as job_totallabourAmt " & _
		"FROM         tbljob INNER JOIN tblinvoice ON tbljob.job_inv_no = tblinvoice.inv_no INNER JOIN " & _
        "tbltechnician ON tbljob.job_tech_code = tbltechnician.tech_code " & _
		"where tbljob.job_id is not null and tbljob.job_status='Posted' " & _
		"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	
		if job_actual_wrty_status <> "" then 
		   sql2 = sql2 & " and tbljob.job_actual_wrty_status like '" & job_actual_wrty_status & "' "
		end if
		
		if job_tech_code <> "All" and job_tech_code<>"" then 
		   sql2 = sql2 & " and tbljob.job_tech_code = '" & job_tech_code & "' "
        'response.write sql2
        'response.End()	
		end if
		
		if tech_type <> "All" and tech_type<>"" then 
		   sql2 = sql2 & " and tbltechnician.tech_type = '" & tech_type & "' "
            'response.write sql2
            'response.End()	
		end if
	
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   total_PartsAmt = rs("job_totalPartsAmt")
   total_labourAmt = rs("job_totallabourAmt")
   
   total_NetPartsAmt = chknumber2(total_PartsAmt * (sparepartPer/100))
   total_NetlabourAmt = chknumber2(total_labourAmt * (labourPer/100))
end if
rs.close

i = 1	
sql2 = "SELECT    tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_linkedcode, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, tbljob.job_cust_postcode, " & _
		"tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email, tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_remark, tbljob.job_createddate,  " & _
		"tbljob.job_createdby, tbljob.job_submittedby, tbljob.job_submitteddate, tbljob.job_pendingby, tbljob.job_pendingdate, tbljob.job_doneby, tbljob.job_donedate, tbljob.job_postedby,  " & _
		"tbljob.job_posteddate, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby, tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type,  " & _
		"tbljob.job_SN_no, tbljob.job_Model, tbljob.job_model_desc, tbljob.job_faulty_reason_cs, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time, " & _ 
		"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate, tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model,  " & _
		"tbljob.job_tech_model_desc, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN, tbljob.job_tech_faulty_code, tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status,  " & _
		"tbljob.job_tech_product_collectdate, tbljob.job_tech_service_date, tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo, tbljob.job_tech_logby,  " & _
		"tbljob.job_tech_logdate, tbljob.job_hq_remark, tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
		"tbljob.inv_totalqty, tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_logbyhq, tbljob.job_logdatehq, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
		"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_reassignedby, tbljob.job_reassigndate, tbljob.job_inv_no, tbljob.job_inv_date, tbljob.job_do_no, tbljob.job_do_date,  " & _
		"tbljob.job_cancelledby, tbljob.job_cancelleddate, tbltechnician.tech_type, tblinvoice.inv_date, tblinvoice.inv_no " & _
		"FROM         tbljob INNER JOIN tblinvoice ON tbljob.job_inv_no = tblinvoice.inv_no INNER JOIN " & _
        "tbltechnician ON tbljob.job_tech_code = tbltechnician.tech_code " & _
		"where tbljob.job_id is not null and tbljob.job_status='Posted' " & _
		"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
		if job_actual_wrty_status <> "" then 
		   sql2 = sql2 & " and tbljob.job_actual_wrty_status like '" & job_actual_wrty_status & "' "
		end if
		
		if job_tech_code <> "All" and job_tech_code<>"" then 
		   sql2 = sql2 & " and tbljob.job_tech_code = '" & job_tech_code & "' "
		end if
		
		if tech_type <> "All" and tech_type<>"" then 
		   sql2 = sql2 & " and tbltechnician.tech_type = '" & tech_type & "' "
		end if
	
       sql2 = sql2 & " order by tblinvoice.inv_no  "


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

link = "&job_tech_type=" & job_tech_type & "&Searchor_date=" & Searchor_date & "&orderby=" & orderby & "&ordertype=" & ordertype & "&job_actual_wrty_status=" & job_actual_wrty_status & "&tech_type=" & tech_type & "&sparepartPer=" & sparepartPer & "&labourPer=" & labourPer & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to

%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Incentive Technician   Report</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1"> Month Tech Report<br />                        <br /></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_IncentiveTechnician_excel.asp?sparepartPer=<%=sparepartPer%>&labourPer=<%=labourPer%>&job_tech_code=<%=job_tech_code%>&job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_type=<%=tech_type%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_IncentiveTechnician.asp">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
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
                        <td class="titlegrey1">Percentage</td>
                        <td colspan="3" class="titlegrey1">Spare Part
                          <input name="sparepartPer" type="text" id="sparepartPer" value="<%=sparepartPer%>" size="10" maxlength="10" />
                        % Labour
                        <input name="LabourPer" type="text" id="LabourPer" value="<%=LabourPer%>" size="10" maxlength="10" />
                        %</td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1">Warranty</td>
                        <td class="titlegrey1">Technician</td>
                        <td class="titlegrey1">Technician Type</td>
                        <td><span class="titlegrey1">
                          <input type="submit" name="button2" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="job_actual_wrty_status" id="job_actual_wrty_status">                          
                          <option value="Over" <%if job_actual_wrty_status="Over" then response.write " selected"%>>Over</option>
                          <option value="Under" <%if job_actual_wrty_status="Under" then response.write " selected"%>>Under</option>
                          <option value="">All</option>
                        </select></td>
                        <td valign="top" class="titlegrey1"><span class="titlegrey1">
                          <select name="job_tech_code" id="job_tech_code">
                            <option value="" <%if job_tech_code="" then response.write " selected"%>>All Technicians</option>
                            <%			
				sql2 = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC' order by tech_code "	
                set rs2 = server.CreateObject("adodb.recordset")
				rs2.Open sql2,strconnect,3,3,&H0001
                while Not rs2.EOF
					  if rs2("tech_code") = job_tech_code then
					  response.write "<option value='" & rs2("tech_code") & "' selected>" & rs2("tech_code") & " - " & rs2("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs2("tech_code") & "'>" & rs2("tech_code") & " - " & rs2("tech_name")  & "</option>"
					  end if 					  
				rs2.movenext
				wend
				rs2.close					
				%>
                          </select>
                        </span></td>
                        <td valign="top" class="titlegrey1"><label for="select">
                          <select name="tech_type" id="tech_type">
                            <option value="">All</option>
                            <option value="IC" <%if tech_type="IC" then response.write " selected"%>>IC</option>
                            <option value="TPC" <%if tech_type="TPC" then response.write " selected"%>>TPC</option>
                            <option value="IHT" <%if tech_type="IHT" then response.write " selected"%>>IHT</option>
                            <option value="IHC" <%if tech_type="IHC" then response.write " selected"%>>IHC</option>
                          </select>
                        </label></td>
                        <td>&nbsp;</td>
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
					Response.Write " <a href='rm_rpt_IncentiveTechnician.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_IncentiveTechnician.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Invoice No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice </strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>JS No</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Spare Part <br />
                      / Labour</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total <br />
                        Contractor <br />
                      Amt Received</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total Riegen Marketing <br />
                      AMT Received</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Percentage<br />
                        %
                      </span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Model</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total Payout<br />
                      (RM)</strong></font></td>
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

technician_tech_type = rs("tech_type")
job_jobcode = rs("job_code")
job_count = rs("job_count")

if technician_tech_type = "TPC" then 
	if (job_count = 1) then
		labourPer = "75"
	end if
	if (job_count > 1) then
		labourPer = "100"
	end if		
end if 

if technician_tech_type = "IHT" then 
	if (job_count = 1) then
		labourPer = "50"
	end if 
	if (job_count > 1) then
		labourPer = "100"
	end if	
end if 

if technician_tech_type = "IHC" then 		
	if (job_count = 1) then
		labourPer = "0"
	end if 
	if (job_count > 1) then
		labourPer = "0"
	end if	
end if 
    
totalsparepart = rs("job_totalPartsAmt") * (sparepartPer/100)
totallabourpart = rs("job_totallabourAmt") * (labourPer/100)
%>                    
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" rowspan="2" align="center" valign="top" nowrap="nowrap"><%=j%></td>
                      <td rowspan="2" align="left" valign="top" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("job_inv_no")%></font></strong></td>
                      <td rowspan="2" align="left" valign="top" nowrap="nowrap"><strong><font color="#0000FF"><%=chkdate(rs("inv_date"))%></font></strong></td>
                      <td rowspan="2" align="left" valign="top" nowrap="nowrap">&nbsp;</td>
                      <td rowspan="2" align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs("job_code")%></td>
                      <td rowspan="2" align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs("job_tech_code")%></td>
					  
                      <td align="left" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><strong>Spare Part</strong></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="#E5E5E5">&nbsp;</td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs("job_totalPartsAmt"))%></td>
                      <td align="right" valign="top" nowrap="nowrap"><%=sparepartPer%></td>
                      <td align="center" valign="top"><%=rs("job_tech_model_desc")%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(totalsparepart)%></td>
                    </tr>
                    <tr bgcolor="<%=nbgcolor%>">
                      <td align="left" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><strong>Labour</strong></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="#E5E5E5">&nbsp;</td>
                      <td align="right" valign="top" nowrap="nowrap"><%=chknumber2(rs("job_totallabourAmt"))%></td>
                      <td align="right" valign="top" nowrap="nowrap"><%=labourPer%></td>
                      <td align="right" valign="top" nowrap="nowrap">&nbsp;</td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(totallabourpart)%></td>
                    </tr>
<%
t_all = t_all + (ccur(totallabourpart) + ccur(totalsparepart))
t_RMreceived = t_RMreceived + (rs("job_totalPartsAmt") + rs("job_totallabourAmt"))
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
    

<tr bgcolor="<%=nbgcolor%>">
                      <td height="40" colspan="7" align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td height="40" align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(t_RMreceived)%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(t_all)%></strong></td>
                    </tr>
<tr bgcolor="<%=nbgcolor%>">
  <td height="40" colspan="7" align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Grand Total</strong></td>
  <td height="40" align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(total_PartsAmt+total_labourAmt)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(ccur(total_NetPartsAmt)+ccur(total_NetlabourAmt))%></strong></td>
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
					Response.Write " <a href='rm_rpt_IncentiveTechnician.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_IncentiveTechnician.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
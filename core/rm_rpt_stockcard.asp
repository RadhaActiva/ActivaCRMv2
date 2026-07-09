<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = "Over"


if ordertype = "" then 
   ordertype = "desc"
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
sql2 = "SELECT tblinvoice.inv_id, tblinvoice.inv_no, tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address,  " & _
		"tblinvoice.inv_cust_postcode, tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id,  " & _
		"tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, tblinvoice.inv_createddate,  " & _
		"tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt,  " & _
		"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate,  " & _
		"tblinvoice.inv_status, tblinvoice.inv_approvedby, tblinvoice.inv_approveddate, tblinvoice.inv_remark, tblinvoice.inv_postedby, tblinvoice.inv_posteddate, " & _
		"tbltechnician.tech_name " & _
		"FROM tblinvoice left join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code " & _
		"where tblinvoice.inv_id is not null " & _
		"and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code in ( '" & job_tech_code & "') "
	end if
	
       sql2 = sql2 & " order by tblinvoice.inv_no  "

'response.write sql2
'response.End()
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

link = "&job_tech_type=" & job_tech_type & "&Searchor_date=" & Searchor_date & "&orderby=" & orderby & "&ordertype=" & ordertype & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Stock Card Report (Stock Movement)</div></td>
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
                      <td width="20%" align="center" class="titlegrey1"><img src="images/excel.jpg" width="57" height="21" /></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_stockcard.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="20" colspan="4" nowrap="nowrap" class="titlegrey1"><strong><strong>
                            Date Range
                            <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                            <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                            <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1"> Store</td>
                        <td width="14%"><span class="titlegrey1">Item</span></td>
                        <td width="24%" align="center">&nbsp;</td>
                        <td rowspan="2"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="job_tech_code" size="6" multiple="multiple" id="job_tech_code">
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
                        </select></td>
                        <td width="14%"><span class="titlegrey1">
                          <select name="job_tech_code2" size="6" multiple="multiple" id="job_tech_code2">
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
                        <td width="24%" align="center" valign="top">
                        <!--
                          <select name="job_actual_wrty_status" id="job_actual_wrty_status">
                            <option value="">All</option>
                            <option value="Over" <%if job_actual_wrty_status="Over" then response.write " selected"%>>Over</option>
                            <option value="Under" <%if job_actual_wrty_status="Under" then response.write " selected"%>>Under</option>
                        </select>
                        --></td>
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
					Response.Write " <a href='rm_rpt_stockcard.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Invoice No</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Invoice Date</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer <br />
                      </span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Technician<span><br />
                      </span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Qty</strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Amt</strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GST Amt</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Parts Amt</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Labour Amt</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Transport Amt</span></strong></font></td>
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
                      <td height="40" align="center"><%=j%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("inv_no")%></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
                      <td align="left"><%=rs("inv_cust_code") & " " & rs("inv_cust_name") %></td>
                      <td align="left"><%=rs("inv_tech_code") &  " " & rs("tech_name")%></td>
                      <td align="center" nowrap="nowrap"><strong> <%=rs("inv_totalqty")%></strong></td>
                      <td align="right" nowrap="nowrap"><strong> <%=chknumber2(rs("inv_totalAmt"))%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("inv_gstAmt"))%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("inv_totalPartsAmt"))%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("inv_labourAmt"))%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("inv_transportAmt"))%></strong></td>
                    </tr>
<%
inv_totalAmt = inv_totalAmt + cint(rs("inv_totalAmt"))
inv_gstAmt = inv_gstAmt + cint(rs("inv_gstAmt"))
inv_totalPartsAmt = inv_totalPartsAmt + cint(rs("inv_totalPartsAmt"))
inv_labourAmt = inv_labourAmt + cint(rs("inv_labourAmt"))
inv_transportAmt = inv_transportAmt + cint(rs("inv_transportAmt"))
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
                    
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="6" align="right" bgcolor="#999999"><strong>Total</strong></td>
                      <!--Open-->
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(inv_totalAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(inv_gstAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(inv_totalPartsAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(inv_labourAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(inv_transportAmt)%></strong></td>
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
					Response.Write " <a href='rm_rpt_stockcard.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_stockcard.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
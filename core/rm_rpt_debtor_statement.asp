<!-- #include file="header.asp" -->
<%
doc_type = request("doc_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
inv_status = request("inv_status")
inv_no = request("inv_no")
cust_code = request("cust_code")
job_tech_code = request("job_tech_code")

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


i = 1

sql2 = "Select v_debtorStmt.CN, sum(v_debtorStmt.cn_totalAmt) as totaltotalAmt from v_debtorStmt inner join tblinvoice on v_debtorStmt.cn_inv_no=tblinvoice.inv_no " & _
	   "and v_debtorStmt.cn_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and v_debtorStmt.cn_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code = '" & job_tech_code & "' "
	end if
	
	if inv_no <> "" then 
	   sql2 = sql2 & " and v_debtorStmt.cn_inv_no like '%" & inv_no & "%' "
	end if
	
	if inv_status <> "All" and inv_status<>"" then 
	   sql2 = sql2 & " and v_debtorStmt.cn_status = '" & inv_status & "' "
	end if

	if doc_type <> "" then 
	   sql2 = sql2 & " and v_debtorStmt.CN like '%" & doc_type & "%' "
	end if

	if cust_code <> "" then 
	   sql2 = sql2 & " and v_debtorStmt.cn_cust_code like '%" & cust_code & "%' "
	end if	
		
sql2 = sql2 & " group by v_debtorStmt.CN"	
	
'response.write sql2
'response.End()		
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
while not rs.eof

if rs("CN") = "INV" then 
   tInvAmt = rs("totaltotalAmt")
elseif rs("CN") = "Pay" then    
   tPayAmt = rs("totaltotalAmt")
elseif rs("CN") = "CN" then    
   tCNAmt = rs("totaltotalAmt")   
end if
rs.movenext
wend
rs.close
	   
sql2 = "Select v_debtorStmt.cn_id, v_debtorStmt.cn_no, v_debtorStmt.cn_date, v_debtorStmt.cn_cust_code, v_debtorStmt.cn_cust_name, " & _
       "v_debtorStmt.cn_job_code, v_debtorStmt.cn_inv_no, v_debtorStmt.cn_totalAmt, v_debtorStmt.cn_status, v_debtorStmt.CN, " & _
       "tblinvoice.inv_job_code, tblinvoice.inv_tech_code from v_debtorStmt inner join tblinvoice on v_debtorStmt.cn_inv_no=tblinvoice.inv_no " & _
	   "and v_debtorStmt.cn_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and v_debtorStmt.cn_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code = '" & job_tech_code & "' "
	end if
	
	if inv_no <> "" then 
	   sql2 = sql2 & " and v_debtorStmt.cn_inv_no like '%" & inv_no & "%' "
	end if
	
	if inv_status <> "All" and inv_status<>"" then 
	   sql2 = sql2 & " and v_debtorStmt.cn_status = '" & inv_status & "' "
	end if

	if doc_type <> "" then 
	   sql2 = sql2 & " and v_debtorStmt.CN like '%" & doc_type & "%' "
	end if

	if cust_code <> "" then 
	   sql2 = sql2 & " and v_debtorStmt.cn_cust_code like '%" & cust_code & "%' "
	end if		
	
sql2 = sql2 & " order by v_debtorStmt.cn_inv_no"		

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

link = "&job_tech_type=" & job_tech_type & "&Searchor_date=" & Searchor_date & "&orderby=" & orderby & "&ordertype=" & ordertype & "&job_date_from=" & job_date_from & "&job_date_to=" & job_date_to & "&inv_status=" & inv_status & "&inv_no=" & inv_no & "&cust_code=" & cust_code & "&job_tech_code=" & job_tech_code
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Debtor Statement</div></td>
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
                      <td width="20%" align="right" class="titlegrey1"><a href="rm_rpt_debtor_statement_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_debtor_statement.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="15%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Document Date<br />
                        </strong></td>
                        <td colspan="5"><div align="left"><strong><font color="#000000"><strong>
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
                        <td width="22%"><span class="titlegrey1">Technician</span></td>
                        <td width="15%" align="left"><span class="titlegrey1"> Status</span></td>
                        <td width="13%" align="left"><span class="titlegrey1">Cust. Code</span></td>
                        <td width="13%" align="left">&nbsp;</td>
                        <td width="22%" rowspan="2" align="right"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="doc_type" id="doc_type">
                          <option value="">All</option>
                          <option value="INV" <%if doc_type="INV" then response.write " selected"%>>INV</option>
                          <option value="CN" <%if doc_type="CN" then response.write " selected"%>>CN</option>
                          <option value="PAY" <%if doc_type="PAY" then response.write " selected"%>>PAY</option>
                        </select></td>
                        <td width="22%"><span class="titlegrey1">
                          <select name="job_tech_code" id="job_tech_code" style="width:200px">
                            <option value="" <%if job_tech_code="" then response.write " selected"%>>All Technicians</option>
                            <%			
				sql1 = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC' order by tech_code "	
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if rs1("tech_code") = job_tech_code then
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
                        <td width="15%" align="left" valign="top">
                          <select name="inv_status" id="inv_status">
                            <option value="All">All</option>
                            <option value="Open" <%if inv_status="Open" then response.write " selected"%>>Open</option>
                            <option value="Submitted" <%if inv_status="Submitted" then response.write " selected"%>>Submitted</option>
                            <option value="Posted" <%if inv_status="Posted" then response.write " selected"%>>Posted</option>                           
                          </select>
                       </td>
                        <td width="13%" align="left" valign="top"><label for="cust_code"></label>
                        <input name="cust_code" type="text" id="cust_code" value="<%=cust_code%>" /></td>
                        <td width="13%" align="left" valign="top">&nbsp;</td>
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
					Response.Write " <a href='rm_rpt_debtor_statement.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_debtor_statement.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="4%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="16%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Document No</span></strong></font></td>
                      <td width="11%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Doc. Date</span></strong></font></td>
                      <td width="7%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td width="11%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Cust. Code<br />
                      </span></strong></font></td>
                      <td width="16%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cust. Name</strong></font></td>
                      <td width="10%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Tech. Code</strong></font></td>
                      <td width="7%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Type<span><br />
                      </span></strong></font></td>
                      <td width="9%" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Debit</strong></font></td>
                      <td width="9%" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Credit</span></strong></font></td>
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
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF">
                      <% if rs("cn") = "INV" then %>
                      <a href="rm_invoice_new.asp?inv_no=<%=rs("cn_no")%>" target="_blank"><%=rs("cn_no")%></a>
                      <% elseif rs("cn") = "CN" then %>
                      <a href="rm_cn_new.asp?cn_no=<%=rs("cn_no")%>" target="_blank"><%=rs("cn_no")%></a>
                      <% elseif rs("cn") = "Pay" then %>
                      <a href="rm_receipt_new.asp?receipt_no=<%=rs("cn_no")%>" target="_blank"><%=rs("cn_no")%></a>
                      <% End if %>
                      
                      </font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("cn_status"))%></td>
                      <td align="left"><%=rs("cn_cust_code") %></td>
                      <td align="left"><%=rs("cn_cust_name") %></td>
                      <td align="left"><%=rs("inv_tech_code") %></td>
                      <td align="center"><%=rs("CN")%></td>
                      <td align="right" nowrap="nowrap"><strong> 
					  <% 
					  if rs("cn") = "INV" then 
					  response.write chknumber2(rs("cn_totalAmt"))
					  end if
					  %></strong></td>
                      <td align="right"><strong><%
					  if rs("cn") = "Pay" or rs("cn") = "CN" then 
					  response.write chknumber2(rs("cn_totalAmt"))
					  end if
					  %></strong></td>
                    </tr>
<%

if rs("cn") = "INV" then    
inv_debitAmt = inv_debitAmt + ccur(rs("cn_totalAmt"))
end if

if (rs("cn") = "Pay" or rs("cn") = "CN") then 
inv_creditAmt = inv_creditAmt + ccur(rs("cn_totalAmt"))
end if

count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
                    
                   
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="8" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <!--Open-->
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_debitAmt)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(inv_creditAmt)%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="8" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(tInvAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(tPayAmt+tCNAmt)%></strong></td>
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
					Response.Write " <a href='rm_rpt_debtor_statement.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_debtor_statement.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
<!-- #include file="header.asp" -->
<%
orderby = request("orderby")
ordertype = request("ordertype")
cn_status = request("cn_status")
cn_no = request("cn_no")

if ordertype = "" then 
   ordertype = "desc"
end if

if cn_status = "" then 
   cn_status = "Posted"
end if


if request("cn_createddate_from") <> "" then
   cn_createddate_from = request("cn_createddate_from")
else
   cn_createddate_from = chkdate(DateAdd("d",-30,date()))
end if

if request("cn_createddate_to") <> "" then
   cn_createddate_to = request("cn_createddate_to")
else
   cn_createddate_to = chkdate(date())
end if

i = 1

sql2 = "SELECT  sum(cn_labourAmt) as cn_labourAmt, " & _
       "sum(cn_transportAmt) as cn_transportAmt, " & _
	   "sum(cn_gstAmt) as cn_gstAmt, " & _
	   "sum(cn_totalAmt) as cn_totalAmt, " & _
	   "sum(cn_totalqty) as cn_totalqty " & _
	   "FROM tblcn where cn_id is not null " & _
		"and  cn_date >= '" & ChkDateYYYYMMDD(cn_createddate_from) & "' and cn_date <= '" & ChkDateYYYYMMDD(cn_createddate_to) & "' "

	if cn_no <> "" then 
	   sql2 = sql2 & " and cn_no like '%" & cn_no & "%' "
	end if
	
	if cn_status <> "All" and cn_status<>"" then 
	   sql2 = sql2 & " and cn_status = '" & cn_status & "' "
	end if
	
'response.write sql2
'response.End()		
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   total_labourAmt = rs("cn_labourAmt")
   total_transportAmt = rs("cn_transportAmt")
   total_gstAmt = rs("cn_gstAmt")
   total_totalAmt = rs("cn_totalAmt")*0.9433962264
   total_totalqty = rs("cn_totalqty")
end if
rs.close


sql2 = "SELECT tblcn.cn_id, tblcn.cn_no, tblcn.cn_status, tblcn.cn_date, tblcn.cn_inv_no, tblcn.cn_inv_date, tblcn.cn_cust_code, tblcn.cn_cust_name, tblcn.cn_cust_address, tblcn.cn_cust_postcode, " & _
	  "tblcn.cn_cust_state, tblcn.cn_cust_state_id, tblcn.cn_cust_city, tblcn.cn_cust_city_id, tblcn.cn_cust_email, tblcn.cn_cust_tel1, tblcn.cn_cust_tel2, tblcn.cn_createddate, tblcn.cn_createdby,  " & _
	  "tblcn.cn_job_code, tblcn.cn_do_no, tblcn.cn_invoice_no, tblcn.cn_totalqty, tblcn.cn_totalPartsAmt, tblcn.cn_remark, tblcn.cn_labourAmt, tblcn.cn_transportAmt, tblcn.cn_gstAmt, tblcn.cn_totalAmt,  " & _
	  "tblcn.cn_emailsent, tblcn.cn_emailsentdate, tblcn.cn_returnedby, tblcn.cn_returneddate, tblcn.cn_submittedby, tblcn.cn_submitteddate, tblcn.cn_doneby, tblcn.cn_donedate, tblcn.cn_postedby,  " & _
	  "tblcn.cn_posteddate, tblcn.cn_cancelledby, tblcn.cn_cancelleddate,tblinvoice.inv_no, tblinvoice.inv_date " & _
	  "FROM tblcn left join tblinvoice on tblcn.cn_inv_no=tblinvoice.inv_no where tblcn.cn_id is not null " & _
	  "and  tblcn.cn_date >= '" & ChkDateYYYYMMDD(cn_createddate_from) & "' and tblcn.cn_date <= '" & ChkDateYYYYMMDD(cn_createddate_to) & "' "
		
	if cn_no <> "" then 
	   sql2 = sql2 & " and tblcn.cn_no like '%" & cn_no & "%' "
	end if
	
	if cn_status <> "All" and cn_status<>"" then 
	   sql2 = sql2 & " and tblcn.cn_status = '" & cn_status & "' "
	end if
	
       sql2 = sql2 & " order by tblcn.cn_no "

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

link = "&orderby=" & orderby & "&ordertype=" & ordertype & "&cn_status=" & cn_status & "&cn_no=" & cn_no & "&cn_createddate_from=" & cn_createddate_from & "&cn_createddate_to=" & cn_createddate_to 
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Credit Note (CN) Summary </div></td>
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
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_cn_summary_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_cn_summary.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> CN Date<br />
                        </strong></td>
                        <td colspan="3"><div align="left"><strong><font color="#000000"><strong>
                          <input name="cn_createddate_from" type="text" id="cn_createddate_from" value="<%=cn_createddate_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.cn_createddate_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="cn_createddate_to" type="text" id="cn_createddate_to" value="<%=cn_createddate_to%>"
                                            size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.cn_createddate_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </div></td>
                      </tr>
                      <tr>
                        <td class="titlegrey1">CN Status</td>
                        <td><span class="titlegrey1">
                          <select name="cn_status" id="cn_status">
                            <option value="All">All</option>
                            <option value="Open" <%if cn_status="Open" then response.write " selected"%>>Open</option>
                            <option value="Submitted" <%if cn_status="Submitted" then response.write " selected"%>>Submitted</option>
                            <option value="Done" <%if cn_status="Done" then response.write " selected"%>>Done</option>
                            <option value="Posted" <%if cn_status="Posted" then response.write " selected"%>>Posted</option>
                            <option value="Cancel" <%if cn_status="Cancel" then response.write " selected"%>>Cancel</option>
                          </select>
                        </span></td>
                        <td width="24%" align="center">&nbsp;</td>
                        <td width="23%" rowspan="2"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1">CN No.</td>
                        <td>
                          <span class="titlegrey1">
                          <input name="cn_no" type="text" id="cn_no" value="<%=cn_no%>" size="15" maxlength="20" />
                          </span></td>
                        <td width="24%" align="center" valign="top"><label for="inv_no"></label></td>
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
					Response.Write " <a href='rm_rpt_cn_summary.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_cn_summary.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>CN No</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> CN Date</span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer <br />
                      </span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>DO No<span><br />
                      </span></strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CN Qty</strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CN Amt</strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>GST Amt</span></strong></font></td>
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
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_cn_new.asp?cn_no=<%=rs("cn_no")%>" target="_blank"><%=rs("cn_no")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("cn_status"))%></td>
                      <td align="left"><%=rs("cn_cust_code") & " " & rs("cn_cust_name") %></td>
                      <td align="left"><%=rs("cn_do_no")%></td>
                      <td align="left"><%=rs("inv_no")%></td>
                      <td align="left"><%=chkdate(rs("inv_date"))%></td>
                      <td align="center" nowrap="nowrap"><strong> <%=rs("cn_totalqty")%></strong></td>
                      <td align="right" nowrap="nowrap"><strong> <%=chknumber2(ccur(rs("cn_totalAmt")))%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("cn_gstAmt"))%></strong></td>
                    </tr>
<%
cn_totalqty = cn_totalqty + ccur(rs("cn_totalqty"))
cn_totalAmt = cn_totalAmt + ccur(rs("cn_totalAmt"))
cn_gstAmt = cn_gstAmt + ccur(rs("cn_gstAmt"))

count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
                    
                   
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="9" align="right" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <!--Open-->
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(cn_totalAmt)%></strong></td>
                      <td align="right" bgcolor="#CCCCCC"><strong><%=chknumber2(cn_gstAmt)%></strong></td>
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="9" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(total_totalAmt)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(total_gstAmt)%></strong></td>
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
					Response.Write " <a href='rm_rpt_cn_summary.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_cn_summary.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
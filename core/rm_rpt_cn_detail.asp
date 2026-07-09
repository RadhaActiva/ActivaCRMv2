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
   total_totalAmt = rs("cn_totalAmt")
   total_totalqty = rs("cn_totalqty")
end if
rs.close

i = 1	
sql2 = "SELECT cnd_id, cnd_cn_no, cnd_inv_no, cnd_job_code,  " & _
		"cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt,  " & _
		"cnd_discounttype, cnd_netcost, cnd_subtotal, " & _
		"cn_date, cn_status, cn_inv_no, cn_cust_code, cn_cust_name, " & _
		"cn_job_code, cn_do_no, inv_no, inv_date, md_averageecost " & _
		"FROM v_creditnoteCost where cnd_id is not null " & _
	    "and  cn_date >= '" & ChkDateYYYYMMDD(cn_createddate_from) & "' and cn_date <= '" & ChkDateYYYYMMDD(cn_createddate_to) & "' " 
	
	if cn_no <> "" then 
	   sql2 = sql2 & " and cnd_cn_no like '%" & cn_no & "%' "
	end if
	
	if cn_status <> "All" and cn_status<>"" then 
	   sql2 = sql2 & " and cn_status = '" & cn_status & "' "
	end if
	
       sql2 = sql2 & " order by cnd_cn_no  "

response.Cookies("GAPS")("sqlexcel") = sql2
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Credit Note (CN)  Detail </div></td>
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
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_cn_detail_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_cn_detail.asp?type=searchdata">
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
                        <td width="14%"><span class="titlegrey1">
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
                        <td rowspan="2"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1">CN No.</td>
                        <td width="14%"><span class="titlegrey1">
                          <input name="cn_no" type="text" id="cn_no" value="<%=cn_no%>" size="15" maxlength="20" />
                        </span></td>
                        <td width="24%" align="center" valign="top">&nbsp;</td>
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
					Response.Write " <a href='rm_rpt_cn_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_cn_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>CN No</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> CN Date</span></strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Customer Code<br />
                        </span></strong></font><font color="#FFFFFF"><strong><span><br />
                      </span></strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Customer Name</strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>DO No</strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Code</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Description</font></strong></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Cost</font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong> CN<br />
                      </strong></font><font color="#FFFFFF"><strong>Qty</strong></font></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>CN<br /> 
                      Amt</span></strong></font></td>
                   <!--   <td height="30" align="right" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Gst<br />
                      Amt</font></strong></td>  20/8/23 Abolition of GST-->
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
                      <td height="40" align="center" nowrap="nowrap"><%=j%>.</td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><a href="rm_cn_new.asp?cn_no=<%=rs("cnd_cn_no")%>" target="_blank"><%=rs("cnd_cn_no")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=(rs("cn_status"))%></td>
                      <td align="left"><%=rs("cn_cust_code")%></td>
                      <td align="left"><%=rs("cn_cust_name")%></td>
                      <td align="left"><%=rs("cn_do_no")%></td>
                      <td align="left"><%=rs("inv_no")%></td>
                      <td align="left" nowrap="nowrap"><%=chkdate(rs("inv_date"))%></td>
                      <td align="left" nowrap="nowrap"><%=rs("cnd_partcode")%></td>
                      <td align="left"><%=rs("cnd_desc")%></td>
                      <td align="right"><strong><%=round(rs("md_averageecost"),2)%></strong></td>
                      <td align="right" nowrap="nowrap"><strong> <%=rs("cnd_qty")%></strong></td>
                      <td align="right"><strong><%=round((rs("cnd_subtotal")*0.9433962264),2) + round((rs("cnd_subtotal")*0.0566037735849057),2) %></strong></td>
                      <!--<td align="right"><strong><%=round((rs("cnd_subtotal")*0.0566037735849057),2)%></strong></td>-->
                    </tr>
<%
cnd_qty = cnd_qty + ccur(rs("cnd_qty"))
cnd_subtotal = cnd_subtotal + round((rs("cnd_subtotal")*0.9433962264),2) + round((rs("cnd_subtotal")*0.0566037735849057),2)
' 20/08/23 abolition of gst - commented the lines below
'cnd_subtotal = cnd_subtotal + round((rs("cnd_subtotal")*0.9433962264),2) 
'cnd_gstAmt = cnd_gstAmt + round((rs("cnd_subtotal")*0.0566037735849057),2) 
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
                    
                   
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="12" align="right" bgcolor="#999999"><strong>Total</strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=(cnd_qty)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(cnd_subtotal)%></strong></td>
                     <!-- <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(cnd_gstAmt)%></strong></td>-->
                    </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="12" align="right" bgcolor="#999999"><strong>Grand Total</strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=(total_totalqty)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(total_totalAmt)%></strong></td>
                   <!--   <td align="right" bgcolor="#999999"><strong><%=chknumber2(total_totalAmt-total_gstAmt)%></strong></td>-->
                     <!-- <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(total_gstAmt)%></strong></td>-->
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
					Response.Write " <a href='rm_rpt_cn_detail.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_cn_detail.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
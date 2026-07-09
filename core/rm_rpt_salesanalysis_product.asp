<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = "Over"
inv_status = request("inv_status")

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
		
sql2 = "SELECT  tblmodel.md_category, tblinvoice_detail.invd_partcode, tblinvoice_detail.invd_parttype, tblinvoice_detail.invd_desc, " & _
	   "sum(tblinvoice_detail.invd_subtotal*0.0566037735849057) as totalBilled, sum(tblinvoice_detail.invd_subtotalcost*0.0566037735849057) as totalcost,  " & _
	   "sum(tblinvoice_detail.invd_subtotal*0.0566037735849057) - sum(tblinvoice_detail.invd_subtotalcost*0.0566037735849057) as margin " & _
	   "FROM tblinvoice_detail inner join tblmodel on tblinvoice_detail.invd_partcode=tblmodel.md_code  " & _	
	   "inner join tblinvoice on tblinvoice_detail.invd_inv_no=tblinvoice.inv_no " & _
	   "where tblinvoice_detail.invd_partcode is not null " & _
	   "and  tblinvoice.inv_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tblinvoice.inv_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
	
	if job_tech_code <> "" then 
	   sql2 = sql2 & " and tblinvoice.inv_tech_code in ( '" & job_tech_code & "') "
	end if

	if inv_status <> "All" and inv_status<>"" then 
	   sql2 = sql2 & " and tblinvoice.inv_status = '" & inv_status & "' "
	end if
	
	sql2 = sql2 & " group by tblinvoice_detail.invd_partcode, tblinvoice_detail.invd_parttype, tblinvoice_detail.invd_desc,  tblmodel.md_category "

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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Sales Analysis by Product Detail </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> BI Module Sales Analysis Report Group by Category, Product</td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_salesanalysis_product_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_salesanalysis_product.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Invoice Date<br />
                        </strong></td>
                        <td colspan="3"><div align="left"><strong><font color="#000000"><strong>
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
                        <td width="24%" align="center"><span class="titlegrey1">Invoice Status</span></td>
                        <td rowspan="2"><span class="titlegrey1">
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
                        <td width="24%" align="center" valign="top">
                          <select name="inv_status" id="inv_status">
                            <option value="All">All</option>
                            <option value="Open" <%if inv_status="Open" then response.write " selected"%>>Open</option>
                            <option value="Submitted" <%if inv_status="Submitted" then response.write " selected"%>>Submitted</option>
                            <option value="Posted" <%if inv_status="Posted" then response.write " selected"%>>Posted</option>
                            <option value="CN" <%if inv_status="CN" then response.write " selected"%>>CN</option>
                          </select>
                       </td>
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
					Response.Write " <a href='rm_rpt_salesanalysis_product.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_salesanalysis_product.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td height="30" align="center" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Category</strong></font></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Code</font></strong></td>
                      <td height="30" align="left" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Item Description</font></strong></td>
                      <td height="30" align="right" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Amount billed <br />
                        (excluded GST)
                      </font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Cost Price<br />
                      (excluded GST) </font></strong></td>
                      <td align="right" valign="top" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Margin Amt</span></strong></font><br />
                      <strong><font color="#FFFFFF">(excluded GST) </font></strong></td>
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
                      <td align="left"><%=rs("invd_parttype")%></td>
                      <td align="left" nowrap="nowrap"><%=rs("invd_partcode")%></td>
                      <td align="left"><%=rs("invd_desc")%></td>
                      <td align="right"><strong><%=chknumber2(rs("totalBilled"))%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("totalcost"))%></strong></td>
                      <td align="right"><strong><%=chknumber2(rs("margin"))%></strong></td>
                    </tr>
<%
totalBilled = totalBilled + ccur(rs("totalBilled"))
totalcost = totalcost + ccur(rs("totalcost"))
margin = margin + ccur(rs("margin"))
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close

%>
                    
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Total</strong></td>
                      <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalBilled)%></strong></td>
                      <td height="40" align="right" bgcolor="#999999"><strong><%=chknumber2(totalcost)%></strong></td>
                      <td align="right" bgcolor="#999999"><strong><%=chknumber2(margin)%></strong></td>
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
					Response.Write " <a href='rm_rpt_salesanalysis_product.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_salesanalysis_product.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
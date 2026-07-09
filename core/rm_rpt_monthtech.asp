<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = request("job_actual_wrty_status")

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

%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Month Tech Report</div></td>
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
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_monthtech_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="action_report.asp?type=rpt_monthtech">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="titlegrey1"><strong>Job Month</strong></td>
                        <td colspan="2"><strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </td>
                      </tr>
                      <tr>
                        <td width="16%" class="titlegrey1">Warranty</td>
                        <td class="titlegrey1">Technician</td>
                        <td><span class="titlegrey1">
                          <input type="submit" name="button2" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="job_actual_wrty_status" id="job_actual_wrty_status">
                          <option value="">All</option>
                          <option value="Over" <%if job_actual_wrty_status="Over" then response.write " selected"%>>Over</option>
                          <option value="Under" <%if job_actual_wrty_status="Under" then response.write " selected"%>>Under</option>
                        </select></td>
                        <td valign="top" class="titlegrey1"><span class="titlegrey1">
                          <select name="job_tech_code" size="6" multiple="multiple" id="job_tech_code">
                            <option value="" <%if job_tech_code="" then response.write " selected"%>>All Technicians</option>
                            <%			
				sql = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC' order by tech_code "	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if checkTechlList(rs("tech_code")) then
					  response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                          </select>
                        </span></td>
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td colspan="3" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Accepted</font></strong></td>
                      <td colspan="3" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Done</font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Posted</font></strong></td>
                    </tr>
                    <tr>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Technician</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>1-3 days</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>4-6 days</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>7 days &amp; above</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>1-3 days</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>4-6 days</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>7 days &amp; above</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1">&nbsp;</td>
                    </tr>
                    
<%
i = 1


sql2 = "SELECT id, tech_code, tech_name, pending_1to3d, pending_4to6d, pending_7above, done_1to3d, done_4to6d, done_7above, posted_qty " & _
	   "FROM tblrpr_monthtech where id is not null order by tech_code "

'response.write sql2
'response.End()
i = 1	   
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
                    
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center" nowrap="nowrap"><%=i%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs1("tech_code")%> - <%=rs1("tech_name")%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><a href="javascript:popup('rm_rpt_monthtech_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_code=<%=rs1("tech_code")%>&stype=Accepted1-3','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs1("pending_1to3d")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><a href="javascript:popup('rm_rpt_monthtech_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_code=<%=rs1("tech_code")%>&stype=Accepted4-6','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs1("pending_4to6d")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><a href="javascript:popup('rm_rpt_monthtech_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_code=<%=rs1("tech_code")%>&stype=Accepted7above','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs1("pending_7above")%></a></strong></td>
                      <td align="center" nowrap="nowrap"><strong><a href="javascript:popup('rm_rpt_monthtech_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_code=<%=rs1("tech_code")%>&stype=Done1-3','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs1("done_1to3d")%></a></strong></td>
                      <td align="center" nowrap="nowrap"><strong><a href="javascript:popup('rm_rpt_monthtech_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_code=<%=rs1("tech_code")%>&stype=Done4-6','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs1("done_4to6d")%></a></strong></td>
                      <td align="center" nowrap="nowrap"><strong><a href="javascript:popup('rm_rpt_monthtech_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_code=<%=rs1("tech_code")%>&stype=Done7above','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs1("done_7above")%></a></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#E5E5E5"><strong><a href="javascript:popup('rm_rpt_monthtech_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&tech_code=<%=rs1("tech_code")%>&stype=Posted','cb19','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs1("posted_qty")%></a></strong></td>
                    </tr>
<%
pending_1to3d = pending_1to3d + rs1("pending_1to3d")
pending_4to6d = pending_4to6d + rs1("pending_4to6d")
pending_7above = pending_7above + rs1("pending_7above")
done_1to3d = done_1to3d + rs1("done_1to3d")
done_4to6d = done_4to6d + rs1("done_4to6d")
done_7above = done_7above + rs1("done_7above") 
posted_qty = posted_qty + rs1("posted_qty") 

i = i + 1
rs1.movenext
wend
rs1.close
%> 

<tr bgcolor="<%=nbgcolor%>">
                      <td height="40" colspan="2" align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=pending_1to3d%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=pending_4to6d%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=pending_7above%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=done_1to3d%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=done_4to6d%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=done_7above%></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=posted_qty%></strong></td>
                    </tr>
                    
                  </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
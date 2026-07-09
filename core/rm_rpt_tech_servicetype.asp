<!-- #include file="header.asp" -->
<%

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
   jobmonth = request("jobmonth")
else
   jobmonth = month(date())
end if

if request("jobyear") <> "" then
   jobyear = request("jobyear")
else
   jobyear = year(date())
end if

if request("job_tech_code") <> "" then
   job_tech_code = request("job_tech_code")
else
   job_tech_code = "All"
end if

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Service Type Summary</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1"> Service Type Summary by month                            
                        <label for="select"></label></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_tech_servicetype_excel.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_tech_servicetype.asp">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Job Month<br />
                        </strong></td>
                        <td width="38"><div align="left"><strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong><span class="titlegrey1">
                          <select name="job_tech_code" id="job_tech_code">
                            <option value="" <%if job_tech_code="" then response.write " selected"%>>All Technicians</option>
                            <%			
				sql = "SELECT tech_code, tech_name FROM tbltechnician where tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC' order by tech_code "	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if rs("tech_code") = job_tech_code then
					  response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                          </select>
                          </span><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                          </span></div></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
               
                <tr>
                  <td width="43%" valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="9%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="75%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Service Status</span></strong></font></td>
                      <td width="16%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
                    </tr>
                    

 <%
i = 1
sql1 = "select job_status, count(job_id) as totalcnt from tbljob where " & _
       "job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 

if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	
	   
sql1 = sql1 & " group by job_status  "
	   'response.write sql1
Response.Cookies("GAPS")("sqlexcel") = sql1

totalcnt = 0
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>

 <tr>
   <td height="40" align="center"><%=i%></td>
   <td align="left"><strong><%=rs1("job_status")%></strong></td>
   <td align="center" nowrap="nowrap" bgcolor="#F3F3F3">
   <a href="javascript:popup('rm_rpt_tech_servicetype_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&job_status=<%=rs1("job_status")%>&jobtype=status&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')">
   <strong><%=rs1("totalcnt")%></strong></a></td>
 </tr>
 
 
 <%
 if rs1("job_status") = "Posted" then 
 %>
	 <%
    withouappointment = 0
    sql1 = "select count(job_id) from tbljob where job_status='Posted' and job_tech_code='resolved_no_appt' " & _
	       "and job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
    withouappointment =  selectid(sql1)
    %>
   <tr>
   <td height="40" align="center">&nbsp;</td>
   <td align="left"><strong> - Resolved Issue without Appointment = <%=withouappointment%></strong></td>
   <td align="center" nowrap="nowrap" bgcolor="#F3F3F3">&nbsp;</td>
   </tr>
   
    <%
	withppointment=0
	sql1 = "select count(job_id) from tbljob where job_status='Posted' and job_tech_code<>'resolved_no_appt' " & _
	       "and job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
    withppointment =  selectid(sql1)
    %>
   <tr>
   <td height="40" align="center">&nbsp;</td>
   <td align="left"><strong>- Resolved Issue with Techician Appointment = <%=withppointment%></strong></td>
   <td align="center" nowrap="nowrap" bgcolor="#F3F3F3">&nbsp;</td>
   </tr>
 <%
 end if
 %>
 
  <%
totalcnt = totalcnt + cint(rs1("totalcnt"))
i = i + 1
rs1.movenext
wend
rs1.close
totalcnt = totalcnt 
%>
 


 
                    
                    
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="2" align="right" bgcolor="#FFFFFF"><strong>Total Job Issued</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
					  <a href="javascript:popup('rm_rpt_tech_servicetype_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&job_status=all&jobtype=status','cb18','scrollbars=yes,resizable=yes,width=600,height=500')">
					  <%=totalcnt%></a></strong></td>
                    </tr>
                  </table></td>
                  <td width="57%" valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="8%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="80%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Service Type</span></strong></font></td>
                      <td width="12%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Qty</span></strong></font></td>
                    </tr>
                    <%
i = 1
sql1 = "select job_reportedby, count(job_id) as totalcnt from tbljob where " & _
       "job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' "  

if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	

sql1 = sql1 & " group by job_reportedby  "
Response.Cookies("GAPS")("sqlexcel") = sql1

totalcnt = 0
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
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
                      <td height="40" align="center"><%=i%></td>
                      <td align="left"><strong> <font color="#0000FF"><%=rs1("job_reportedby")%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
					  <a href="javascript:popup('rm_rpt_tech_servicetype_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&job_reportedby=<%=rs1("job_reportedby")%>&jobtype=job_reportedby&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')">
					  <%=rs1("totalcnt")%></a></font></strong></strong></td>
                    </tr>
                    <%

totalcnt = totalcnt + cint(rs1("totalcnt"))
i = i + 1
rs1.movenext
wend
rs1.close
%>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="2" align="right" bgcolor="#FFFFFF"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
					  <a href="javascript:popup('rm_rpt_tech_servicetype_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&job_reportedby=all&jobtype=job_reportedby&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')">
					  <%=totalcnt%></a></strong></td>
                    </tr>
                  </table></td>
                </tr>
                 <tr>
                   <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                   <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                 </tr>
                 <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td height="30" colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
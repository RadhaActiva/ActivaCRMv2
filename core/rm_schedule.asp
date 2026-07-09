<!-- #include file="header.asp" -->
<%
if request("job_date") <> "" then 
   job_date = request("job_date")
else
   job_date = date()
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
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left">View Schedule </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
                        <tr>
                          <td colspan="2" class="style21"><font size="4"><strong>Technician Schedule</strong></font></td>
                        </tr>
                        <tr>
                          <td class="style21"><form action="rm_schedule.asp?searchitem=<%=searchitem%>&amp;<%=searchvalue%>&amp;<%=searchvalue%>&amp;md_type=<%=md_type%>&amp;formname=<%=formname%>&amp;fieldname=<%=fieldname%>" method="post" name="form1" id="form1">
                            <font color="#000000"><strong> Job Date
                              <input name="job_date" type="text" id="job_date" value="<%=chkdate(job_date)%>" size="12" maxlength="20" />
                              <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font><strong>Techinician</strong>
                            <select name="job_tech_code" id="job_tech_code">
                              <option value="All">All</option>
                              <%			
				sql1 = "SELECT tech_id, tech_code, tech_name FROM tbltechnician where tech_status = 'Y' and (tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC') "	
				if Request.Cookies("GAPS")("slevel") = "technician" then 
				sql1 = sql1 & "and tech_code='" & Request.Cookies("GAPS")("job_tech_code") & "'"
				end if
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  if (job_tech_code) = (rs1("tech_code")) then
					  response.write "<option value='" & rs1("tech_code") & "' selected>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs1("tech_code") & "'>" & rs1("tech_code") & " - " & rs1("tech_name")  & "</option>"
					  end if 					  
				rs1.movenext
				wend
				rs1.close					
				%>
                            </select>
                            <input type="submit" name="Submit" value="Submit" />
                          </form></td>
                          <td class="style21"><font color="#FFFFFF">
                            <input type="button" name="button2" id="button7" value="&lt;" onclick="document.location.href='rm_schedule.asp?job_date=<%=chkdate(dateadd("d",-1,job_date))%>'" />
                            <input type="button" name="button5" id="button5" value="Today" onclick="document.location.href='rm_schedule.asp?job_date=<%=chkdate(date())%>'" />
                            <input type="button" name="button6" id="button6" value="&gt;" onclick="document.location.href='rm_schedule.asp?job_date=<%=chkdate(dateadd("d",1,job_date))%>'" />
                          </font></td>
                        </tr>
                        <tr>
                          <td colspan="2" align="left" valign="top">&nbsp;</td>
                        </tr>
                        <tr>
                          <td colspan="2" valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
                            <tr valign="top" bgcolor="#88c0a7">
                              <td width="3%"><strong>No.</strong></td>
                              <td width="12%"><strong>Tech Code</strong></td>
                              <td width="12%"><strong>Tech Name</strong></td>
                              <td><strong>Area</strong></td>
                              <td class='tktTotals'><strong>Job Assigned</strong></td>
                            </tr>
                            <% 
i = 1
sql ="SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_city, tech_email, tech_tel1, tech_tel2, " & _
     "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area " & _
	 "FROM tbltechnician where tech_status='Y' and (tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC') " 

if Request.Cookies("GAPS")("slevel") = "technician" then 
sql = sql & "and tech_code='" & Request.Cookies("GAPS")("job_tech_code") & "'"
end if
	 
sql = sql & " order by tech_code"

'response.write sql

set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
while not rs.eof 
if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
                            <tr valign="top" bgcolor="<%=nbgcolor%>">
                              <td nowrap="nowrap"><%=i%>.</td>
                              <td>
                              <%if Request.Cookies("GAPS")("slevel") = "cs" or Request.Cookies("GAPS")("slevel") = "sc" then %>
                              <a href="rm_contractor_new.asp?tech_code=<%=rs("tech_code")%>"><%=rs("tech_code")%></a>
                              <%else%>
                              <%=rs("tech_code")%>
                              <%end if%>
                              </td>
                              <td><%=rs("tech_name")%></td>
                              <td><%=rs("tech_area")%><br />
                                <%=rs("tech_tel1")%><br />
                                <%=rs("tech_tel2")%></td>
                              <td><%       j = 1 
	            sql1 = "SELECT tbljob.job_code, tbljob.job_date, tbljob.job_cust_name, tbljob.job_cust_state, tbljob.job_cust_city, tbljob.job_tech_code, tbltechnician.tech_code, tbltechnician.tech_name, tbltechnician.tech_tel1, tbltechnician.tech_area FROM " & _
                       "tbljob inner join tbltechnician on tbljob.job_tech_code=tbltechnician.tech_code where tbljob.job_appointment_date = '" & ChkDateYYYYMMDD(job_date) & "' and tbljob.job_tech_code='" & rs("tech_code") & "'"
            	'response.write sql1 & "<br>"
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  response.write j & ". <a href='rmtech_jobsheet.asp?job_code=" & rs1("job_code") & "'>" & rs1("job_code") & "</a> : " & rs1("job_cust_name") & ", " & rs1("job_cust_state") & ", " & rs1("job_cust_city") & "<br>"    
				j = j + 1
				rs1.movenext
				wend
				rs1.close	
	   %></td>
                            </tr>
                            <%
count = count + 1 
i = i + 1
rs.MoveNext
wend
rs.Close
Set rs = Nothing
%>
                          </table></td>
                        </tr>
                        <tr valign="top">
                          <td colspan="10" align="right">&nbsp;</td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td align="center">&nbsp;</td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
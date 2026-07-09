<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            background-color: #FFFFFF;
        }
    </style>
</head>
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

'response.Cookies("GAPS")("sqlexcel") = sql2

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>KPI Service Summary</div></td>
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
                   <!--   <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_tech_servicekpi_detail_excel.asp" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>-->
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_tech_servicekpi.asp">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Job Month  </strong>(dd-MMM-yyyy)<strong><br />
                        </strong></td>
                        <td><div align="left"><strong><font color="#000000"><strong>
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
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                          </span></div></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
 <%
sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Done', 'Posted') and DATEDIFF(day, job_submitteddate, job_donedate) < 4 and DATEDIFF(day, job_submitteddate, job_donedate) >= 0 "
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	    
totalcntJob1t3 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Done', 'Posted') and DATEDIFF(day, job_submitteddate, job_donedate) > 3 and DATEDIFF(day, job_submitteddate,job_donedate) < 8" 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	
totalcntJob4t7 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Done', 'Posted') and DATEDIFF(day, job_submitteddate, job_donedate) > 7 and DATEDIFF(day, job_submitteddate,job_donedate) < 15" 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	
totalcntJob8t14 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Done', 'Posted') and DATEDIFF(day, job_submitteddate, job_donedate) > 14 " 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	
totalcntJobmore14 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Done', 'Posted') and DATEDIFF(day, job_submitteddate, job_donedate) is null " 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	
totalcntJobmoreNull = selectid(sql1)

totalJobCompleted = ChkNumber(totalcntJob1t3) + ChkNumber(totalcntJob4t7) + ChkNumber(totalcntJob8t14) + ChkNumber(totalcntJobmore14) + ChkNumber(totalcntJobmoreNull)
    %> 
              <tr>
                  <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Jobsheet Completed </strong>(Submitted to Done) </font></td>
                    </tr>
                    <tr>
                      <td width="58%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Days</span></strong></font></td>
                      <td width="58%" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Quantities of Job</span></strong></font></td>
                    </tr>
                    <tr>
                      <td height="40" align="center"><strong>1-3 </strong>days</td>
                      <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&jobrange=1t3&jobrangestatus=completed&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"> <%=totalcntJob1t3%></a></strong></td>
                    </tr>
                    <tr>
                      <td height="40" align="center"><strong>4-7 </strong>days</td>
                      <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&jobrange=4t7&jobrangestatus=completed&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"> <%=totalcntJob4t7%></a></strong></td>
                    </tr>
                    <tr>
                      <td height="40" align="center"><strong>8-14 </strong>days</td>
                      <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&jobrange=8t14&jobrangestatus=completed&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"> <%=totalcntJob8t14%></a></strong></td>
                    </tr>
                   <tr>
                   <td height="40" align="center"><strong>&gt;14 </strong>days</td>
                   <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&jobrange=more14&jobrangestatus=completed&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=totalcntJobmore14%></a></strong></td>
                   </tr>
                   <tr bgcolor="#F3F3F3">
                     <td height="40" align="center" bgcolor="#FFFFFF">Incomplete Done Date</td>
                     <td height="40" align="center" bgcolor="#FFFFFF"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;jobrange=incompletedate&amp;jobrangestatus=completed&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=totalcntJobmoreNull%></a></strong></td>
                   </tr>
                   <tr bgcolor="#F3F3F3">
                      <td height="40" align="center" bgcolor="#FFFFFF"><strong>Total</strong></td>
                      <td height="40" align="center" bgcolor="#FFFFFF"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&jobrange=all&jobrangestatus=completed&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=totalJobCompleted%></a></strong></td>
                    </tr>
                </table></td>
                  <td width="51%" align="center" valign="top" bgcolor="#FFFFFF">

 <%
sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Submitted', 'Accepted') and DATEDIFF(day, job_submitteddate, job_JS_receiveddate) < 4" 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if
totalPendingJob1t3 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Submitted', 'Accepted') and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) > 3 and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) < 8" 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if
totalPendingJob4t7 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Submitted', 'Accepted') and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) > 7 and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) < 15" 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if
totalPendingJob8t14 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Submitted', 'Accepted') and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) > 14 " 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if
totalPendingJobmore14 = selectid(sql1)

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Submitted', 'Accepted') and DATEDIFF(day,job_submitteddate, job_JS_receiveddate) is null " 
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if
totalPendingJobNull = selectid(sql1)

totalJobPending = ChkNumber(totalPendingJob1t3) + ChkNumber(totalPendingJob4t7) + ChkNumber(totalPendingJob8t14) + ChkNumber(totalPendingJobmore14) + ChkNumber(totalPendingJobNull)
    %> 
                      
                  <table width="90%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="2" align="center" bgcolor="#000099" class="style1"><font color="#FFFFFF"><strong>Jobsheet Pending </strong>(Submitted to Received)</font></td>
                    </tr>
                    <tr>
                      <td width="58%" align="center" bgcolor="#000099" class="style1"><font color="#FFFFFF"><strong><span>Days</span></strong></font></td>
                      <td width="58%" align="center" nowrap="nowrap" bgcolor="#000099" class="style1"><font color="#FFFFFF"><strong><span>Quantities of Job</span></strong></font></td>
                    </tr>
                    <tr>
                      <td height="40" align="center"><strong>1-3 </strong>days</td>
                      <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&jobrange=1t3&jobrangestatus=pending&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"> <%=totalPendingJob1t3%></a></strong></td>
                    </tr>
                    <tr>
                      <td height="40" align="center"><strong>4-7 </strong>days</td>
                      <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&jobrange=4t7&jobrangestatus=pending&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"> <%=totalPendingJob4t7%></a></strong></td>
                    </tr>
                    <tr>
                      <td height="40" align="center"><strong>8-14 </strong>days</td>
                      <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&jobrange=8t14&jobrangestatus=pending&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"> <%=totalPendingJob8t14%></a></strong></td>
                    </tr>
                    <tr>
                      <td height="40" align="center"><strong>&gt;14 </strong>days</td>
                      <td align="center"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&jobrange=more14&jobrangestatus=pending&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=totalPendingJobmore14%></a></strong></td>
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" align="center" bgcolor="#FFFFFF">Incomplete Job Sheet Received<span class="auto-style1"> </span>Date</td>
                      <td height="40" align="center" bgcolor="#FFFFFF"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;jobrange=incompletedate&amp;jobrangestatus=pending&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=totalPendingJobNull%></a></strong></td>
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" align="center" bgcolor="#FFFFFF"><strong>Total</strong></td>
                      <td height="40" align="center" bgcolor="#FFFFFF"><strong><a href="javascript:popup('rm_rpt_tech_servicekpi_detail.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>&jobrange=all&jobrangestatus=pending&job_tech_code=<%=job_tech_code%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=totalJobPending%></a></strong></td>
                    </tr>
                  </table></td>
                </tr>
                 <tr>
                   <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                   <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                 </tr>
                 <tr>
                   <td valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
                    
                    
<%
if totalJobCompleted < 1 then 
   totalJobCompleted=1
end if
completedjobP1t3 = ChkNumber2(ChkNumber(totalcntJob1t3)/ChkNumber(totalJobCompleted)*100) 
completedjobP4t7 = ChkNumber2(ChkNumber(totalcntJob4t7)/ChkNumber(totalJobCompleted)*100) 
completedjobP8t14 = ChkNumber2(ChkNumber(totalcntJob8t14)/ChkNumber(totalJobCompleted)*100) 
completedjobPmore14 = ChkNumber2(ChkNumber(totalcntJobmore14)/ChkNumber(totalJobCompleted)*100)
completedjobPNull = ChkNumber2(ChkNumber(totalcntJobmoreNull)/ChkNumber(totalJobCompleted)*100)
completedjobP = ChkNumber(completedjobP1t3) + ChkNumber(completedjobP4t7) + ChkNumber(completedjobP8t14) + ChkNumber(completedjobPmore14) + ChkNumber(completedjobPNull)
%>
                     <tr>
                       <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Jobsheet Completed % </strong>(Submitted to Done) </font></td>
                     </tr>
                     <tr>
                       <td width="58%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Days</span></strong></font></td>
                       <td width="58%" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Quantities of Job</span></strong></font></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>1-3 </strong>days</td>
                       <td align="center"><strong> <%=completedjobP1t3%>%</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>4-7 </strong>days</td>
                       <td align="center"><strong> <%=completedjobP4t7%> %</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>8-14 </strong>days</td>
                       <td align="center"><strong> <%=completedjobP8t14%> %</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>&gt;14 </strong>days</td>
                       <td align="center"><strong><%=completedjobPmore14%>%</strong></td>
                     </tr>
                     <tr bgcolor="#F3F3F3">
                       <td height="40" align="center" bgcolor="#FFFFFF">Incomplete Done Date</td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=completedjobPNull%>%</strong></td>
                     </tr>
                     <tr bgcolor="#F3F3F3">
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong>Total</strong></td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=completedjobP%>%</strong></td>
                     </tr>
                   </table></td>
                   <td align="center" valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
                     <%

if totalJobPending < 1 then 
   totalJobPending=1
end if
PendingjobP1t3 = ChkNumber2(ChkNumber(totalPendingJob1t3)/ChkNumber(totalJobPending)*100) 
PendingjobP4t7 = ChkNumber2(ChkNumber(totalPendingJob4t7)/ChkNumber(totalJobPending)*100) 
PendingjobP8t14 = ChkNumber2(ChkNumber(totalPendingJob8t14)/ChkNumber(totalJobPending)*100) 
PendingjobPmore14 = ChkNumber2(ChkNumber(totalPendingJobmore14)/ChkNumber(totalJobPending)*100)
PendingjobPNull = ChkNumber2(ChkNumber(totalPendingJobNull)/ChkNumber(totalJobPending)*100)
PendingjobP = ChkNumber(PendingjobP1t3) + ChkNumber(PendingjobP4t7) + ChkNumber(PendingjobP8t14) + ChkNumber(PendingjobPmore14) + ChkNumber(PendingjobPNull)
%>
                     <tr>
                       <td colspan="2" align="center" bgcolor="#000099" class="style1"><font color="#FFFFFF"><strong>Jobsheet Pending % </strong>&nbsp;(Submitted to Received)</font></td>
                     </tr>
                     <tr>
                       <td width="58%" align="center" bgcolor="#000099" class="style1"><font color="#FFFFFF"><strong><span>Days</span></strong></font></td>
                       <td width="58%" align="center" nowrap="nowrap" bgcolor="#000099" class="style1"><font color="#FFFFFF"><strong><span>Quantities of Job</span></strong></font></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>1-3 </strong>days</td>
                       <td align="center"><strong> <%=PendingjobP1t3%>%</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>4-7 </strong>days</td>
                       <td align="center"><strong> <%=PendingjobP4t7%> %</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>8-14 </strong>days</td>
                       <td align="center"><strong> <%=PendingjobP8t14%> %</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>&gt;14 </strong>days</td>
                       <td align="center"><strong><%=PendingjobPmore14%>%</strong></td>
                     </tr>
                     <tr bgcolor="#F3F3F3">
                       <td height="40" align="center" bgcolor="#FFFFFF">Incomplete Accepted Date</td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=PendingjobPNull%>%</strong></td>
                     </tr>
                     <tr bgcolor="#F3F3F3">
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong>Total</strong></td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=PendingjobP%>%</strong></td>
                     </tr>
                   </table></td>
                 </tr>
                 <tr>
                   <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                   <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                 </tr>
                 <tr>
                   <td valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
<%

sql1 = "select count(rcn_id) as totaljob " & _
       "from tblrcn where  rcn_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and rcn_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and rcn_status in ('Posted') and right(rcn_job_code,2)='CF' " 
totalRCNCF = selectid(sql1)
    
sql1 = "select count(rcn_id) as totaljob " & _
       "from tblrcn where  rcn_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and rcn_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and rcn_status in ('Posted') and right(rcn_job_code,2)='WH' " 
totalRCNWH = selectid(sql1)

totalRCN= ChkNumber(totalRCNCF) + ChkNumber(totalRCNWH) 

if ChkNumber(totalRCNCF) <> 0 then 
totalRCNCFP = ChkNumber2(ChkNumber(totalRCNCF)/ChkNumber(totalRCN)*100)
else
totalRCNCFP = 0
end if

if ChkNumber(totalRCNWH) <> 0 then 
totalRCNWHP = ChkNumber2(ChkNumber(totalRCNWH)/ChkNumber(totalRCN)*100)
else
totalRCNWHP = 0
end if

if ChkNumber(totalRCNCFP) <> 0 then 
totalRCNP = ChkNumber(totalRCNCFP) + ChkNumber(totalRCNWHP)
else
totalRCNP = 0
end if
%>
                     <tr>
                       <td colspan="3" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>RCN</strong></font></td>
                     </tr>
                     <tr>
                       <td width="58%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Type</span></strong></font></td>
                       <td width="29%" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Quantities of Job</span></strong></font></td>
                       <td width="29%" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>%</strong></font></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>WH</strong></td>
                       <td align="center"><strong> <%=totalRCNWH%></strong></td>
                       <td align="center"><strong><%=totalRCNWHP%>%</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>CF</strong></td>
                       <td align="center"><strong> <%=totalRCNCF%></strong></td>
                       <td align="center"><strong><%=totalRCNCFP%>%</strong></td>
                     </tr>
                     <tr bgcolor="#F3F3F3">
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong>Total</strong></td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=totalRCN%></strong></td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=totalRCNP%>%</strong></td>
                     </tr>
                   </table></td>
                   <td align="center" valign="top" bgcolor="#FFFFFF"><table width="90%" border="1" cellpadding="4" cellspacing="0">
                     <%
sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Cancel') and right(job_code,2)='CF' "
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	   
totalCancelCF = selectid(sql1) 

sql1 = "select count(job_id) as totaljob " & _
       "from tbljob where  job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " & _
	   "and job_status in ('Cancel') and right(job_code,2)='WH' "
if job_tech_code <> "All" then 
sql1 = sql1 & " and job_tech_code = '" & job_tech_code & "' "
end if	   
totalCancelWH = selectid(sql1) 

totalCancel= ChkNumber(totalCancelCF) + ChkNumber(totalCancelWH) 

if ChkNumber(totalCancelCF) <> 0 then 
totalCancelCFPer = ChkNumber2(ChkNumber(totalCancelCF)/ChkNumber(totalCancel)*100)
else
totalCancelCFPer = 0
end if

if ChkNumber(totalCancelWH) <> 0 then 
totalCancelWHPer = ChkNumber2(ChkNumber(totalCancelWH)/ChkNumber(totalCancel)*100)
else
totalCancelWHPer = 0
end if

if ChkNumber(totalCancel) <> 0 then 
totalCancelPer = ChkNumber(totalCancelCFPer) + ChkNumber(totalCancelWHPer)
else
totalCancelPer = 0
end if
%>
                     <tr>
                       <td colspan="3" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cancel Job</strong></font></td>
                     </tr>
                     <tr>
                       <td width="58%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Type</span></strong></font></td>
                       <td width="29%" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Quantities of Job</span></strong></font></td>
                       <td width="29%" align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>%</strong></font></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>WH</strong></td>
                       <td align="center"><strong> <%=totalCancelWH%></strong></td>
                       <td align="center"><strong><%=totalCancelWHPer%>%</strong></td>
                     </tr>
                     <tr>
                       <td height="40" align="center"><strong>CF</strong></td>
                       <td align="center"><strong> <%=totalCancelCF%></strong></td>
                       <td align="center"><strong><%=totalCancelCFPer%>%</strong></td>
                     </tr>
                     <tr bgcolor="#F3F3F3">
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong>Total</strong></td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=totalCancel%></strong></td>
                       <td height="40" align="center" bgcolor="#FFFFFF"><strong><%=totalCancelPer%>%</strong></td>
                     </tr>
                   </table></td>
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
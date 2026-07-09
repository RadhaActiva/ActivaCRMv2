<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            width: 10%;
        }
    </style>
</head>
<%
jobmonth = request("jobmonth")
jobyear = request("jobyear")
job_tech_code = request("job_tech_code")

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

if request("claim_status") <> "" then
   claim_status = request("claim_status")
else
   claim_status = "All"
end if

if request("job_status") <> "" then
   job_status = request("job_status")
else
   job_status = "All"
end if

totalAmt=0
total_job=0
total_jobcount1=0
total_jobcount2=0
total_overwrtty=0
total_installation=0
total_jobcount2_TPC_IC_IHC_IHT=0

sql2 = "Select tbljob.job_tech_code, tbltechnician.tech_name,tbltechnician.tech_type, tbljob.job_count, tbljob.job_code, tbljob.job_tech_faulty_action,tbljob.job_donedate,job_tech_model,tbljob.job_faulty_reason_cs, tbljob.job_model_desc, tbljob.job_tech_sn, tbljob.job_actual_wrty_status, tbljob.job_totalAmt, tbljob.job_status,job_claim_approved from tbljob " & _
    "join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
    "and month(job_submitforclaims_date) = " & jobmonth & " and year(job_submitforclaims_date) = " & jobyear & " and job_submitforclaims='Yes' "
    
if job_tech_code <> "All" then 
sql2 = sql2 & " and tbljob.job_tech_code = '" & job_tech_code & "' "
end if	    

if claim_status="Yes" then
    sql2 = sql2 & " and tbljob.job_claim_approved = 'Yes' "
elseif claim_status="No" then
    sql2 = sql2 & " and tbljob.job_claim_approved IS NULL "
elseif claim_status="All" then
    sql2=sql2
end if

if job_status="Done" then
    sql2 = sql2 & " and tbljob.job_status = 'Done' "
elseif job_status="Posted" then
    sql2 = sql2 & " and tbljob.job_status = 'Posted' "
elseif job_status="All" then
    sql2=sql2
end if

sql2 = sql2 & "order by tbljob.job_status,tbljob.job_donedate"

'response.write sql2
    
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
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&job_tech_code=" & job_tech_code & "&claim_status=" & claim_status & "&job_status=" & job_status
%> 


<script>
function DisplayReport() 
{
	document.form1.action = "rm_rpt_weekly_job_submission.asp";
	document.form1.submit();
}
</script> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                         <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Service Weekly Job Submission</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%"> <strong></strong></td>
                      <td width="20%" align="center" class="titlegrey1">
                          <% if job_tech_code <> "All" then %>
                            <a href="rm_rpt_weekly_job_submission_excel.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&job_tech_code=<%=job_tech_code%>&claim_status=<%=claim_status%>" target="_blank">
                            <img src="images/excel.jpg" width="57" height="21" border="0" /></a>
                          <%end if%>
                     </td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="rm_rpt_weekly_job_submission.asp">
                <tr>
                  <td height="30" align="left" bgcolor="#FFFFFF"><font color="#000000">
                    Claim Submission Month 
                     <select name="jobyear" id="jobyear">
                        <option value="2022"<%if jobyear="2022" then response.write " selected"%>>2022</option>
                        <option value="2023"<%if jobyear="2023" then response.write " selected"%>>2023</option>
                        <option value="2024"<%if jobyear="2024" then response.write " selected"%>>2024</option>
                        <option value="2025"<%if jobyear="2025" then response.write " selected"%>>2025</option>
                        <option value="2026"<%if jobyear="2026" then response.write " selected"%>>2026</option>
                    </select>
                    <select name="jobmonth" id="jobmonth">
                      <option value="1" <%if jobmonth="1" then response.write " selected"%>>Jan</option>
                      <option value="2" <%if jobmonth="2" then response.write " selected"%>>Feb</option>
                      <option value="3" <%if jobmonth="3" then response.write " selected"%>>Mar</option>
                      <option value="4" <%if jobmonth="4" then response.write " selected"%>>Apr</option>
                      <option value="5" <%if jobmonth="5" then response.write " selected"%>>May</option>
                      <option value="6" <%if jobmonth="6" then response.write " selected"%>>Jun</option>
                      <option value="7" <%if jobmonth="7" then response.write " selected"%>>Jul</option>
                      <option value="8" <%if jobmonth="8" then response.write " selected"%>>Aug</option>
                      <option value="9" <%if jobmonth="9" then response.write " selected"%>>Sep</option>
                      <option value="10" <%if jobmonth="10" then response.write " selected"%>>Oct</option>
                      <option value="11" <%if jobmonth="11" then response.write " selected"%>>Nov</option>
                      <option value="12" <%if jobmonth="12" then response.write " selected"%>>Dec</option>
                    </select></font>
                  <td height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                      <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>                    
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF">Technician
                 <select name="job_tech_code" id="job_tech_code">
                      <option value="" <%if job_tech_code="" then response.write " selected"%>>All Technicians</option>
                        <%			
				    sql2="SELECT tech_code, tech_name FROM tbltechnician where tech_code in " &_
                    "(Select distinct tbljob.job_tech_code from tbljob join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
                    "and month(job_submitforclaims_date) = '" & jobmonth & "' and year(job_submitforclaims_date) = '" & jobyear & "' and job_submitforclaims='Yes')  order by tech_code"
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
                      &nbsp;&nbsp;&nbsp;
                      Claim Processed
                       <select name="claim_status" id="claim_status">
                            <option value="" <%if claim_status="" then response.write " selected"%>>All</option>
                            <option value="Yes" <%if claim_status="Yes" then response.write " selected"%>>Yes</option>
                            <option value="No" <%if claim_status="No" then response.write " selected"%>>No</option>
                    </select>
                     &nbsp;&nbsp;&nbsp;
                      Job Status
                       <select name="job_status" id="job_status">
                            <option value="" <%if job_status="" then response.write " selected"%>>All</option>
                            <option value="Done" <%if job_status="Done" then response.write " selected"%>>Done</option>
                            <option value="Posted" <%if job_status="Posted" then response.write " selected"%>>Posted</option>
                    </select>
                    <span class="titlegrey1">
                     &nbsp;
                    <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" />
                  </span></td>
                  <td width="20%" height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                 </form>
                
              <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="6" align="right" class="style1"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_weekly_job_submission.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_weekly_job_submission.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="4%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="8%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Tech</span></strong></font></td>
                      <td width="12%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Jobsheet No</span></strong></font></td>
                      <td width="8%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Service Date</span></strong></font></td>
                      <td width="30%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Service Action Taken</span></strong></font></td>
                      <td width="55%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Model Desc</strong></font></td>
                      <td width="20%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Serial No</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Warranty</span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="auto-style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Job Status</strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Claim<br/>Processed</strong></font></td>
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
                      <td align="left" bgcolor="#FFFFFF"><%=rs("job_tech_code")%></td>
                      <td align="left" nowrap="nowrap"><font color="#0000FF"><a href="rm_jobsheet.asp?job_code=<%=rs("job_code")%>" target="_blank"><%=rs("job_code")%></a></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=chkdate(rs("job_donedate"))%></td>
                      <td align="center"><%=rs("job_tech_faulty_action")%></td>
                      <td align="left" bgcolor="#FFFFFF"><%=rs("job_model_desc")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs("job_tech_sn")%></td>
                      <td align="center"><%=rs("job_actual_wrty_status")%></td>
                      <td align="center" class="auto-style1"><%=ChkNumber2(rs("job_totalAmt"))%></td>
                       <td align="center"><strong><%=rs("job_status")%></strong></td>
                       <td align="center"><strong><%=rs("job_claim_approved")%></strong></td>
                    </tr>
<%
'dont include 4 models of water storage in the calculations
if rs("job_tech_model") <> "TAE07-800" and rs("job_tech_model") <> "TAE07-810" and rs("job_tech_model") <> "TAE07-811" and rs("job_tech_model") <> "TAE07-812" then
        if isnumeric(rs("job_totalAmt")) then 
            totalAmt = totalAmt + rs("job_totalAmt")
        end if

        if rs("job_count") = 1 and rs("job_actual_wrty_status") <> "Over" and rs("job_faulty_reason_cs") <>  "Installation" then 
            total_jobcount1 = total_jobcount1 + 1
        end if
  
        if rs("job_actual_wrty_status") = "Over" then 
            total_overwrtty = total_overwrtty + 1
        end if

        if rs("job_faulty_reason_cs") = "Installation" then 
            total_installation = total_installation + 1
        end if

         if rs("tech_type") = "TPC" or rs("tech_type") = "IC" or rs("tech_type") = "IHC" or rs("tech_type") = "IHT" then
            if rs("job_count") > 1 and rs("job_actual_wrty_status") = "Under" and rs("job_faulty_reason_cs") <> "Installation" then 'logic applicable for TPC and IC and IHC '181225 IHT added
                total_jobcount2_TPC_IC_IHC_IHT = total_jobcount2_TPC_IC_IHC_IHT + 1
            end if
        end if

        total_job = total_job + 1
end if

tech_type = rs("tech_type")
    
count = count + 1 
i = i + 1
rs.MoveNext
Next

rs.Close
Set rs = Nothing

'different technicians type has diff logic for footer figures

'if tech_type = "IHT" then
   ' total_jobcount1 = total_job - total_overwrtty
   ' total_jobcount2 = 0 'NO 2ND JOB AS THIS IS INCLUDED IN 1ST JOB ALREADY
'end if 

'if tech_type = "IHC" then
'    total_overwrtty = 0  'zero for IHC
    'total_jobcount1 = count - (total_overwrtty + total_jobcount)    
'end if 

if tech_type = "TPC" or tech_type = "IC" or tech_type = "IHC" or tech_type = "IHT" then
    total_jobcount2 = total_jobcount2_TPC_IC_IHC_IHT
    'total_jobcount1 = total_job - (total_overwrtty + total_jobcount2  + total_installation)
    'total_jobcount1 = total_jobcount1 - (total_overwrtty + total_installation)
    'total_jobcount2 = total_jobcount2_TPC_IC
end if 


%>                 
                    <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="9" align="right" bgcolor="#CCCCCC"><strong>Total :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=chknumber2(totalAmt)%>&nbsp;</strong></td>
                     <td align="center" colspan="2" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="9" align="right" bgcolor="#CCCCCC"><strong>Total Job :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_job%>&nbsp;</strong></td>
                     <td align="center" colspan="2" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="9" align="right" bgcolor="#CCCCCC"><strong>1st Unit :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_jobcount1%>&nbsp;</strong></td>
                     <td align="center" colspan="2" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="9" align="right" bgcolor="#CCCCCC"><strong>> 2nd Unit :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_jobcount2%>&nbsp;</strong></td>
                     <td align="center" colspan="2" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                    <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="9" align="right" bgcolor="#CCCCCC"><strong>O/Wrtty :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_overwrtty%>&nbsp;</strong></td>
                     <td align="center" colspan="2" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                   <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="9" align="right" bgcolor="#CCCCCC"><strong>Installlation :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_installation%>&nbsp;</strong></td>
                     <td align="center" colspan="2" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="6" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rpt_weekly_job_submission.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rpt_weekly_job_submission.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                </table></td>
                </tr>
                <tr>
                  <td height="30" colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
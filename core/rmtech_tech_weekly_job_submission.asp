<!-- #include file="header.asp" -->
<%
job_tech_code = request.cookies("GAPS")("sloginid")

jobmonth = request("jobmonth")
jobyear = request("jobyear")

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


totalAmt=0
total_job=0
total_jobcount1=0
total_jobcount2=0
total_overwrtty=0
total_jobcount2_TPC_IC_IHC=0

sql2 = "Select tbljob.job_tech_code, tbltechnician.tech_name,tbltechnician.tech_type, tbljob.job_count, tbljob.job_code, tbljob.job_tech_faulty_action,tbljob.job_date,job_tech_model, tbljob.job_model_desc, tbljob.job_tech_sn, tbljob.job_actual_wrty_status, tbljob.job_totalAmt, tbljob.job_status from tbljob " & _
    "join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
    "and month(job_submitforclaims_date) = " & jobmonth & " and year(job_submitforclaims_date) = " & jobyear & " and job_submitforclaims='Yes' and tbljob.job_status = 'Posted' and tbljob.job_tech_code = '" & job_tech_code & "' "

'sql2 = sql2 & "Order by tbljob.job_date desc" 'to sort by date
sql2 = sql2 & "Order by tbljob.job_code"

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
link = "&jobyear=" & jobyear & "&jobmonth=" & jobmonth & "&job_tech_code=" & job_tech_code
%> 


<script>
function DisplayReport() 
{
	document.form1.action = "tech_rpt_weekly_job_submission.asp";
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
                         <td class="titleblue1"><div align="left">Report Service Weekly Job Submission (Posted Jobs)</div></td>
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
                            <a href="tech_rpt_weekly_job_submission_excel.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&job_tech_code=<%=job_tech_code%>" target="_blank">
                            <img src="images/excel.jpg" width="57" height="21" border="0" /></a>
                          <%end if%>
                     </td>
                    </tr>
                  </table></td>
                </tr>
                <form id="form1" name="form1" method="post" action="tech_rpt_weekly_job_submission.asp">
                  <tr>                    
                  <td width="80%" height="30" align="left" bgcolor="#FFFFFF">Technician : <%=job_tech_code%></td>
                  <td width="20%" height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                    <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
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
                    <span class="titlegrey1">
                    <input type="button" name="button2" id="button" value="Display Report" onclick="javascript:DisplayReport();" />
                  </span>
                      </td>
                  <td height="30" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                      <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
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
					Response.Write " <a href='tech_rpt_weekly_job_submission.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='tech_rpt_weekly_job_submission.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                    </tr>
                    <tr>
                      <td width="4%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="8%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Tech</span></strong></font></td>
                      <td width="12%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Jobsheet No</span></strong></font></td>
                      <td width="8%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Service Date</span></strong></font></td>
                      <td width="30%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Service Action Taken</span></strong></font></td>
                      <td width="55%" align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Model Desc</strong></font></td>
                      <td width="20%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Serial No</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Warranty</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                      <td width="10%" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Status</strong></font></td>
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
                      <td align="left" nowrap="nowrap"><%=rs("job_code")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=chkdate(rs("job_date"))%></td>
                      <td align="center"><%=rs("job_tech_faulty_action")%></td>
                      <td align="left" bgcolor="#FFFFFF"><%=rs("job_model_desc")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs("job_tech_sn")%></td>
                      <td align="center"><%=rs("job_actual_wrty_status")%></td>
                      <td align="center"><%=ChkNumber2(rs("job_totalAmt"))%></td>
                       <td align="center"><strong><%=rs("job_status")%></strong></td>
                    </tr>
<%
'dont include 4 models of water storage in the calculations
if rs("job_tech_model") <> "TAE07-800" and rs("job_tech_model") <> "TAE07-810" and rs("job_tech_model") <> "TAE07-811" and rs("job_tech_model") <> "TAE07-812" then
        if isnumeric(rs("job_totalAmt")) then 
            totalAmt = totalAmt + rs("job_totalAmt")
        end if

        if rs("job_count") = 1 then 
            total_jobcount1 = total_jobcount1 + 1
        end if

        if rs("job_count") > 1 then 
            total_jobcount2 = total_jobcount2 + 1
        end if

        if rs("job_actual_wrty_status") = "Over" then 
            total_overwrtty = total_overwrtty + 1
        end if

       if rs("tech_type") = "TPC" or rs("tech_type") = "IC" or rs("tech_type") = "IHC" then
            if rs("job_count") > 1 and rs("job_actual_wrty_status") = "Under" then 'logic applicable for TPC and IC and IHC
                total_jobcount2_TPC_IC_IHC = total_jobcount2_TPC_IC_IHC + 1
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

if tech_type = "IHT" then
    total_jobcount1 = total_job - total_overwrtty
    total_jobcount2 = 0
end if 

'if tech_type = "IHC" then
'    total_overwrtty = 0  'zero for IHC
    'total_jobcount1 = count - (total_overwrtty + total_jobcount)    
'end if 

if tech_type = "TPC" or tech_type = "IC" or tech_type = "IHC" then
    total_jobcount2 = total_jobcount2_TPC_IC_IHC
    total_jobcount1 = total_job - (total_overwrtty + total_jobcount2)
    'total_jobcount2 = total_jobcount2_TPC_IC
end if

%>
                   
                    <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>Total :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=chknumber2(totalAmt)%>&nbsp;</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>Total Job :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_job%>&nbsp;</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>1st Unit :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_jobcount1%>&nbsp;</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>> 2nd Unit :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_jobcount2%>&nbsp;</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                    <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>O/Wrtty :</strong></td>
                     <td align="left" bgcolor="#CCCCCC"><strong><%=total_overwrtty%>&nbsp;</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>  
                   </tr>
                     <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="6" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                       <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='tech_rpt_weekly_job_submission.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='tech_rpt_weekly_job_submission.asp?num=" & Showed+row & link & "'> Next >></a>"
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
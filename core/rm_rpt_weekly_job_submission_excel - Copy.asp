<%  
    
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=weekly_service_job_submission_" & year(date()) & month(date()) & day(date()) & ".xls"
job_tech_code = request("job_tech_code")
'job_tech_name = request("job_tech_name")
jobyear = request("jobyear")
jobmonth = request("jobmonth")

%>
<!-- #include file="database/datastore.asp" -->
<%
sql2 = request.Cookies("GAPS")("sqlexcel")

i=1
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open

total_job = 0
total_jobcount1=0
total_jobcount2=0
total_overwrtty=0
total_installation=0
total_jobcount2_TPC_IC_IHC_IHT=0
%>

<h2>Weekly Service Job Submission Form</h2>
<h4>Date : <%=day(date())%>/<%=month(date())%>/<%=Year(date())%> </h4>
<h4>Contractor Name : <%=rs1("tech_name")%></h4>
<h4>Contractor Code : <%=job_tech_code%></h4>

<table width="100%" border="0" cellpadding="4" cellspacing="0">
  <tr>
    <td width="50" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
    <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Job Sheet No</span></strong></font></td>
    <td width="90" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Service Date</strong></font></td>
    <td width="343" height="30" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Service Action Taken</span></strong></font></td>
    <td width="109" height="30" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Model Desc</span></strong></font></td>
    <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Serial No<span></span></strong><br/></td>
    <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Warranty Info</strong><br/></td>
    <td width="104" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Job/Status</strong><br/></td>
    <td width="139" height="30" align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total&nbsp;</strong></font></td>
  </tr>
  <%
  
t_totalqty = 0
t_totalvalue = 0


while not rs1.eof
if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if


%>
  <tr bgcolor="<%=nbgcolor%>">
    <td align="left" nowrap="nowrap"><%=i%></td>
    <td align="left" nowrap="nowrap"><strong><%=rs1("job_code")%></strong></td>
    <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><%=chkdate(rs1("job_donedate"))%></td>
    <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs1("job_tech_faulty_action")%></td>
    <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs1("job_model_desc")%></td>   
    <td align="left" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs1("job_tech_sn")%></td>
    <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs1("job_actual_wrty_status")%></td>
    <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><%=rs1("job_status")%></td>
    <td align="right" nowrap="nowrap" bgcolor="#F3F3F3"><%=ChkNumber2(rs1("job_totalAmt"))%></td>
  </tr>
  <%

if rs1("job_tech_model") <> "TAE07-800" and rs1("job_tech_model") <> "TAE07-810" and rs1("job_tech_model") <> "TAE07-811" and rs1("job_tech_model") <> "TAE07-812" then

        t_amount = t_amount + cint(rs1("job_totalAmt")) 

        if rs1("job_count") = 1 then 
            total_jobcount1 = total_jobcount1 + 1
        end if

        'if rs1("job_count") > 1 then 
         '   total_jobcount2 = total_jobcount2 + 1
        'end if

        if rs1("job_actual_wrty_status") = "Over" then 
            total_overwrtty = total_overwrtty + 1
        end if

        if rs1("job_faulty_reason_cs") = "Installation" then 
            total_installation = total_installation + 1
        end if

         if rs1("tech_type") = "TPC" or rs1("tech_type") = "IC" or rs1("tech_type") = "IHC" or rs1("tech_type") = "IHT" then
            if rs1("job_count") > 1 and rs1("job_actual_wrty_status") = "Under" then 'logic applicable for TPC and IC and IHC '181225 IHT added
                total_jobcount2_TPC_IC_IHC_IHT = total_jobcount2_TPC_IC_IHC_IHT + 1
            end if
        end if
        total_job = total_job + 1 'need to count jobs that are not water storage only
end if

tech_type = rs1("tech_type")

count = count + 1 
i = i + 1
rs1.MoveNext
wend
rs1.Close
Set rs1 = Nothing

If Err.Number <> 0 Then
  Response.Write (Err.Description)   
  Response.End 
End If

'different technicians type has diff logic for footer figures

if tech_type = "IHT" then
    total_jobcount1 = total_job - total_overwrtty
    total_jobcount2 = 0 'NO 2ND JOB AS THIS IS INCLUDED IN 1ST JOB ALREADY
end if 

'if tech_type = "IHC" then
'    total_overwrtty = 0  'zero for IHC
    'total_jobcount1 = count - (total_overwrtty + total_jobcount)    
'end if 

if tech_type = "TPC" or tech_type = "IC" or tech_type = "IHC" or tech_type = "IHT" then
    total_jobcount2 = total_jobcount2_TPC_IC_IHC_IHT
    total_jobcount1 = total_job - (total_overwrtty + total_jobcount2)
    'total_jobcount2 = total_jobcount2_TPC_IC
end if 

%>
  <tr bgcolor="#F3F3F3">
    <td height="30" colspan="8" align="right" bgcolor="#999999"><strong>Grand</strong> <strong>Total</strong></td>
    <td align="rightr" nowrap="nowrap" bgcolor="#999999"><strong><%=t_amount%></strong></td>
  </tr>
                 <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>Total Job :</strong></td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=total_job%>&nbsp;</strong></td>
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>1st Unit :</strong></td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=total_jobcount1%>&nbsp;</strong></td>
                   </tr>
                     <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>> 2nd Unit :</strong></td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=total_jobcount2%>&nbsp;</strong></td>
                   </tr>
                    <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>O/Wrtty :</strong></td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=total_overwrtty%>&nbsp;</strong></td>
                   </tr>
                    <tr bgcolor="#F3F3F3">
                     <td height="20" colspan="8" align="right" bgcolor="#CCCCCC"><strong>Installation :</strong></td>
                     <td align="right" bgcolor="#CCCCCC"><strong><%=total_installation%>&nbsp;</strong></td>
                   </tr>
</table>

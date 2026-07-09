<%  
tech_code = request("tech_code")    
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=TPC_Extra_Mileage_claim_"& tech_code & year(date()) & month(date()) & day(date()) & ".xls"
'job_tech_name = request("job_tech_name")
jobyear = request("jobyear")
jobmonth = request("jobmonth")

%>
<!-- #include file="database/datastore.asp" -->
<%
sql2 = request.Cookies("GAPS")("sqlexcel")

i=1

set rs1 = server.CreateObject("adodb.recordset")
sql1 = "SELECT te_id,te_date,te_year,te_month,te_tech_code,te_trip_date, te_job_sheet,te_loc_start,te_loc_end,te_mileage,te_offset_mileage,te_net_mileage,te_claim_amount " & _      
        "FROM tbltech_claim_exmileage where te_year ='" & jobyear & "' and te_month ='" & jobmonth & "' and te_tech_code = '" & tech_code & "'  order by te_date desc"
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open

total_net_claim=0

sql = "select tech_name from tbltechnician where tech_code='" & tech_code & "'"
tech_name = selectid(sql)
%>


<h2>Monthly Extra Mileage Claim (TPC)</h2>
<h4>Date : <%=day(date())%>/<%=month(date())%>/<%=Year(date())%> </h4>
Contractor Name : <%=tech_name%> <br/>
Contractor Code : <%=tech_code%> <br/>

<table width="100%" border="0" cellpadding="4" cellspacing="0">
                   <tr>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>No</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Submission Date</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>MM / YY</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Trip Date</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Sheet No</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Location Start</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Location End</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Mileage (KM)</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Offset (KM)</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Net Mileage (KM)</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Claim Amt (RM)</strong></font></td>
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
                    <tr bgcolor="#FFFFFF">
                      <td align="center"><%=i%></td>
                      <td align="center"> <%=chkdate(rs1("te_date"))%></td>
                      <td align="center"> <%=rs1("te_month")%>/<%=rs1("te_year")%></td>
                      <td align="center"> <%=chkdate(rs1("te_trip_date"))%></td>
                      <td align="center"> <%=rs1("te_job_sheet")%></td>
                      <td align="center"> <%=rs1("te_loc_start")%> </td>
                      <td align="center"> <%=rs1("te_loc_end")%> </td>
                      <td align="center"> <%=rs1("te_mileage")%> </td>
                      <td align="center"> <%=rs1("te_offset_mileage")%> </td>
                      <td align="center"> <%=rs1("te_net_mileage")%> </td>
                      <td align="center"> <%=chknumber2(rs1("te_claim_amount"))%> </td>                      
                    </tr>
  <%
total_net_claim=total_net_claim + rs1("te_claim_amount")

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


%>
  <tr bgcolor="#F3F3F3">
    <td height="30" colspan="10" align="right" bgcolor="#999999"><strong>Total</strong></td>
    <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(total_net_claim)%></strong></td>
  </tr>
</table>

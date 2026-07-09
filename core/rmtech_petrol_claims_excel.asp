<%  
tech_code = request("tech_code")    
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=Monthly_Petrol_Claims_"& tech_code & "_" & year(date()) & month(date()) & day(date()) & ".xls"
'job_tech_name = request("job_tech_name")
jobyear = request("jobyear")
jobmonth = request("jobmonth")

%>
<!-- #include file="database/datastore.asp" -->
<%
sql2 = request.Cookies("GAPS")("sqlexcel")

i=1

set rs1 = server.CreateObject("adodb.recordset")
sql1 = "SELECT tp_id,tp_date,tp_year,tp_month,tp_vehicle_no,tp_tech_code,tp_trip_date,tp_job_sheet,tp_mileage_start,tp_mileage_end,tp_distance,tp_claim_amount " & _      
    "FROM tbltech_claim_petrol where tp_tech_code = '" & request.cookies("GAPS")("sloginid") & "' and tp_year ='" & jobyear & "' and tp_month ='" & jobmonth & "' order by tp_date desc"

rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open

total_petrol=0

sql = "select tech_name from tbltechnician where tech_code='" & tech_code & "'"
tech_name = selectid(sql)
%>


<h2>Monthly Petrol Claims</h2>
<h4>Date : <%=day(date())%>/<%=month(date())%>/<%=Year(date())%> </h4>
Contractor Name : <%=tech_name%> <br/>
Contractor Code : <%=tech_code%> <br/>
<br/>

<table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" width="10%" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>No</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Submission Date</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>MM / YY</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Trip Date</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Sheet</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Mileage Begin</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Mileage End</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Distance (KM)</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Claim Amount (RM)</strong></font></td>
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
claimed_amount = 0

if not isnull(rs1("tp_claim_amount")) then
    claimed_amount = rs1("tp_claim_amount")* 0.9
end if

%>
               <tr bgcolor="#FFFFFF">
                      <td align="center"><%=i%></td>
                      <td align="center"> <%=chkdate(rs1("tp_date"))%></td>
                      <td align="center"> <%=rs1("tp_month")%>/<%=rs1("tp_year")%></td>
                      <td align="center"> <%=chkdate(rs1("tp_trip_date"))%></td>
                      <td align="center"> <%=rs1("tp_job_sheet")%> </td>
                      <td align="center"> <%=rs1("tp_mileage_start")%> </td>
                      <td align="center"> <%=rs1("tp_mileage_end")%> </td>
                      <td align="center"> <%=rs1("tp_distance")%></td>
                      <td align="center"> <%=ChkNumber2(claimed_amount)%></td>  
                    </tr>
  <%

total_petrol = total_petrol + claimed_amount

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
    <td height="30" colspan="8" align="right" bgcolor="#999999"><strong>Total</strong></td>
    <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(total_petrol)%></strong></td>
  </tr>
</table>

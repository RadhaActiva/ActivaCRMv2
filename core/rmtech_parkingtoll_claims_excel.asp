<%  
tech_code = request("tech_code")    
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=Monthly_Parking_Toll_Claims_"& tech_code & year(date()) & month(date()) & day(date()) & ".xls"
'job_tech_name = request("job_tech_name")
jobyear = request("jobyear")
jobmonth = request("jobmonth")

%>
<!-- #include file="database/datastore.asp" -->
<%
sql2 = request.Cookies("GAPS")("sqlexcel")

i=1

set rs1 = server.CreateObject("adodb.recordset")

sql1 = "SELECT tpt_id,tpt_date,tpt_year,tpt_month,tpt_tech_code,tpt_job_sheet,tpt_parking_amount, tpt_toll_amount,tpt_trip_date " & _      
       "FROM tbltech_claim_parkingtoll where tpt_tech_code = '" & tech_code & "' and tpt_year ='" & jobyear & "' and tpt_month ='" & jobmonth & "' order by tpt_date"
    

rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open

total_toll=0
total_parking=0

sql = "select tech_name from tbltechnician where tech_code='" & tech_code & "'"
tech_name = selectid(sql)
%>


<h2>Monthly Parking and Toll Claims</h2>
<h4>Date : <%=day(date())%>/<%=month(date())%>/<%=Year(date())%> </h4>
Contractor Name : <%=tech_name%> <br/>
Contractor Code : <%=tech_code%> <br/>

<table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" width="10%" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Trans ID</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Submission Date</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>MM / YY</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Trip Date</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Job Sheet</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Parking Amount</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Toll Amount</strong></font></td>
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
                      <td align="center"><%=chkdate(rs1("tpt_date"))%></td>
                      <td align="center"><%=rs1("tpt_month")%>/<%=rs1("tpt_year")%></td>
                      <td align="center"><%=chkdate(rs1("tpt_trip_date"))%></td>
                      <td align="center"><%=rs1("tpt_job_sheet")%> </td>
                      <td align="center"><%=ChkNumber2(rs1("tpt_parking_amount"))%> </td>
                      <td align="center"><%=ChkNumber2(rs1("tpt_toll_amount"))%> </td>                    
                    </tr>
  <%

total_toll = total_toll + rs1("tpt_toll_amount")
total_parking = total_parking + rs1("tpt_parking_amount")

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
    <td height="30" colspan="5" align="right" bgcolor="#999999"><strong>Total</strong></td>
    <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(total_parking)%></strong></td>
    <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=chknumber2(total_toll)%></strong></td>
  </tr>
</table>

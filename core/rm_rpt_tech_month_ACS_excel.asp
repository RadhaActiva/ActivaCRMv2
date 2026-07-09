<%  
Response.ContentType = "application/vnd.ms-excel"
tech_code = request("tech_code")
tech_name = request("tech_name")
Response.AddHeader "content-disposition", " filename=Arm_Statement_" & tech_code & "_" & year(date()) & month(date()) & day(date()) & ".xls"
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
%>

<h2>ACS Claim Statement</h2>
<h4>Date : <%=day(date())%>/<%=month(date())%>/<%=Year(date())%> </h4>
<h4>Contractor Name : <%=tech_name%> </h4>
<h4>Contractor Code : <%=tech_code%></h4>

<table width="100%" border="0" cellpadding="4" cellspacing="0">
  <tr>
                  <td width="6%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>JS Date</span></strong></font></td>
                  <td width="5%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>SJS No</span></strong></font></td>
                  <td width="5%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Complete Date</span></strong></font></td>
                  <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Purchase Date<br /></span></strong></font></td>
                  <td width="15%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Dealer Name<br /></span></strong></font></td>
                  <td width="40%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Model / Item</span></strong></font></td>
                  <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Service</strong></font></td>
                  <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Trip</span></strong></font></td>
                  <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Parking/Toll</span></strong></font></td>
                  <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Hotel</span></strong></font></td>
                  <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total (RM)</span></strong></font></td>
                  <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Inv</span></strong></font></td>
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

totalperjob=0
parkingtoll=0
trip_amt=0
labor=0
hotel_amount=0
exmileage_amount=0

if isnumeric(rs1("job_totallabourAmt")) then 
    labor = rs1("job_totallabourAmt")
end if

if isnumeric(rs1("trip_amount")) then 
    trip_amt = rs1("trip_amount") 'petrol claims for IHT/IHC/IC
end if

if isnumeric(rs1("exmileage_amount")) then 
    exmileage_amount = rs1("exmileage_amount") 'petrol claims for TPC only
end if

if isnumeric(rs1("parkingtoll_amount")) then 
    parkingtoll = rs1("parkingtoll_amount")
end if

if isnumeric(rs1("hotel_amount")) then 
    hotel_amount = rs1("hotel_amount")
end if

petrol_amt = trip_amt + exmileage_amount 'this includes petrol claims for all technicians
totalperjob = labor + petrol_amt + parkingtoll + hotel_amount

%>
  <tr bgcolor="<%=nbgcolor%>">
      <td height="40" align="center"><%=rs1("job_date")%></td>
      <td align="left" nowrap="nowrap"><%=rs1("job_code")%></td>
      <td align="left" bgcolor="#FFFFFF"><%=ChkDate(rs1("job_donedate"))%></td>
      <td align="center" bgcolor="#F3F3F3"><%=ChkDate(rs1("job_purchase_date"))%></td>
      <td align="center" bgcolor="#F3F3F3"></td>
      <td align="center"><%=rs1("job_model_desc")%></td>
      <td align="center" bgcolor="#FFFFFF"><%=ChkNumber2Decimal(rs1("job_totallabourAmt"))%></td>
      <td align="center" bgcolor="#F3F3F3"><%=ChkNumber2Decimal(petrol_amt)%></td>
      <td align="center" bgcolor="#F3F3F3"><%=ChkNumber2Decimal(rs1("parkingtoll_amount"))%></td>
      <td align="center" bgcolor="#F3F3F3"><%=ChkNumber2Decimal(rs1("hotel_amount"))%></td>
      <td align="center" bgcolor="#F3F3F3"><%=ChkNumber2Decimal(totalperjob)%></td>
      <td align="center" bgcolor="#FFFFFF"><%=rs1("job_inv_no")%></td>
  </tr>
  <%

if isnumeric(rs1("job_totallabourAmt")) then 
totalservice = totalservice + rs1("job_totallabourAmt")
end if

if isnumeric(petrol_amt) then 
totaltrip = totaltrip + petrol_amt
end if

if isnumeric(rs1("parkingtoll_amount")) then 
totalparkingtoll = totalparkingtoll + rs1("parkingtoll_amount")
end if

if isnumeric(rs1("hotel_amount")) then 
totalhotel = totalhotel + rs1("hotel_amount")
end if

grandtotal = totalservice + totaltrip + totalparkingtoll + totalhotel

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

%>
  <tr bgcolor="#F3F3F3">
                   <tr bgcolor="#F3F3F3">
                     <tr bgcolor="#F3F3F3">
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Total (RM)</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totalservice)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totaltrip)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totalparkingtoll)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totalhotel)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(grandtotal)%></td>                 
                     <td align="right" bgcolor="#CCCCCC"></td>
                   </tr>

                   </tr>
  </tr>
</table>

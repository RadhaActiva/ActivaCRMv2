<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
job_actual_wrty_status = request("job_actual_wrty_status")
tech_code = request("tech_code")
jobyear = request("jobyear")
jobmonth = request("jobmonth")

sql = "select tech_name FROM tbltechnician WHERE tech_code ='" & tech_code & "'" 
tech_name = selectid(sql)
%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Schedule C </font>ACS Monthly Claim Statement</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">
                      <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <td><strong>For Contractor : <%=tech_code%></strong>&nbsp;<%=tech_name%></td>
                    <tr>
                      <td width="80%" class="titlegrey1">&nbsp;</td>
                        <td width="20%" align="center" class="titlegrey1">
                            <a href="rm_rpt_tech_month_Arm_excel.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&tech_code=<%=tech_code%>&tech_name=<%=tech_name%>" target="_blank">
                            <img src="images/excel.jpg" width="57" height="21" border="0" /></a>
                     </td>
                    </tr>
                    <td><strong>For Period : <%=jobmonth%>/<%=jobyear%></strong></td>
                  </td>
                    </table>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_tech_month_ACS.asp">
                  </form></td>
                </tr>        
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="6%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>JS Date</span></strong></font></td>
                      <td width="10%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>SJS No</span></strong></font></td>
                      <td width="10%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Complete Date</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Purchase Date<br /></span></strong></font></td>
                      <td width="15%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Dealer Name<br /></span></strong></font></td>
                      <td width="40%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Model / Item</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Service</strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Trip</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Parking/Toll</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Hotel</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total</span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Inv</span></strong></font></td>
                    </tr>
                    
<%
i = 1
'sql2 = "SELECT rpc_id, rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_serviceQty1, rpc_serviceAmt1, rpc_serviceQty2, rpc_serviceAmt2, " & _
'		"rpc_techfees, rpc_car_allow, rpc_phone_allow, rpc_toll, rpc_parking, rpc_petrol, rpc_hotel, rpc_service_allow, rpc_overwarranty_fee, rpc_others,  " & _
'		"rpc_deduction_ow, rpc_deduction_sparepart, rpc_total,rpc_submitted_date,rpc_checkedby,rpc_checked_date,rpc_verifiedby,rpc_verified_date,rpc_approvedby,rpc_approved_date " & _
'		"FROM tblrpr_techcommission where rpc_tech_type='IHT' and rpc_month=" & jobmonth & " and rpc_year = " & jobyear & " order by rpc_tech_code "
  
'sql2 ="Select tbljob.job_tech_code, tbljob.job_code, tbljob.job_date, tbljob.job_donedate, tbljob.job_purchase_date, tbljob.job_model_desc, tbljob.job_totallabourAmt, tbljob.job_tech_sn, tbljob.job_inv_no, " & _ 
'   " (select sum(tp_claim_amount) from tbltech_claim_petrol where tp_job_sheet = tbljob.job_code and tp_tech_code=tbljob.job_tech_code and tp_month ='" & jobmonth & "' and tp_year ='" & jobyear & "') as trip_amount,  " & _
'   " (select sum(tpt_parking_amount+tpt_toll_amount) from tbltech_claim_parkingtoll where tpt_job_sheet = tbljob.job_code and tpt_tech_code=tbljob.job_tech_code and tpt_month ='" & jobmonth & "' and tpt_year ='" & jobyear & "') as parkingtoll_amount, " & _
'   " (select sum(th_claim_amount) from tbltech_claim_hotel where th_job_sheet = tbljob.job_code and th_tech_code=tbljob.job_tech_code and th_month ='" & jobmonth & "' and th_year ='" & jobyear & "') as hotel_amount, " & _
'   " (select sum(te_claim_amount) from tbltech_claim_exmileage where te_job_sheet = tbljob.job_code and te_tech_code=tbljob.job_tech_code and te_month ='" & jobmonth & "' and te_year ='" & jobyear & "') as exmileage_amount from tbljob  " & _
'   "join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
'   "and month(job_submitforclaims_date) = '" & jobmonth & "' and year(job_submitforclaims_date) = '" & jobyear & "' and job_submitforclaims='Yes' and tbljob.job_tech_code='" & tech_code & "' " & _
'   "and job_actual_wrty_status='Under' and job_submitforclaims='Yes' and job_status='Posted'"


 sql2  = "select * from (Select tbljob.job_tech_code, tbljob.job_code, tbljob.job_date, tbljob.job_donedate, tbljob.job_purchase_date, tbljob.job_model_desc, tbljob.job_totallabourAmt, tbljob.job_tech_sn, " & _
"tbljob.job_inv_no, " & _
"(select sum(tp_claim_amount) from tbltech_claim_petrol where tp_job_sheet = tbljob.job_code and tp_tech_code=tbljob.job_tech_code and tp_month =tbltech_claim.tc_month and tp_year = tbltech_claim.tc_year) as trip_amount,  " & _
"(select sum(tpt_parking_amount+tpt_toll_amount) from tbltech_claim_parkingtoll where tpt_job_sheet = tbljob.job_code " & _
"and tpt_tech_code=tbljob.job_tech_code and tpt_month =tbltech_claim.tc_month and tpt_year =tbltech_claim.tc_year) as parkingtoll_amount, " & _
"(select sum(th_claim_amount) from tbltech_claim_hotel where th_job_sheet = tbljob.job_code and th_tech_code=tbljob.job_tech_code and th_month =tbltech_claim.tc_month and th_year =tbltech_claim.tc_year) as hotel_amount, " & _
"(select sum(te_claim_amount) from tbltech_claim_exmileage where te_job_sheet = tbljob.job_code and te_tech_code=tbljob.job_tech_code and te_month =tbltech_claim.tc_month " & _
"and te_year =tbltech_claim.tc_year) as exmileage_amount " & _
"from tbljob join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
"join tbltech_claim ON  tbltechnician.tech_code = tbltech_claim.tc_tech_code " & _
"where tc_year = tbltech_claim.tc_year and tc_month=tbltech_claim.tc_month AND tc_tech_code ='" & tech_code & "' and tc_year_process = " & jobyear & " and tc_month_process = " & jobmonth & " " & _
"and tbljob.job_id is not null and tbljob.job_submitforclaims='Yes' and tbljob.job_claim_approved is NOT NULL " & _
"and tbljob.job_tech_code='" & tech_code & "' and job_actual_wrty_status='Under' and job_submitforclaims='Yes' and job_status='Posted') A where  A.job_totallabourAmt > 1 or A.trip_amount > 1 or A.parkingtoll_amount > 1 or A.hotel_amount > 1 or A.exmileage_amount > 1"
  
totalservice = 0
totaltrip = 0
totalparkingtoll = 0
grandtotal = 0
response.Cookies("GAPS")("sqlexcel") = sql2

i = 1	   
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql2
rs1.CursorLocation  = 3
rs1.Open
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
                    <tr bgcolor="#FFFFFF">
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
                      <td align="center" bgcolor="#FFFFFF"><%=rs1("job_inv_no")%><%=anyfigure%></td>
                    </tr>
<%'end if %>

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

i = i + 1
rs1.movenext
wend
rs1.close
%>       
                   <tr bgcolor="#F3F3F3">
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Total</strong></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totalservice)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totaltrip)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totalparkingtoll)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(totalhotel)%></td>
                     <td align="center" nowrap="nowrap" bgcolor="#CCCCCC"><%=chknumber2(grandtotal)%></td>                 
                     <td align="right" bgcolor="#CCCCCC"></td>
                   </tr>
                </table></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
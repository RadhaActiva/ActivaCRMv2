<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=IncentiveTechnician_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->

<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = request("job_actual_wrty_status")
tech_type = request("tech_type")
sparepartPer = request("sparepartPer")
labourPer = request("labourPer")
job_tech_code = request("job_tech_code")
job_date_from = request("job_date_from")
job_date_to = request("job_date_to")
jobmonth = request("jobmonth")
jobyear = request("jobyear")

sql="select top 1 tech_name from tbltechnician where tech_code='" & job_tech_code & "'"
tech_name = selectid(sql)
%>

<table width="80" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="11" class="style1"><strong>Over Warranty Incentives Calculation<br />
                      Technician Type : <%=tech_type%> - <%=tech_name%></strong></td>
                    </tr>
                      <tr>
                      <td colspan="11" class="style1"><strong>For the month of MM/YY : <%=jobmonth%> / <%=jobyear%></strong></td>
                    </tr>
                     <tr>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Invoice No</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>JS No</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Model</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total (A) <br/>Contractor Amt <br />Received <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Labor <br/> Charge <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Overwriting <br/> % (Labour)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F5B041" class="style1"><font color="#FFFFFF"><strong>Labour <br/> Payout <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Spare Part <br/> Charge <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Overwriting <br/>  % (Spare <br/>Part)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F5B041" class="style1"><font color="#FFFFFF"><strong>Sparepart <br/> Payout <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total (B) <br/> Payout <br/>(RM)</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Net Riegen Marketing <br/>Received (RM)<br/> Total (A) - (B)</strong></font></td>
                    </tr>   
<%
i = 1
sql1 = request.Cookies("GAPS")("sqlexcel") 
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

'totalsparepart = rs("job_totalPartsAmt") * sparepartPer/100
'totallabourpart = rs("job_totallabourAmt") * labourPer/100

'for grand total figures
gtotal_received = gtotal_received + rs1("inc_received_amt")
gtotal_labourCharge = gtotal_labourCharge + rs1("inc_labor_charge_amt")
gtotal_labourpayout = gtotal_labourpayout + rs1("inc_labor_payout")
gtotal_partsamount = gtotal_partsamount + rs1("inc_part_charge")
gtotal_sparepartscharge = gtotal_sparepartscharge + rs1("inc_part_payout")
gtotal_Bpayout =  gtotal_Bpayout + rs1("inc_total_payout")
gtotal_netReceived = gtotal_netReceived + rs1("inc_net_received")

%>
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center" valign="top" nowrap="nowrap"><%=i%></td>
                      <td align="left" valign="top" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs1("inc_invoice_no")%></font></strong></td>
                      <td align="left" valign="top" nowrap="nowrap"><strong><font color="#0000FF"><%=rs1("inc_jobno")%></font></strong></td>
                      <td align="left" valign="top" nowrap="nowrap"><%=rs1("md_desc")%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_received_amt"))%></td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_labor_charge_amt"))%></td>
					  <td align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("inc_labor_overwriting")%>%</td>
                      <td align="center" valign="top" nowrap="nowrap" bgcolor="#F5B041"><%=chknumber2(rs1("inc_labor_payout"))%></td> 
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_part_charge"))%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs1("inc_part_overwriting")%>%</td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="#F5B041"><%=chknumber2(rs1("inc_part_payout"))%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_total_payout"))%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs1("inc_net_received"))%></td>
                    </tr>   
<%

i = i + 1
rs1.movenext
wend
rs1.close
%>
<tr bgcolor="<%=nbgcolor%>">
  <td height="40" colspan="4" align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Grand Total</strong></td>
  <td height="40" align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_received)%></strong></td>
  <td height="40" align="center" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_labourCharge)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"></td>
  <td align="right" nowrap="nowrap" bgcolor="#F5B041"><strong><%=chknumber2(gtotal_labourpayout)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_partsamount)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"></td>
  <td align="right" nowrap="nowrap" bgcolor="#F5B041"><strong><%=chknumber2(gtotal_sparepartscharge)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_Bpayout)%></strong></td>
  <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(gtotal_netReceived)%></strong></td>
</tr>
</table>
<tr>
<td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
</tr>
<tr><td align="right"><strong>G/Total <%=chknumber2(gtotal_Bpayout)%></strong>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
<tr><td align="right">(Riegen to pay out)&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>

<table width="80" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
</tr>
<tr>
<td width ="20%"></td>
<td><strong><u>OVER WARRANTY INCENTIVE FEES</u></strong></td>
</tr>
<tr>
<td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
</tr>
<tr>
<td></td>
<td>a) Riegen Marketing Received  :</td>
<td><%=chknumber2(gtotal_received)%></td>
</tr>
<tr>
<td></td>
<td>b) Contractor Overiding  :</td>
<td><%=chknumber2(gtotal_Bpayout)%></td>
</tr>
<tr>
<td></td>
<td>c) Balance (A-B)  :</td>
<td><%=chknumber2(gtotal_netReceived)%></td>
</tr>
</table>


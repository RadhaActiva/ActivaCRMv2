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

sql="select top 1 tech_name from tbltechnician where tech_code='" & job_tech_code & "'"
tech_name = selectid(sql)
%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td colspan="11" class="style1"><strong>Over Warranty Incentives Calculation<br />
                      Technician Type: <%=tech_type%> (Home Appliances and Services / <%=tech_name%> )</strong></td>
                    </tr>
                    <tr>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Invoice No</span></strong></font></td>
                      <td align="left" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>JS No</span></strong></font></td>
                      <td align="center" nowrap="nowrap" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                      <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Spare Part <br />
                      / Labour</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total <br />
                        Contractor <br />
                      Amt Received</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Total Riegen Marketing <br />
                      AMT Received</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Percentage<br />
                        %
                      </span></strong></font></td>
                      <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Model</span></strong></font></td>
                      <td align="right" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total Payout<br />
                      (RM)</strong></font></td>
                    </tr>
<%
i = 1
sql1 = request.Cookies("GAPS")("sqlexcel") 
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql1
rs.CursorLocation  = 3
rs.Open
while not rs.eof 

'totalsparepart = rs("job_totalPartsAmt") * sparepartPer/100
'totallabourpart = rs("job_totallabourAmt") * labourPer/100

technician_tech_type = rs("tech_type")
job_jobcode = rs("job_code")
job_count = rs("job_count")

if technician_tech_type = "TPC" then 
	if (job_count = 1) then
		labourPer = "75"
	end if
	if (job_count > 1) then
		labourPer = "100"
	end if		
end if 

if technician_tech_type = "IHT" then 
	if (job_count = 1) then
		labourPer = "50"
	end if 
	if (job_count > 1) then
		labourPer = "100"
	end if	
end if 

if technician_tech_type = "IHC" then 		
	if (job_count = 1) then
		labourPer = "0"
	end if 
	if (job_count > 1) then
		labourPer = "0"
	end if	
end if 



totalsparepart = rs("job_totalPartsAmt") * (sparepartPer/100)
totallabourpart = rs("job_totallabourAmt") * (labourPer/100)

%>
              <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" rowspan="2" align="center" valign="top" nowrap="nowrap"><%=i%></td>
                      <td rowspan="2" align="left" valign="top" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs("job_inv_no")%></font></strong></td>
                      <td rowspan="2" align="left" valign="top" nowrap="nowrap"><strong><font color="#0000FF"><%=chkdate(rs("inv_date"))%></font></strong></td>
                      <td rowspan="2" align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs("job_code")%></td>
                      <td rowspan="2" align="center" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=rs("job_tech_code")%></td>
                      <td align="left" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><strong>Spare Part</strong></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="#E5E5E5">&nbsp;</td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(rs("job_totalPartsAmt"))%></td>
                      <td align="right" valign="top" nowrap="nowrap"><%=sparepartPer%></td>
                      <td align="center" valign="top"><%=rs("job_tech_model_desc")%></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(totalsparepart)%></td>
                    </tr>
                    <tr bgcolor="<%=nbgcolor%>">
                      <td align="left" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><strong>Labour</strong></td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="#E5E5E5">&nbsp;</td>
                      <td align="right" valign="top" nowrap="nowrap"><%=chknumber2(rs("job_totallabourAmt"))%></td>
                      <td align="right" valign="top" nowrap="nowrap"><%=labourPer%></td>
                      <td align="right" valign="top" nowrap="nowrap">&nbsp;</td>
                      <td align="right" valign="top" nowrap="nowrap" bgcolor="<%=nbgcolor%>"><%=chknumber2(totallabourpart)%></td>
                    </tr>
              <%
t_all = t_all + (ccur(totallabourpart) + ccur(totalsparepart))
t_ftcrmreceived = t_RMreceived + (rs("job_totalPartsAmt") + rs("job_totallabourAmt"))
i = i + 1
rs.movenext
wend
rs.close


%>

            <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" colspan="6" align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong>Total</strong></td>
                      <td height="40" align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(t_RMreceived)%></strong></td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC">&nbsp;</td>
                      <td align="right" nowrap="nowrap" bgcolor="#CCCCCC"><strong><%=chknumber2(t_all)%></strong></td>
                    </tr>
</table>

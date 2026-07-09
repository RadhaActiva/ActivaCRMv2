<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = request("job_actual_wrty_status")

if ordertype = "" then 
   ordertype = "desc"
end if

if request("job_date_from") <> "" then
   job_date_from = request("job_date_from")
else
   job_date_from = chkdate(DateAdd("d",-90,date()))
end if

if request("job_date_to") <> "" then
   job_date_to = request("job_date_to")
else
   job_date_to = chkdate(date())
end if

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

%> 
     
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>Technician Monthly Claim Report - TPC Technician</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1">&nbsp;</td>
                      <td width="20%" align="center" class="titlegrey1"><img src="images/excel.jpg" width="57" height="21" /></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="action_report.asp?type=monthcommisionTPC">
                    <select name="jobyear" id="jobyear">
                        <option value="2021"<%if jobyear="2021" then response.write " selected"%>>2021</option>
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
                    </select>
                    <span class="titlegrey1">
                    <input type="submit" name="button" id="button3" value="Generate Report" />
                     &nbsp;&nbsp;[NOTE : Run Generate Incentive (Over-Warranty) Report before running this report]
                    </span>
                  </form></td>
                </tr>
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="34" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="204" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Technician</span></strong></font></td>
                      <td width="276" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Technician Name</span></strong></font></td>
                      <td width="54" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Service 1st Unit<br />
                      </span></strong></font></td>
                      <td width="55" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Service 2nd Onward<br />
                      </span></strong></font></td>
                      <td width="55" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Petrol<br />
                        (Extra Milege)
                      </span></strong></font></td>
                      <td width="42" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Other</span></strong></font></td>
                      <td width="71" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Deduction<br />
                        Over Wrty
                        <br />
                      </span></strong></font></td>
                      <td width="71" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Deduction<br />
                        Spare Part
                      </span></strong></font></td>
                      <td width="39" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                      <td width="8" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                    </tr>
                    
<%
i = 1
sql2 = "SELECT rpc_id, rpc_month, rpc_year, rpc_tech_code, rpc_tech_name, rpc_serviceQty1, rpc_serviceAmt1, rpc_serviceQty2, rpc_serviceAmt2, " & _
		"rpc_techfees, rpc_car_allow, rpc_phone_allow, rpc_toll, rpc_parking, rpc_petrol, rpc_hotel, rpc_service_allow, rpc_overwarranty_fee, rpc_others,  " & _
		"rpc_deduction_ow, rpc_deduction_sparepart, rpc_total,rpc_submitted_date,rpc_checkedby,rpc_checked_date,rpc_verifiedby,rpc_verified_date " & _
		"FROM tblrpr_techcommission where rpc_tech_type='TPC' and rpc_month=" & jobmonth & " and rpc_year = " & jobyear & " order by rpc_tech_code "
'response.write sql2
'response.End()
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

%>
                    <tr bgcolor="#FFFFFF">
                      <td height="40" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><strong><font color="#0000FF"><a href="rm_rpt_tech_monthcommisionTPC_print.asp?rpc_id=<%=rs1("rpc_id")%>&jobyear=<%=jobyear%>&jobmonth=<%=jobmonth%>&tech_code=<%=rs1("rpc_tech_code")%>" target="_blank"><%=rs1("rpc_tech_code")%></a></font></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><%=rs1("rpc_tech_name")%></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=rs1("rpc_serviceAmt1")%></strong></td>
                      <td align="center"><strong><%=rs1("rpc_serviceAmt2")%></strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong><%=rs1("rpc_petrol")%></strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong><%=rs1("rpc_others")%></strong></td>
                      <td align="center" bgcolor="#FFFFFF"><strong><%=rs1("rpc_deduction_ow")%></strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong><%=rs1("rpc_total")%></strong></td>
                      <td align="center" bgcolor="#F3F3F3"><strong><%=ChkNumber2(rs1("rpc_total"))%></strong></td>
                       <%if  rs1("rpc_verifiedby") <> "" then %>
                            <td align="center" bgcolor="33FCFF"><strong>Verified</strong></td>
                       <%else if  (rs1("rpc_verifiedby") = "" or isnull(rs1("rpc_verifiedby"))) and rs1("rpc_checkedby") <> "" then %>
                            <td align="center" bgcolor="E3FF33"><strong>Checked</strong></td>
                       <%else if  (rs1("rpc_verifiedby") = "" or isnull(rs1("rpc_verifiedby"))) and (rs1("rpc_checkedby") = "" or isnull(rs1("rpc_checkedby"))) then %>
                            <td align="center" bgcolor="#F3F3F3"><strong>New</strong></td>
                       <%end if %>
                       <%end if %>
                      <%end if %>
                    </tr>
                    
<%
i = i + 1
rs1.movenext
wend
rs1.close
%> 
     
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
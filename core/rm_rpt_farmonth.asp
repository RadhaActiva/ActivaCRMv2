<!-- #include file="header.asp" -->
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
updatemonth = request("updatemonth")

if request("TotalSales") <> "" then
   TotalSales = request("TotalSales")
else
   TotalSales = 100
end if

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

currentdate = "01-" & convertmonth(jobmonth) & "-" & jobyear

jobmonth1 = jobmonth
jobyear1 =  jobyear
jobmonth2 = month(DateAdd("m",-1,currentdate))
jobyear2 =  year(DateAdd("m",-1,currentdate))
jobmonth3 = month(DateAdd("m",-2,currentdate))
jobyear3 =  year(DateAdd("m",-2,currentdate))

if request("job_tech_model") <> "" then
   job_tech_model = replace(request("job_tech_model"), " ", "")
   arrjob_tech_model = split(job_tech_model,",")
   job_tech_model = replace(job_tech_model, ",", "','")
   
   listjob_tech_model = listjob_tech_model & job_tech_model
   
else
   listjob_tech_model = ""
   arrjob_tech_model = split("0,0",",")
   
end if

function checkModelList(strv)
for k = 0 to ubound(arrjob_tech_model)
    if arrjob_tech_model(k) = strv then 
	   checkModelList = true
	   exit for
	else
	   checkModelList = false
	end if
next
end function

%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>FAR (One Month)</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="80%" class="titlegrey1"> Failure Analysis Rate (FAR) by month                            
                      <label for="select"></label></td>
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_farmonth_excel.asp?jobmonth=<%=jobmonth%>&jobyear=<%=jobyear%>&TotalSales=<%=TotalSales%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="action_report.asp?type=rpt_farmonth">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Job Posted Month<br />
                        </strong></td>
                        <td colspan="4"><div align="left">
                          <select name="jobyear" id="jobyear">
                            <option value="2016"<%if jobyear="2016" then response.write " selected"%>>2016</option>
                            <option value="2017"<%if jobyear="2017" then response.write " selected"%>>2017</option>
                            <option value="2018"<%if jobyear="2018" then response.write " selected"%>>2018</option>
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
                        </div></td>
                      </tr>
                      <tr>
                        <td class="titlegrey1">Type                        </td>
                        <td width="30%" align="left"><span class="titlegrey1">Model</span></td>
                        <td width="13%" align="center"><span class="titlegrey1">Total Sales </span></td>
                        <td width="19%" align="center"><strong><span class="titlegrey1">Update Month</span></strong></td>
                        <td width="22%" rowspan="2" align="center"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span><span class="titlegrey1">
                        <br />
                        <br />
<input type="button" name="button2" id="button" value="Reset Report" onclick="document.location.href='action_report.asp?type=rpt_farmonth_reset&job_tech_type='+ form1.job_tech_type.options[form1.job_tech_type.selectedIndex].value + '&jobmonth=' + form1.jobmonth.value + '&jobyear=' + form1.jobyear.value " />
                        </span></td>
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="job_tech_type" id="job_tech_type">
                          <option value="CF" <%if job_tech_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                          <option value="WH" <%if job_tech_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                        </select></td>
                        <td width="30%" align="left"><span class="titlegrey1">
                          <select name="job_tech_model" size="6" multiple="multiple" id="job_tech_model">
                            <option value="" <%if job_tech_model="" then response.write " selected"%>>All Model</option>
                            <%			
				sql = "SELECT md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_unitprice FROM tblmodel where (md_category = 'FAN : CEILING FAN' or md_category = 'WHEAT : WATER HEATER') and md_code is not null "
				
				if job_tech_type <> "" then 
				sql = sql & " and md_type='" & job_tech_type & "' "
				end if
				
				sql = sql & " order by md_code "	
				
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if checkModelList(rs("md_code")) then
					  response.write "<option value='" & rs("md_code") & "' selected>" & rs("md_code") & " - " & rs("md_desc")  & "</option>"
					  else
					  response.write "<option value='" & rs("md_code") & "'>" & rs("md_code") & " - " & rs("md_desc")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                          </select>
                        </span></td>
                        <td width="13%" align="center" valign="top"><label for="TotalSales"></label>
                        <input name="TotalSales" type="text" id="TotalSales" value="<%=TotalSales%>" size="10" maxlength="10" /></td>
                        <td align="center" valign="top">
                          <select name="updatemonth" id="updatemonth">
                            <option value="jobmonth1" <%if updatemonth="jobmonth1" then response.write " selected"%>>jobmonth1</option>
                            <option value="jobmonth2" <%if updatemonth="jobmonth2" then response.write " selected"%>>jobmonth2</option>
                            <option value="jobmonth3" <%if updatemonth="jobmonth3" then response.write " selected"%>>jobmonth3</option>
                            <option value="jobmonth4" <%if updatemonth="jobmonth4" then response.write " selected"%>>jobmonth4</option>
                            <option value="jobmonth5" <%if updatemonth="jobmonth5" then response.write " selected"%>>jobmonth5</option>
                            <option value="jobmonth6" <%if updatemonth="jobmonth6" then response.write " selected"%>>jobmonth6</option>
                            <option value="jobmonth7" <%if updatemonth="jobmonth7" then response.write " selected"%>>jobmonth7</option>
                            <option value="jobmonth8" <%if updatemonth="jobmonth8" then response.write " selected"%>>jobmonth8</option>
                            <option value="jobmonth9" <%if updatemonth="jobmonth9" then response.write " selected"%>>jobmonth9</option>
                            <option value="jobmonth10" <%if updatemonth="jobmonth10" then response.write " selected"%>>jobmonth10</option>
                            <option value="jobmonth11" <%if updatemonth="jobmonth11" then response.write " selected"%>>jobmonth11</option>
                            <option value="jobmonth12" <%if updatemonth="jobmonth12" then response.write " selected"%>>jobmonth12</option>
                        </select></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
                    <tr>
                      <td align="center" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(jobmonth1) & " " & jobyear1%></strong></font></td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">MD</font></strong></td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>DS</strong></font></td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">WI</font></strong></td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CF</strong></font></td>
                    </tr>
                    <tr>
                      <td width="4%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="10%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Fault(s)</span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                      </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                      </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                      </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                      </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                      </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                      </span></strong></font></td>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                    </tr>
                    
<%
job_tech_model = replace(job_tech_model, "'", "")
i = 1
sql1 = "SELECT id, faulth_code, faulth_desc, fa_month1_over, fa_month1_under, fa_month2_over, fa_month2_under, fa_month3_over, fa_month3_under, " & _
	   "fa_month4_over, fa_month4_under, fa_month5_over, fa_month5_under, fa_month6_over, fa_month6_under, fa_month7_over, fa_month7_under,  " & _
	   "fa_month8_over, fa_month8_under, fa_month9_over, fa_month9_under, fa_month10_over, fa_month10_under, fa_month11_over, fa_month11_under,  " & _
	   "fa_month12_over, fa_month12_under, fa_month_total_over, fa_month_total_under, fa_MD1_over, fa_MD1_under, fa_MD2_over, fa_MD2_under, fa_MD3_over,  " & _
	   "fa_MD3_under, fa_MD4_over, fa_MD4_under, fa_MD5_over, fa_MD5_under, fa_MD6_over, fa_MD6_under, fa_MD7_over, fa_MD7_under, fa_MD8_over,  " & _
	   "fa_MD8_under, fa_MD9_over, fa_MD9_under, fa_MD10_over, fa_MD10_under, fa_MD11_over, fa_MD11_under, fa_MD12_over, fa_MD12_under, fa_DS1_over,  " & _
	   "fa_DS1_under, fa_DS2_over, fa_DS2_under, fa_DS3_over, fa_DS3_under, fa_DS4_over, fa_DS4_under, fa_DS5_over, fa_DS5_under, fa_DS6_over,  " & _
	   "fa_DS6_under, fa_DS7_over, fa_DS7_under, fa_DS8_over, fa_DS8_under, fa_DS9_over, fa_DS9_under, fa_DS10_over, fa_DS10_under, fa_DS11_over, " & _ 
	   "fa_DS11_under, fa_DS12_over, fa_DS12_under, fa_WI1_over, fa_WI1_under, fa_WI2_over, fa_WI2_under, fa_WI3_over, fa_WI3_under, fa_WI4_over,  " & _
	   "fa_WI4_under, fa_WI5_over, fa_WI5_under, fa_WI6_over, fa_WI6_under, fa_WI7_over, fa_WI7_under, fa_WI8_over, fa_WI8_under, fa_WI9_over,  " & _
	   "fa_WI9_under, fa_WI10_over, fa_WI10_under, fa_WI11_over, fa_WI11_under, fa_WI12_over, fa_WI12_under, fa_CF1_over, fa_CF1_under, fa_CF2_over,  " & _
	   "fa_CF2_under, fa_CF3_over, fa_CF3_under, fa_CF4_over, fa_CF4_under, fa_CF5_over, fa_CF5_under, fa_CF6_over, fa_CF6_under, fa_CF7_over,  " & _
	   "fa_CF7_under, fa_CF8_over, fa_CF8_under, fa_CF9_over, fa_CF9_under, fa_CF10_over, fa_CF10_under, fa_CF11_over, fa_CF11_under, fa_CF12_over,  " & _
	   "fa_CF12_under, fa_MD_over, fa_MD_under, fa_DS_over, fa_DS_under, fa_WI_over, fa_WI_under, fa_CF_over, fa_CF_under FROM tblrpr_farmonth where id is not null and (fa_month_total_over > 0 or fa_month_total_under > 0) order by fa_month_total_over desc, fa_month_total_under desc  "
		   
Response.Cookies("GAPS")("sqlexcel") = sql1
Response.Cookies("GAPS")("TotalSales") = TotalSales

'response.write sql1

set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
                     <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"> <%=i%> </td>
                      <td align="left"><strong> <font color="#0000FF"><%=rs1("faulth_desc")%> </font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=<%=jobmonth1%>&jobyear=<%=jobyear1%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month1_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=<%=jobmonth1%>&jobyear=<%=jobyear1%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month1_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_MD_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_MD_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_MD_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_MD_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_DS_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_DS_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_DS_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_DS_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_WI_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_WI_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_WI_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_WI_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_CF_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_CF_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_CF_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_CF_under")%></a></font></strong></td>
                    </tr>
                    
<%
fa_month1_over = fa_month1_over + rs1("fa_month1_over")
fa_month1_under = fa_month1_under + rs1("fa_month1_under")
fa_month2_over = fa_month2_over + rs1("fa_month2_over")
fa_month2_under = fa_month2_under + rs1("fa_month2_under")
fa_month3_over = fa_month3_over + rs1("fa_month3_over")
fa_month3_under = fa_month3_under + rs1("fa_month3_under")
fa_month_total_over = fa_month_total_over + rs1("fa_month_total_over")
fa_month_total_under = fa_month_total_under + rs1("fa_month_total_under")

fa_MD_over = fa_MD_over + rs1("fa_MD_over")
fa_MD_under = fa_MD_under + rs1("fa_MD_under")
fa_DS_over = fa_DS_over + rs1("fa_DS_over")
fa_DS_under = fa_DS_under + rs1("fa_DS_under")
fa_WI_over = fa_WI_over + rs1("fa_WI_over")
fa_WI_under = fa_WI_under + rs1("fa_WI_under")
fa_CF_over = fa_CF_over + rs1("fa_CF_over")
fa_CF_under = fa_CF_under + rs1("fa_CF_under")

total_over = total_over + rs1("fa_month_total_over")
total_under = total_under + rs1("fa_month_total_under") 

if rs1("faulth_desc") ="Set Tested OK"  then
   setok_over = setok_over + rs1("fa_month_total_over")
   setok_under = setok_under + rs1("fa_month_total_under")
end if

if rs1("faulth_desc") ="Cancel Service" then 
   cancel_service_over = cancel_service_over + rs1("fa_month_total_over")
   cancel_service_under = cancel_service_under + rs1("fa_month_total_under")
end if 


i = i + 1
rs1.movenext
wend
rs1.close

fa_MD = (cint(fa_MD_over)+cint(fa_MD_under))
fa_DS = (cint(fa_DS_over)+cint(fa_DS_under))
fa_WI = (cint(fa_WI_over)+cint(fa_WI_under))
fa_CF = (cint(fa_CF_over)+cint(fa_CF_under))

grandtotal = total_over+total_under
totalsetok_over = setok_over + setok_under
totalcancel_service = cancel_service_over + cancel_service_under
totalokcancel = totalsetok_over +  totalcancel_service
%>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="2" align="right" bgcolor="#FFFFFF"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=<%=jobmonth1%>&jobyear=<%=jobyear1%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_month1_over%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=<%=jobmonth1%>&jobyear=<%=jobyear1%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_month1_under%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#97ACFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_month_total_over%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#97ACFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_month_total_under%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_MD_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_MD_over%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_MD_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_MD_under%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_DS_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_DS_over%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_DS_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_DS_under%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_WI_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_WI_over%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_WI_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_WI_under%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_CF_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_CF_over%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=all&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_CF_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=fa_CF_under%></a></font></strong></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td width="48%" height="30" align="left" bgcolor="#FFFFFF"><table width="60%" border="0" cellpadding="4" cellspacing="0">
                    <tr bgcolor="#FFFFFF">
                      <td width="10%" height="20" align="left" nowrap="nowrap"><strong> Total Sales </strong></td>
                      <td width="5%" height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=TotalSales%></strong></td>
                      <!--Open-->                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Total Reject </strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=grandtotal%></strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>% of Reject</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=chknumber2((grandtotal/TotalSales)*100)%> %</strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Total Set Tested OK  &amp; Cancel Service</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=totalokcancel%></strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>% Set Tested OK  &amp; Cancel Service</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> <%=chknumber2((totalokcancel/TotalSales)*100)%> %</strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" bgcolor="#FFFFFF"><strong>% Reject Actual Failure<br />
                        <br />
                      </strong>((total Reject - Total Set Tested OK  &amp; Cancel Service)/Total Sales )x 100</td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong> 
					  <%=chknumber2(((grandtotal-totalokcancel)/TotalSales)*100)%> %</strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>No of MD</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=fa_MD%></strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>No of DS</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=fa_DS%></strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>No of WI</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=fa_WI%></strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>No of CF</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=fa_CF%></strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Over Warranty</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=total_over%></strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Under Warranty</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><%=total_under%></strong></td>
                      <!--Open-->                    </tr>
                  </table></td>
                  <td width="52%" align="left" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr bgcolor="#FFFFFF">
                      <td width="10%" height="20" align="left" nowrap="nowrap"><strong> % MD </strong></td>
                      <td width="5%" height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
                        <%
					  if TotalSales > 0 then 
					     response.write chknumber2((fa_MD/TotalSales)*100)
					  else
					     response.write "0"
					  end if	 
						 %>
                        %</strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>% DS</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
                        <%
					  if TotalSales > 0 then 
					     response.write chknumber2((fa_DS/TotalSales)*100)
					  else
					     response.write "0"
					  end if	 
						 %>
                        %</strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>% WI</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
                        <%
					   if TotalSales > 0 then 
					     response.write chknumber2((fa_WI/TotalSales)*100)
					   else
					     response.write "0"
					   end if	 
						 %>
                        %</strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>% CF</strong></td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong>
                        <%
					  if TotalSales > 0 then 
					     response.write chknumber2((fa_CF/TotalSales)*100)
					  else
					     response.write "0"
					  end if 
						 %>
                        %</strong></td>
                      <!--Open-->
                    </tr>
                    <tr bgcolor="#F3F3F3">
                      <td height="20" align="left" nowrap="nowrap" bgcolor="#FFFFFF">&nbsp;</td>
                      <td height="20" align="center" nowrap="nowrap" bgcolor="#F3F3F3">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
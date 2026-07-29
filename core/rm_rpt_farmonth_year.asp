<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        #button3 {
            height: 43px;
            width: 160px;
        }
    </style>
</head>
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")

if request("TotalSales") <> "" then
   TotalSales = request("TotalSales")
else
   TotalSales = 100
end if

if ordertype = "" then 
   ordertype = "desc"
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

<script language="javascript">

    function confirmForm() {

        document.form1.action = "rm_rpt_farmonth_year.asp";
        document.form1.method = "get";
        document.form1.submit();
    }

    function filterModels() {
        var input = document.getElementById("modelSearch");
        var filter = input.value.toUpperCase();
        var select = document.getElementById("job_tech_model");

        for (var i = 0; i < select.options.length; i++) {
            var option = select.options[i];
            var text = option.text.toUpperCase();
            // Keep option visible if it matches filter OR it's selected
            if (text.includes(filter) || option.selected) {
                option.style.display = "";
            } else {
                option.style.display = "none";
            }
        }
    }

    function clearSelection() {
        var select = document.getElementById("job_tech_model");
        for (var i = 0; i < select.options.length; i++) {
            select.options[i].selected = false;
            select.options[i].style.display = ""; // Reset all to visible
        }
        document.getElementById("modelSearch").value = "";
    }

    // -->
</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report </font>FAR (One Year) - Faulty Code</div></td>
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
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="action_report_farmonthyear.asp?type=rpt_farmonth_year_reset">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" height="20" nowrap="nowrap" class="titlegrey1"><strong> Job Month<br />
                        </strong></td>
                        <td colspan="3"><div align="left">
                          <select name="jobyear" id="jobyear">
                            <option value="2016"<%if jobyear="2016" then response.write " selected"%>>2016</option>
                            <option value="2017"<%if jobyear="2017" then response.write " selected"%>>2017</option>
                            <option value="2018"<%if jobyear="2018" then response.write " selected"%>>2018</option>
                            <option value="2019"<%if jobyear="2019" then response.write " selected"%>>2019</option>
							<option value="2020"<%if jobyear="2020" then response.write " selected"%>>2020</option>
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
                        </div></td>
                      </tr>
                      <tr>
                        <td class="titlegrey1">Type                        </td>
                        <td width="14%"><span class="titlegrey1">Model</span></td>
                       
                        
                      </tr>
                      <tr>
                        <td valign="top" class="titlegrey1"><select name="job_tech_type" id="job_tech_type" onchange="javascript:confirmForm();">
                          <option value="CF" <%if job_tech_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                          <option value="WH" <%if job_tech_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                        </select>
                            <br />
                            <br />
                         <button type="button" onclick="clearSelection()" style="margin-bottom: 10px;">Clear Selection</button>

                        </td>
                        <td width="14%"><span class="titlegrey1">
                            <input type="text" id="modelSearch" placeholder="Search model..." onkeyup="filterModels()" style="width: 100%; margin-bottom: 5px;">
                            <select name="job_tech_model" multiple size="10" id="job_tech_model">
                       <!--   <select name="job_tech_model" size="6" multiple="multiple" id="job_tech_model">-->
                            <option value="" <%if job_tech_model="" then response.write " selected"%>>All Model</option>
                              <%			
				sql = "SELECT md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_unitprice FROM tblmodel where (md_category = 'FAN : CEILING FAN' or md_category = 'WHEAT : WATER HEATER' or md_category ='SWHEA : STORAGE WATER HEATER') and md_code is not null "
				
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
                        <td width="24%" align="center" valign="top"><label for="TotalSales"></label>
                        Total Sales &nbsp;<input name="TotalSales" type="text" id="TotalSales" value="<%=TotalSales%>" size="10" maxlength="10" />&nbsp;&nbsp;&nbsp;<span class="titlegrey1"><input type="submit" name="button1" id="button3" value="Generate Report" /></span></td>
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
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(jobmonth3) & " " & jobyear3%></strong></font></td>
                      <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(jobmonth2) & " " & jobyear2%></strong></font></td>
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
sql1 = "SELECT id, faulth_code, faulth_desc, fa_month1_over, fa_month1_under, fa_month2_over, fa_month2_under, " & _
		"fa_month3_over, fa_month3_under, fa_month4_over, fa_month4_under, fa_month5_over, fa_month5_under, fa_month6_over, fa_month6_under,  " & _
		"fa_month7_over, fa_month7_under, fa_month8_over, fa_month8_under, fa_month9_over, fa_month9_under, fa_month10_over, fa_month10_under,  " & _
		"fa_month11_over, fa_month11_under, fa_month12_over, fa_month12_under, fa_month_total_over, fa_month_total_under, fa_MD_over, fa_MD_under,  " & _
		"fa_DS_over, fa_DS_under, fa_WI_over, fa_WI_under, fa_CF_over, fa_CF_under " & _
		"FROM tblrpr_farmonth where id is not null and (fa_month_total_over > 0 or fa_month_total_under > 0) order by fa_month_total_over desc, fa_month_total_under desc "
    
Response.Cookies("GAPS")("sqlexcel") = sql1
Response.Cookies("GAPS")("TotalSales") = TotalSales

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
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=<%=jobmonth3%>&jobyear=<%=jobyear3%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month3_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=<%=jobmonth3%>&jobyear=<%=jobyear3%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month3_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=<%=jobmonth2%>&jobyear=<%=jobyear2%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month2_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=<%=jobmonth2%>&jobyear=<%=jobyear2%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month2_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=<%=jobmonth1%>&jobyear=<%=jobyear1%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month1_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=<%=jobmonth1%>&jobyear=<%=jobyear1%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month1_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=Under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_MD_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_MD_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_MD_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_MD_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_DS_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_DS_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_DS_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_DS_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_WI_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_WI_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_WI_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_WI_under")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_CF_over&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_CF_over")%></a></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_farmonth_year_detail.asp?jobmonth=0&jobyear=<%=jobyear%>&faulth_code=<%=rs1("faulth_code")%>&job_tech_type=<%=job_tech_type%>&job_tech_model=<%=job_tech_model%>&stype=fa_CF_under&lastmonth=<%=jobmonth1%>','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_CF_under")%></a></font></strong></td>
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
'totalsetok_over = setok_over + setok_under
'totalcancel_service = cancel_service_over + cancel_service_under
'totalokcancel = totalsetok_over +  totalcancel_service

total_under_warranty = fa_MD_under+fa_DS_under+fa_WI_under+fa_CF_under
total_over_warranty = fa_MD_over+fa_DS_over+fa_WI_over+fa_CF_over

md_under_reject = (fa_MD_under/TotalSales)*100
ds_under_reject = (fa_ds_under/TotalSales)*100
wi_under_reject = (fa_wi_under/TotalSales)*100
cf_under_reject = (fa_cf_under/TotalSales)*100
total_under_reject_rate=md_under_reject+ds_under_reject+wi_under_reject+cf_under_reject

md_over_reject = (fa_MD_over/TotalSales)*100
ds_over_reject = (fa_ds_over/TotalSales)*100
wi_over_reject = (fa_wi_over/TotalSales)*100
cf_over_reject = (fa_cf_over/TotalSales)*100
total_over_reject_rate=md_over_reject+ds_over_reject+wi_over_reject+cf_over_reject

total_md_reject_per=md_under_reject+md_over_reject
total_ds_reject_per=ds_under_reject+ds_over_reject
total_wi_reject_per=wi_under_reject+wi_over_reject
total_cf_reject_per=cf_under_reject+cf_over_reject
total_reject_per = total_md_reject_per + total_ds_reject_per + total_wi_reject_per + total_cf_reject_per

total_MD_reject_units=fa_MD_under+fa_MD_over
total_ds_reject_units=fa_ds_under+fa_ds_over
total_wi_reject_units=fa_wi_under+fa_wi_over
total_cf_reject_units=fa_cf_under+fa_cf_over
total_reject_units = total_md_reject_units+total_ds_reject_units+total_wi_reject_units+total_cf_reject_units
%>
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="2" align="right" bgcolor="#FFFFFF"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month3_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month3_under%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=fa_month2_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=fa_month2_under%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month1_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month1_under%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#97ACFF"><strong><font color="#0000FF"><%=fa_month_total_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#97ACFF"><strong><font color="#0000FF"><%=fa_month_total_under%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_MD_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_MD_under%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=fa_DS_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=fa_DS_under%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_WI_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_WI_under%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=fa_CF_over%></font></strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=fa_CF_under%></font></strong></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="30" colspan="2" align="right" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                    <table width="60%" border="1" cellpadding="4" cellspacing="0">
                        <tr>
                            <td width="40%" height="20" align="left" nowrap="nowrap"><strong> Sales Quantity </strong></td>
                            <td height="60%" nowrap="nowrap" bgcolor="#FFFFFF" align="center" colspan="6"><strong>Total Reject Rate </strong></td>
                        </tr>
                        <tr>
                            <td width="40%" height="20" align="left" nowrap="nowrap" valign="top"><strong> <%=TotalSales%> </strong></td>
                            <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF" rowspan="2"><strong>Under <br />Warranty</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>MD</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>DS</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>WI</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>CF</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Total</strong></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td><%=fa_MD_under%></td>
                            <td><%=fa_DS_under%></td>
                            <td><%=fa_WI_under%></td>
                            <td><%=fa_CF_under%></td>
                            <td><%=total_under_warranty%></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Reject Rate %</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(md_under_reject,1)%>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(ds_under_reject,1) %>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(wi_under_reject,1) %>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(cf_under_reject,1) %>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_under_reject_rate,1)%>%</strong></td>
                        </tr>
                        <tr><td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td></tr>
                        <tr>
                            <td width="40%" height="20" align="left" nowrap="nowrap"></td>
                            <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF" rowspan="2"><strong>Over <br />Warranty</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>MD</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>DS</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>WI</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>CF</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Total</strong></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td><%=fa_MD_over%></td>
                            <td><%=fa_DS_over%></td>
                            <td><%=fa_WI_over%></td>
                            <td><%=fa_CF_over%></td>
                            <td><%=total_over_warranty%></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Reject Rate %</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(md_over_reject,1) %>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(ds_over_reject,1) %>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(wi_over_reject,1) %>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(cf_over_reject,1) %>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_over_reject_rate,1)%>%</strong></td>
                        </tr>
                          <tr><td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td></tr>
                        <tr>
                            <td width="40%" height="20" align="left" nowrap="nowrap"></td>
                            <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF" rowspan="2"><strong>TOTAL REJECT <br/>(U/W + O/W)</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>MD</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>DS</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>WI</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>CF</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Total</strong></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td><%=total_md_reject_units%></td>
                            <td><%=total_ds_reject_units%></td>
                            <td><%=total_wi_reject_units%></td>
                            <td><%=total_cf_reject_units%></td>
                            <td><%=total_reject_units%></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong>Reject Rate %</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_md_reject_per,1)%>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_ds_reject_per,1)%>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_wi_reject_per,1)%>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_cf_reject_per,1)%>%</strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_reject_per,1)%>%</strong></td>
                        </tr>
                    </table>
                  <!--</table></td>
                  
              </tr>
              </table></td>-->
        </tr>
<!-- #include file="footer.asp" -->
<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            height: 26px;
        }
    </style>
</head>
<%
job_tech_type = request("job_tech_type")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
job_actual_wrty_status = "Under"

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

if request("job_cust_state1") <> "" then
    sql = "select  state_name FROM tblstate where state_id=" & request("job_cust_state1")
    job_cust_state1 = request("job_cust_state1")
	job_cust_state1_name = selectid(sql)
else
    job_cust_state1 = "0"
end if

if request("job_cust_state2") <> "" then
    sql = "select  state_name FROM tblstate where state_id=" & request("job_cust_state2")
    job_cust_state2 = request("job_cust_state2")
	job_cust_state2_name = selectid(sql)
else
    job_cust_state2 = "0"
end if

if request("job_cust_state3") <> "" then
    sql = "select  state_name FROM tblstate where state_id=" & request("job_cust_state3")
    job_cust_state3 = request("job_cust_state3")
	job_cust_state3_name = selectid(sql)
else
    job_cust_state3 = "0"
end if

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
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Report 10 </font>Most Problem under Warranty</div></td>
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
                      <td width="20%" align="center" class="titlegrey1"><a href="rm_rpt_10mostprobUnder_excel.asp?job_date_from=<%=job_date_from%>&job_date_to=<%=job_date_to%>" target="_blank"><img src="images/excel.jpg" width="57" height="21" border="0" /></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="rm_rpt_10mostprobUnder.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                        <td class="titlegrey1"><strong>Job Month</strong></td>
                        <td colspan="2"><strong><font color="#000000"><strong>
                          <input name="job_date_from" type="text" id="job_date_from" value="<%=job_date_from%>" size="15" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_from;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>to
                          <input name="job_date_to" type="text" id="job_date_to" value="<%=job_date_to%>"
                                            size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.form1.job_date_to;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"> <img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></strong> Date must be (dd-MMM-yyyy) eg: 21-May-2015 </td>
                      </tr>
                      <tr>
                        <td class="titlegrey1" width="16%">Stk Type</td>
                          <td valign="top" class="titlegrey1">
                          <select name="job_tech_type" id="job_tech_type">
                          <option value="">All</option>
                          <option value="CF" <%if job_tech_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                          <option value="WH" <%if job_tech_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                          </select> 
                          </td>
                         </tr>

                        <tr>
                        <td width="14%" class="auto-style1"><span class="titlegrey1">Model</span></td>
                            <td width="14%" class="auto-style1"><span class="titlegrey1">
                          <select name="job_tech_model" id="job_tech_model">
                            <option value="" <%if job_tech_model="" then response.write " selected"%>>All</option>
                            <%			
				                sql = "SELECT md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_unitprice FROM tblmodel where md_code is not null "
				
				                if job_tech_type <> "" then 
				                sql = sql & " and md_type='" & job_tech_type & "' "
				                end if
				                
                                sql = sql & " order by md_code "	

                                set rs = server.CreateObject("adodb.recordset")
				                rs.Open sql,strconnect,3,3,&H0001
                                while Not rs.EOF
					                  if job_tech_model = rs("md_code") then
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

                        <td width="24%" align="center" class="auto-style1"></td>
                        <td rowspan="5"><span class="titlegrey1">
                          <input type="submit" name="button" id="button3" value="Generate Report" />
                        </span></td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td align="left" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="97%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td width="5%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td width="15%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Item Code.</span></strong></font></td>
                      <td width="30%" align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Model Description</span></strong></font></td>
                      <td width="13%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Type<br />
                      </span></strong></font></td>
                      <td width="13%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Sales Quanlity<br />
                      </span></strong></font></td>
                      <td width="10%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Cases of Fault</strong></font></td>
                      <td width="9%" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>% of Fault</span></strong></font></td>
                    </tr>
                    
<%
i = 1


sql2 = "SELECT tbljob.job_Model, tblmodel.md_desc, tbljob.job_tech_type, count(tbljob.job_id) as totaljob FROM tbljob inner join tblmodel on tbljob.job_Model = tblmodel.md_code " & _
          "left join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code " & _
	      "where tbljob.job_id is not null and tbljob.job_status='Posted' " & _
		  "and  tbljob.job_date >= '" & ChkDateYYYYMMDD(job_date_from) & "' and tbljob.job_date <= '" & ChkDateYYYYMMDD(job_date_to) & "' " 
		  
	if job_actual_wrty_status <> "" then 
	   sql2 = sql2 & " and job_actual_wrty_status = '" & job_actual_wrty_status & "' "
	end if
	
	if job_tech_type <> "" then 
	   sql2 = sql2 & " and job_tech_type = '" & job_tech_type & "' "
	end if
	
	if job_tech_model <> "" then 
	   sql2 = sql2 & " and job_model in ( '" & job_tech_model & "') "
	end if
	
	   sql2 = sql2 & "group by tbljob.job_Model, tblmodel.md_desc, tbljob.job_tech_type "
	   
	   if orderby <> "" then
	      sql2 = sql2 & " order by totaljob desc "
	   else
          sql2 = sql2 & " order by totaljob desc "
	   end if
	   
Response.Cookies("GAPS")("sqlexcel") = sql2
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
                   <tr bgcolor="<%=nbgcolor%>">
                      <td height="40" align="center"><%=i%></td>
                      <td align="left" nowrap="nowrap"><strong> <font color="#0000FF"><%=rs1("job_Model")%></font></strong></td>
                      <td align="left" nowrap="nowrap"><%=rs1("md_desc")%></td>
                      <td align="center" nowrap="nowrap"><%=rs1("job_tech_type")%></td>
                      <td align="center" nowrap="nowrap"><strong> 0</strong></td>
                      <td align="center" nowrap="nowrap"><strong> <%=rs1("totaljob")%></strong></td>
                      <td align="center"><strong>0</strong></td>
                    </tr>
<%
totaljob = totaljob + cint(rs1("totaljob"))
i = i + 1
rs1.movenext
wend
rs1.close

grandtotal = total_over+total_under
%>
                    
                    <tr bgcolor="#F3F3F3">
                      <td height="40" colspan="4" align="right" bgcolor="#999999"><strong>Total</strong></td>
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong> 0</strong></td>
                      <!--Open-->
                      <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=totaljob%></strong></td>
                      <td align="center" bgcolor="#999999"><strong>0</strong></td>
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
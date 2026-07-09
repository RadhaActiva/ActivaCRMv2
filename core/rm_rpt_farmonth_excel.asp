<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=farmonth_" & year(date()) & month(date()) & day(date()) & ".xls"
%>
<!-- #include file="database/datastore.asp" -->

<%
jobmonth = request("jobmonth")
jobyear = request("jobyear")
if request("job_tech_model") <> "" then
   listjob_tech_model = "'" & replace(request("job_tech_model"), " ", "") & "'"
else
   listjob_tech_model = ""
end if

currentdate = "01-" & convertmonth(jobmonth) & "-" & jobyear
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="2" valign="top" bgcolor="#FFFFFF"><table border="1" cellpadding="4" cellspacing="0">
              <tr>
                <td colspan="36" align="left" class="style1"><strong>Failure Analysis Rate (FAR) by month</strong></td>
              </tr>
              <tr>
                <td colspan="36" align="left" class="style1"><strong>Model / Category: <br>
				<%
				sql1 = "SELECT md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_unitprice FROM tblmodel " & _
				      " where md_code is not null "
					  
				if listjob_tech_model <> "" then 
				sql1 = sql1 & " and md_code in (" & listjob_tech_model & ") " 
				end if	  
				
				sql1 = sql1 & " order by md_code "	
				
				'response.write sql1 
                set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
					  response.write rs1("md_desc") & " - " & rs1("md_category")  & "<br>"
				rs1.movenext
				wend
				rs1.close	
				%></strong></td>
              </tr>
              <tr>
                <td align="center" bgcolor="#666666" class="style1">&nbsp;</td>
                <td align="left" bgcolor="#666666" class="style1">&nbsp;</td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-11,currentdate))) & " " & year(DateAdd("m",-11,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-10,currentdate))) & " " & year(DateAdd("m",-10,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-9,currentdate))) & " " & year(DateAdd("m",-9,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-8,currentdate))) & " " & year(DateAdd("m",-8,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-7,currentdate))) & " " & year(DateAdd("m",-7,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-6,currentdate))) & " " & year(DateAdd("m",-6,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-5,currentdate))) & " " & year(DateAdd("m",-5,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-4,currentdate))) & " " & year(DateAdd("m",-4,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-3,currentdate))) & " " & year(DateAdd("m",-3,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-2,currentdate))) & " " & year(DateAdd("m",-2,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(month(DateAdd("m",-1,currentdate))) & " " & year(DateAdd("m",-1,currentdate))%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><%=convertmonth(jobmonth) & " " & jobyear%></strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Total</strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">MD</font></strong></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>DS</strong></font></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">WI</font></strong></td>
                <td colspan="2" align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>CF</strong></font></td>
              </tr>
              <tr>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                <td align="left" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span>Fault(s)</span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong><span> Over<br />
                </span></strong></font></td>
                <td align="center" bgcolor="#666666" class="style1"><font color="#FFFFFF"><strong>Under</strong></font></td>                
              </tr>
              <%
i = 1
sql1 = request.Cookies("GAPS")("sqlexcel") 
TotalSales = request.Cookies("GAPS")("TotalSales") 
set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof 
%>
              <tr>
                <td height="40" align="center"><%=i%></td>
                <td align="left"><strong> <font color="#0000FF"><%=rs1("faulth_desc")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month12_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month12_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=rs1("fa_month11_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=rs1("fa_month11_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month10_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month10_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month9_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month9_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month8_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month8_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month7_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month7_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month6_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month6_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month5_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month5_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month4_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month4_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month3_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month3_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month2_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month2_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month1_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_month1_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><%=rs1("fa_month_total_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><%=rs1("fa_month_total_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_MD_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_MD_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=rs1("fa_DS_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=rs1("fa_DS_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_WI_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><%=rs1("fa_WI_under")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=rs1("fa_CF_over")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><%=rs1("fa_CF_under")%></font></strong></td>
              </tr>
              <%
fa_month1_over = fa_month1_over + ChkNumber(rs1("fa_month1_over"))
fa_month1_under = fa_month1_under + ChkNumber(rs1("fa_month1_under"))
fa_month2_over = fa_month2_over + ChkNumber(rs1("fa_month2_over"))
fa_month2_under = fa_month2_under + ChkNumber(rs1("fa_month2_under"))
fa_month3_over = fa_month3_over + ChkNumber(rs1("fa_month3_over"))
fa_month3_under = fa_month3_under + ChkNumber(rs1("fa_month3_under"))
fa_month4_over = fa_month4_over + ChkNumber(rs1("fa_month4_over"))
fa_month4_under = fa_month4_under + ChkNumber(rs1("fa_month4_under"))
fa_month5_over = fa_month5_over + ChkNumber(rs1("fa_month5_over"))
fa_month5_under = fa_month5_under + ChkNumber(rs1("fa_month5_under"))
fa_month6_over = fa_month6_over + ChkNumber(rs1("fa_month6_over"))
fa_month6_under = fa_month6_under + ChkNumber(rs1("fa_month6_under"))
fa_month7_over = fa_month7_over + ChkNumber(rs1("fa_month7_over"))
fa_month7_under = fa_month7_under + ChkNumber(rs1("fa_month7_under"))
fa_month8_over = fa_month8_over + ChkNumber(rs1("fa_month8_over"))
fa_month8_under = fa_month8_under + ChkNumber(rs1("fa_month8_under"))
fa_month9_over = fa_month9_over + ChkNumber(rs1("fa_month9_over"))
fa_month9_under = fa_month9_under + ChkNumber(rs1("fa_month9_under"))
fa_month10_over = fa_month10_over + ChkNumber(rs1("fa_month10_over"))
fa_month10_under = fa_month10_under + ChkNumber(rs1("fa_month10_under"))
fa_month11_over = fa_month11_over + ChkNumber(rs1("fa_month11_over"))
fa_month11_under = fa_month11_under + ChkNumber(rs1("fa_month11_under"))
fa_month12_over = fa_month12_over + ChkNumber(rs1("fa_month12_over"))
fa_month12_under = fa_month12_under + ChkNumber(rs1("fa_month12_under"))
fa_month_total_over = fa_month_total_over + ChkNumber(rs1("fa_month_total_over"))
fa_month_total_under = fa_month_total_under + ChkNumber(rs1("fa_month_total_under"))

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

total_under_warranty = fa_MD_under+fa_DS_under+fa_WI_under+fa_CF_under
total_over_warranty = fa_MD_over+fa_DS_over+fa_WI_over+fa_CF_over

md_under_reject = (fa_MD_under/TotalSales)*100
ds_under_reject = (fa_ds_under/TotalSales)*100
wi_under_reject = (fa_wi_under/TotalSales)*100
cf_under_reject = (fa_cf_under/TotalSales)*100
total_under_reject_rate=md_under_reject+ds_under_reject+wi_under_reject+cf_under_reject+it_under_reject

md_over_reject = (fa_MD_over/TotalSales)*100
ds_over_reject = (fa_ds_over/TotalSales)*100
wi_over_reject = (fa_wi_over/TotalSales)*100
cf_over_reject = (fa_cf_over/TotalSales)*100
total_over_reject_rate=md_over_reject+ds_over_reject+wi_over_reject+cf_over_reject+it_over_reject

total_md_reject_per=md_under_reject+md_over_reject
total_ds_reject_per=ds_under_reject+ds_over_reject
total_wi_reject_per=wi_under_reject+wi_over_reject
total_cf_reject_per=cf_under_reject+cf_over_reject
total_reject_per = total_md_reject_per + total_ds_reject_per + total_wi_reject_per + total_cf_reject_per + total_it_reject_per

total_MD_reject_units=fa_MD_under+fa_MD_over
total_ds_reject_units=fa_ds_under+fa_ds_over
total_wi_reject_units=fa_wi_under+fa_wi_over
total_cf_reject_units=fa_cf_under+fa_cf_over
total_reject_units = total_md_reject_units+total_ds_reject_units+total_wi_reject_units+total_cf_reject_units+total_it_reject_units
%>
              <tr bgcolor="#F3F3F3">
                <td height="40" colspan="2" align="center" bgcolor="#FFFFFF"><strong>Total</strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month12_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month12_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month11_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month11_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month10_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month10_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month9_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month9_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month8_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month8_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month7_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month7_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month6_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month6_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month5_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month5_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month4_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month4_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month3_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month3_under%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month2_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month2_under%></font></strong></td>
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
            <td width="48%" height="30" align="left" bgcolor="#FFFFFF">
                
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
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(md_under_reject,1)%></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(ds_under_reject,1) %></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(wi_under_reject,1) %></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(cf_under_reject,1) %></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_under_reject_rate,1)%></strong></td>
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
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(md_over_reject,1) %></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(ds_over_reject,1) %></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(wi_over_reject,1) %></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(cf_over_reject,1) %></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_over_reject_rate,1)%></strong></td>
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
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_md_reject_per,1)%></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_ds_reject_per,1)%></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_wi_reject_per,1)%></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_cf_reject_per,1)%></strong></td>
                             <td height="60%" align="left" nowrap="nowrap" bgcolor="#FFFFFF"><strong><%=formatnumber(total_reject_per,1)%></strong></td>
                        </tr>
                    </table>      
            </table></td>
          </tr>
</table>

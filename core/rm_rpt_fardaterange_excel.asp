<%  
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", " filename=fardaterange_" & year(date()) & month(date()) & day(date()) & ".xls"

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

if request("TotalSales") <> "" then
   TotalSales = request("TotalSales")
else
   TotalSales = 100
end if

if request("job_tech_model") <> "" then
   listjob_tech_model = "'" & replace(request("job_tech_model"), " ", "") & "'"
else
   listjob_tech_model = ""
end if

%>
<!-- #include file="database/datastore.asp" -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="100%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="4" cellspacing="0">
              <tr>
                <td colspan="14" align="left" class="style1"><strong>Failure Analysis Rate (FAR) by Date Range</strong></td>
              </tr>
              <tr>
                <td colspan="14" align="left" class="style1">
                <strong>Model / Category: <br>
				<%
				sql1 = "SELECT md_id, md_code, md_desc, md_category, md_model, md_barcode, md_type, md_status, md_unitprice FROM tblmodel " & _
				      " where md_code is not null and (md_category = 'FAN : CEILING FAN' or md_category = 'WHEAT : WATER HEATER') "
					  
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
				%></strong>
                </td>
              </tr>
              <tr>
                <td colspan="4" align="left" bgcolor="#666666" class="style1"><strong><font color="#FFFFFF">Selected Date Range: <%=job_date_from%> to <%=job_date_to%></font></strong></td>
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
i = 1
sql1 = "SELECT id, faulth_code, faulth_desc, fa_month_total_over, fa_month_total_under, fa_MD_over, fa_MD_under,  " & _
		"fa_DS_over, fa_DS_under, fa_WI_over, fa_WI_under, fa_CF_over, fa_CF_under " & _
		"FROM tblrpr_fardaterange where id is not null order by fa_month_total_over desc, fa_month_total_under desc "

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
                <td height="40" align="center"><%=i%></td>
                <td align="left"><strong> <font color="#0000FF"><%=rs1("faulth_desc")%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=Over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_over")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=Under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_under")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=Over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_over")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#C6D1FF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=Under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_month_total_under")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_MD_over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_MD_over")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_MD_under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_MD_under")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_DS_over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_DS_over")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_DS_under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_DS_under")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_WI_over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_WI_over")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#F3F3F3"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_WI_under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_WI_under")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_CF_over','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_CF_over")%></a></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#FFFFFF"><strong><font color="#0000FF"><a href="javascript:popup('rm_rpt_fardaterange_detail.asp?job_date_from=<%=job_date_from%>&amp;job_date_to=<%=job_date_to%>&amp;faulth_code=<%=rs1("faulth_code")%>&amp;job_tech_type=<%=job_tech_type%>&amp;job_tech_model=<%=job_tech_model%>&amp;stype=fa_CF_under','cb18','scrollbars=yes,resizable=yes,width=600,height=500')"><%=rs1("fa_CF_under")%></a></font></strong></td>
              </tr>
              <%
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
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month_total_over%></font></strong></td>
                <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><font color="#0000FF"><%=fa_month_total_under%></font></strong></td>
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
</table>
<br />
<br />
<table>
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
                      <td height="20" align="left" bgcolor="#FFFFFF"><strong>% Reject Actual Failure</strong></td>
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
</table>

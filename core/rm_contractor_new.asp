<!-- #include file="header.asp" -->
<%

tech_code=request.querystring("tech_code") 
tech_name=request.querystring("tech_name")
tech_tel1=request.querystring("tech_tel1")
tech_tel2=request.querystring("tech_tel2")
tech_icno=request.querystring("tech_icno")
tech_address=request.querystring("tech_address")
tech_id=request.querystring("tech_id") 
tech_postcode = request.form("itech_postcode")
tech_salary="0.00"

if tech_postcode = "" then
    tech_postcode = request.QueryString("tech_postcode")
end if

set rs = server.CreateObject("adodb.recordset")

if request("tech_code") <> "" then 'and request("tech_id") <> ""  then	  
sql = "SELECT tech_id, tech_code, tech_type, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id, tech_wh_code, tech_salary," & _
	  "tech_creditlimit, tech_terms " & _
	  "FROM tbltechnician WHERE tech_code = '" & request("tech_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			tech_id = rs("tech_id") 
			tech_code = rs("tech_code") 
			tech_type = rs("tech_type") 
			tech_name = rs("tech_name") 
			tech_icno = rs("tech_icno")
			tech_address = rs("tech_address")
			tech_postcode = rs("tech_postcode") 
			tech_state = rs("tech_state") 
			tech_state_id = rs("tech_state_id") 
			tech_city = rs("tech_city") 
			tech_city_id = rs("tech_city_id") 
			tech_email = rs("tech_email") 
			tech_tel1 = rs("tech_tel1") 
			tech_tel2 = rs("tech_tel2") 
			tech_createdby = rs("tech_createdby") 
			tech_cretateddate = rs("tech_cretateddate") 
			tech_carmodel = rs("tech_carmodel")  
			tech_carplateno = rs("tech_carplateno") 
			tech_carcolour = rs("tech_carcolour") 
			tech_password = rs("tech_password") 
			tech_status = rs("tech_status") 
			tech_area = rs("tech_area")
			tech_area_id = rs("tech_area_id")
			tech_wh_code = rs("tech_wh_code")
			tech_creditlimit = rs("tech_creditlimit")
			tech_terms = rs("tech_terms")
            tech_salary = rs("tech_salary")
		End If
		rs.Close
	  'stype = "editTechnician"	
      stype = "addTechnician"
	  actionname = "Save" 
 else    
	  stype = "addTechnician"
	  actionname = "Save" 
end if

'tech_city_id=""
'tech_state_id=""
    
if tech_postcode <> "" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & tech_postcode & "' "
		rs1.Open sql1,strconnect,0,1,&H0001   
		If Not rs1.EOF Then
             tech_state_id = rs1("state_id") 'will auto populate state
             tech_state =  rs1("state_name")
             tech_city_id = rs1("city_id") 'will auto populate city
             tech_city = rs1("post_office")    
        end if
    rs1.close
end if
%>

<script language="javascript">

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
    {
	document.forminvoice.action = orderlinks;
	document.forminvoice.submit();
    } 
}

    function getPostcode(p) {
        //document.getElementById('cust_name').value = s;
        document.getElementById('tech_postcode').value = p;
        document.formorder.submit();  }

// -->
</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
           <table width="100%"><!--for UI Purpose-->
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Technician</div></td>
                        <td width="23%" align="right" class="titleblue1">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                <td colspan="2" bgcolor="#FFFFFF">
                <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                      <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
                </tr>
                <tr>
                  <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV4">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Technician  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician Code *</strong></font></td>
                        <td align="left"><input name="tech_code" type="text" id="tech_code" value="<%=tech_code%>" size="20" maxlength="50" />
                          <input type="hidden" name="tech_id" id="tech_id" value="<%=tech_id%>" /></td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type</strong></font></td>
                        <td align="left">
                          <select name="tech_type" id="tech_type">
                            <option value="IC" <%if tech_type="IC" then response.write " selected"%>>IC</option>
                            <option value="IHT" <%if tech_type="IHT" then response.write " selected"%>>IHT</option>
                            <option value="TPC" <%if tech_type="TPC" then response.write " selected"%>>TPC</option>                    
                            <option value="SGT" <%if tech_type="SGT" then response.write " selected"%>>SGT</option>                    
                            <option value="Sales" <%if tech_type="Sales" then response.write " selected"%>>Sales</option>
                            <option value="Walk-in" <%if tech_type="Walk-in" then response.write " selected"%>>Walk-in</option>
                          </select></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician Name *</strong></font></td>
                        <td align="left"><label for="tech_icno"></label>
                          <input name="tech_name" type="text" id="tech_name" value="<%=tech_name%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician ICNO *</strong></font></td>
                        <td align="left"><label for="tech_name2"></label>
                          <input name="tech_icno" type="text" id="tech_name2" value="<%=tech_icno%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><strong>
                          <textarea name="tech_address" cols="50" rows="3" id="tech_address"><%=tech_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left"><strong>
                          <input name="tech_postcode" type="text" id="tech_postcode" value="<%=tech_postcode%>" onchange="getPostcode(this.value)" size="20" maxlength="50" />
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left">
                        <input name="tech_state" type="text" id="tech_state" value="<%=tech_state%>" size="30" readonly maxlength="50" />
                        <input name="tech_state_id" type="hidden" id="tech_state_id" value="<%=tech_state_id%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                        <input name="tech_city" type="text" id="tech_city" value="<%=tech_city%>" size="30" readonly maxlength="50" />
                        <input name="tech_city_id" type="hidden" id="tech_city_id" value="<%=tech_city_id%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="tech_email" type="text" id="tech_emai" value="<%=tech_email%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><label for="textfield9"></label>
                          <input name="tech_tel1" type="text" id="tech_tel1" value="<%=tech_tel1%>" maxlength="50" />
                          e.g 012-1234657</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="tech_tel2" type="text" id="tech_tel2" value="<%=tech_tel2%>" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                        <td valign="top"><%=tech_createdby%><br />
                        <%=chkdatetime(tech_cretateddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong>Other Info</strong></td>
                      </tr>
                      <tr>
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Location&nbsp; *</strong></font></td>
                        <td align="left"><select name="tech_area_id" id="tech_area_id" onchange="setOption(this.form, this.options[this.selectedIndex].value);">
                          <option value="0"></option>
                            <%			
				                sql = "SELECT state_id, state_cnty_id, state_code, state_name FROM tblstate order by state_name"	
                                set rs = server.CreateObject("adodb.recordset")
				                rs.Open sql,strconnect,3,3,&H0001
                                while Not rs.EOF
					                  if (tech_area_id) = (rs("state_id")) then
					                  response.write "<option value='" & rs("state_id") & "' selected>" & rs("state_name") & "</option>"
					                  else
					                  response.write "<option value='" & rs("state_id") & "'>" & rs("state_name") & "</option>"
					                  end if 					  
				                rs.movenext
				                wend
				                rs.close					
				            %>
                        </select></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Car Model</strong></font></td>
                        <td align="left"><input name="tech_carmodel" type="text" id="tech_carmodel" value="<%=tech_carmodel%>" maxlength="50" /></td>
                      </tr>
                     
                      <tr>
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Car Plate No. </strong></font></td>
                        <td align="left"><input name="tech_carplateno" type="text" id="tech_carplateno" value="<%=tech_carplateno%>" maxlength="50" /></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Car colour</strong></font></td>
                        <td><input name="tech_carcolour" type="text" id="tech_carcolour" value="<%=tech_carcolour%>" /></td>
                      </tr>
                       <tr>
                         <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF">Password</font></strong></td>
                         <td align="left"><input name="tech_password" type="text" id="tech_password" value="<%=tech_password%>" size="20" maxlength="50" /></td>
                         <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Status</font></strong></td>
                         <td><select name="tech_status" id="tech_status">
                           <option value="Y" <%if tech_status="Y" then response.write " selected"%>>Y</option>
                           <option value="N" <%if tech_status="N" then response.write " selected"%>>N</option>
                         </select></td>
                       </tr>
                       <tr>
                         <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF">Store&nbsp; *</font></strong></td>
                         <td align="left"><select name="tech_wh_code" id="tech_wh_code" style="width:100px">
                           <option value="0"></option>
                           <%			
				                sql = "SELECT wh_id, wh_code, wh_name FROM tblwarehouse order by wh_code"	
                                set rs = server.CreateObject("adodb.recordset")
				                rs.Open sql,strconnect,3,3,&H0001
                                while Not rs.EOF
					                  if (tech_wh_code) = (rs("wh_code")) then
					                  response.write "<option value='" & rs("wh_code") & "' selected>" & rs("wh_code") & " - " & rs("wh_name") & "</option>"
					                  else
					                  response.write "<option value='" & rs("wh_code") & "'>" & rs("wh_code") & " - " & rs("wh_name") & "</option>"
					                  end if 					  
				                rs.movenext
				                wend
				                rs.close					
				            %>
                         </select></td>
                         <td align="left" bgcolor="#CD6155">&nbsp;</td>
                         <td>&nbsp;</td>
                       </tr>
                       <tr>
                         <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF">Credit Limit&nbsp; *</font></strong></td>
                         <td align="left">RM
                          <input name="tech_creditlimit" type="text" id="tech_creditlimit" value="<%=tech_creditlimit%>" size="10" maxlength="50" /></td>
                         <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Terms</font></strong></td>
                         <td><input name="tech_terms" type="text" id="tech_terms" value="<%=tech_terms%>" size="10" maxlength="50" />days</td>
                       </tr>
                       <tr>
                            <td bgcolor="#CD6155"><strong><font color="#FFFFFF">Spareparts Incentive</font></strong></td>
                            <td align="left"><input name="tech_creditlimit2" type="text" id="tech_creditlimit2" value="<%=tech_creditlimit%>" size="10" maxlength="50" />
                            %</td>
                            <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Labour Incentive</font></strong></td>
                            <td><input name="tech_creditlimit3" type="text" id="tech_creditlimit3" value="<%=tech_creditlimit%>" size="10" maxlength="50" />%</td>
                      </tr>
                       <tr >
                            <td bgcolor="#CD6155"><strong><font color="#FFFFFF">Basic Salary</font></strong></td>
                            <td align="left"><input name="tech_salary" type="text" id="tech_salary" value="<%=tech_salary%>" size="10" maxlength="10" /> (For IHT)</td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
              </table>
                
                <p></p>
                <div>
                    <input align="right" type="submit" name="button" id="button" value="<%=actionname%>"/>
                </div>
                </form>
                <p></p>

                </td>
                </tr>
              <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                                     
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" scope="col">
                         <table width="100%" border="0" cellpadding="8" cellspacing="0">
                            <tr bgcolor="#CD6155">
                              <td colspan="10" bgcolor="#E8E8E8"><strong><font size="2">Related Jobs</font></strong></td>
                            </tr>

                            <tr bgcolor="#475387">
                              <td align="center"><font color="#FFFFFF"><strong>No</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong>Job No.</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong>Date</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong>Model</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong>SN</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong>Customer</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong>Appointment Date</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong> Status</strong></font></td>
                              <td align="center"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                              <td align="center"><strong><font color="#FFFFFF">City</font></strong></td>
                            </tr>
                        <%	i = 1
				            sql1 = "SELECT tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
				            "tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
				            "tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
				            "tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
				            "tbljob.job_Model, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
				            "tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
				            "tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN,  " & _
				            "tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
				            "tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo, tbljob.job_hq_remark,  " & _
				            "tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
				            "tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
				            "tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbltechnician.tech_name, tbltechnician.tech_tel1 " & _
				            "FROM tbljob inner join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null and " & _
				            "tbljob.job_tech_code = '" & tech_code & "' order by tbljob.job_id desc"
			    
                            set rs1 = server.CreateObject("adodb.recordset")
                            rs1.ActiveConnection = strconnect
                            rs1.Source = sql1
                            rs1.CursorLocation  = 3
                            rs1.Open
                            if rs1.eof then
                               norecord = "There is no record found."
                            end if

                            If Not rs1.EOF Then

                            if request("rowno") <> "" then
	                              row = cint(request("rowno"))
                            else
	                              row = 50
                            end if
			
                            Showed = Request("jobnum")
                            If Showed = "" Then Showed = 0
                            TotalRecord = rs1.RecordCount
                            Remain = TotalRecord - Showed

                            If Remain > row Then
                              LoopMax = Showed + row
                            Else
                              LoopMax = Showed + Remain
                            End If

	                        If Int(TotalRecord/row) <> TotalRecord/row Then
	                          pgCount = Int(TotalRecord/row) + 1
	                        Else
	                          pgCount = TotalRecord/row
	                        End If

	                        if LoopMax mod row = 0 then
		                        pagestartno = LoopMax/row
	                        else
		                        pagestartno = pgCount
	                        end if		
                        end if

                        count = count + Showed
                        link = "&tech_code=" & tech_code 

                        if not rs1.eof then
                        rs1.Move Showed
                        count = Showed + 1
                        end if

                        For j = Showed + 1 To LoopMax
                        %>

                                                <tr>
                                                  <td align="center"><%=j%>.</td>
                                                  <td align="center" nowrap="nowrap"><strong><a href="rm_jobsheet.asp?job_code=<%=rs1("job_code")%>" target="_blank"><%=rs1("job_code")%></a></strong></td>
                                                  <td align="center" nowrap="nowrap"><%=chkdate(rs1("job_date"))%></td>
                                                  <td align="center"><%=rs1("job_Model")%></td>
                                                  <td align="center"><%=rs1("job_SN_no")%></td>
                                                  <td align="center"><%=rs1("job_cust_name")%></td>
                                                  <td align="center"><%=chkdate(rs1("job_appointment_date"))%></td>
                                                  <td align="center"><%=rs1("job_status")%></td>
                                                  <td align="center"><%=rs1("job_tech_code") & "-" & rs1("tech_name") %></td>
                                                  <td align="center"><%=rs1("job_cust_city")%></td>
                                                </tr>
                                                <%	
                        count = count + 1 
                        i = i + 1
                        rs1.MoveNext
                        Next
                        'rs.Close
                        'Set rs = Nothing
                        %>
                         <tr>
                          <td colspan="10" align="right"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                            <%	
	                            i = 0
	                            For j = 1 To pgCount
				                            If CInt(Showed) = ((j-1) * row) Then
					                            Response.Write "<font color=#000000><b>"& j &"</b></font>"
				                            Else
						                            Response.Write " <a href='rm_contractor_new.asp?jobnum=" & (j-1) * row & link & "'>"& j &"</a>"
				                            End If
			                            If Not j = pgCount Then Response.Write " "
	                            i = i + 1
	                            Next
	
	                            If Remain > row Then
	                               Response.Write "<a href='rm_contractor_new.asp?jobnum=" & Showed+row & link & "'> Next >></a>"
	                            End If
                            %>
                          </td>
                          </tr>
                      </table></td>
                    </tr>
                    <tr><td colspan="2" align="right" bgcolor="#FFFFFF" scope="col">&nbsp;</td></tr>
                    <tr align="right"><td colspan="2" bgcolor="#FFFFFF"></td></tr>                   
                  </table></td>
                </tr>  
              </table>
              </td>
        </tr>
<!-- #include file="footer.asp" -->
<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("job_code") <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city,job_cust_cnty_id, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_model_desc, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_code, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_service_date, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo,job_wrty_photo2,job_wrty_photo3, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date, job_submittedby, job_submitteddate, " & _
		"job_doneby, job_donedate, job_postedby, job_posteddate, job_cancelledby, job_cancelleddate  " & _
	    "FROM tbljob WHERE job_code = '" & request("job_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			job_id = rs("job_id") 
			job_code = rs("job_code") 
			job_date = rs("job_date")
			job_cust_code = rs("job_cust_code")
			job_cust_name = rs("job_cust_name")
			job_cust_address = rs("job_cust_address") 
			job_cust_postcode = rs("job_cust_postcode") 
			job_cust_state = rs("job_cust_state") 
			job_cust_city = rs("job_cust_city") 
            job_cust_cnty_id = rs("job_cust_cnty_id") 
			job_cust_email = rs("job_cust_email") 
			job_cust_tel1 = rs("job_cust_tel1") 
			job_cust_tel2 = rs("job_cust_tel2")  
			job_remark = rs("job_remark")
			job_createddate = rs("job_createddate") 
			job_createdby = rs("job_createdby") 
			job_JS_receiveddate = rs("job_JS_receiveddate") 
			job_JS_receivedby = rs("job_JS_receivedby") 
			job_status = rs("job_status") 
            job_purchase_date = rs("job_purchase_date") 
			job_onlineWrtyNo = rs("job_onlineWrtyNo") 
			job_onlineWrtyStatus = rs("job_onlineWrtyStatus") 
			job_type = rs("job_type") 
			job_SN_no = rs("job_SN_no") 
			job_Model = rs("job_Model") 
			job_model_desc = rs("job_model_desc") 
			job_faulty_reason_cs = rs("job_faulty_reason_cs") 
			job_faulty_desc = rs("job_faulty_desc") 
			job_reportedby = rs("job_reportedby") 
			
			if isdate(rs("job_appointment_date")) then 
			job_appointment_date = rs("job_appointment_date") 
			end if
			
			job_appointment_time = rs("job_appointment_time") 
			job_tech_code = rs("job_tech_code") 
			job_appointment_remark = rs("job_appointment_remark") 
			job_emailsentdate = rs("job_emailsentdate") 
			job_emailsent = rs("job_emailsent") 
			job_smssentdate = rs("job_smssentdate") 
			job_smssent = rs("job_smssent") 
			job_tech_type = rs("job_tech_type") 
			job_tech_model = rs("job_tech_model")
			job_tech_model_desc = rs("job_tech_model_desc")  
			job_tech_tax_invoice = rs("job_tech_tax_invoice") 
			job_tech_SN = rs("job_tech_SN") 
			job_tech_faulty_code = rs("job_tech_faulty_code") 
			job_tech_faulty_reason = rs("job_tech_faulty_reason") 
			job_tech_faulty_action = rs("job_tech_faulty_action") 
			job_tech_status = rs("job_tech_status") 
			job_tech_product_collectdate = rs("job_tech_product_collectdate") 
			job_tech_returntoCustDate = rs("job_tech_returntoCustDate") 
			job_actual_wrty_status = rs("job_actual_wrty_status") 
			job_wrty_photo = rs("job_wrty_photo") 
            job_wrty_photo2 = rs("job_wrty_photo2") 
            job_wrty_photo3 = rs("job_wrty_photo3") 
			job_tech_logby  = rs("job_tech_logby")
			job_tech_logdate  = rs("job_tech_logdate")
			job_hq_remark = rs("job_hq_remark") 
			job_hq_category_code = rs("job_hq_category_code") 
			job_hq_received_date = rs("job_hq_received_date") 
			job_totalPartsAmt = rs("job_totalPartsAmt") 
			job_totallabourAmt = rs("job_totallabourAmt") 
			job_totaltransportAmt = rs("job_totaltransportAmt") 
			job_totalAmt = rs("job_totalAmt") 
			job_repair_date = rs("job_repair_date") 
			job_return_tech_date = rs("job_return_tech_date") 
			job_office_issueRemark = rs("job_office_issueRemark") 
			job_office_supervisor = rs("job_office_supervisor") 
			job_office_taxinvoice = rs("job_office_taxinvoice") 
			job_rcn_no = rs("job_rcn_no")
			job_rcn_Date = rs("job_rcn_Date")
			job_inv_no = rs("job_inv_no") 
			job_do_no = rs("job_do_no") 
			job_submittedby = rs("job_submittedby")
			job_submitteddate = rs("job_submitteddate")
			job_doneby = rs("job_doneby")
			job_donedate = rs("job_donedate")
			job_postedby = rs("job_postedby")
			job_posteddate = rs("job_posteddate")
			job_cancelledby = rs("job_cancelledby")
			job_cancelleddate = rs("job_cancelleddate")
		End If
		rs.Close
	  stype = "editjob"	
	  actionname = "Save" 
 else    
	  stype = "addorder"
	  actionname = "Save" 		
	  job_date = date()	 
	  job_status = "Open"   	
end if
  
      if job_emailsent = "" or isnull(job_emailsent) then 
	     job_emailsent = job_cust_email
	  end if
	  
	  if job_smssent = "" or isnull(job_smssent) then 
	     job_smssent = job_cust_tel1
	  end if
 
sql = "select cnty_name from tblcountry where cnty_id =" & job_cust_cnty_id	
job_cust_city_name = selectid(sql)
%>


<script language="javascript">


function confirmForm(id,orderlinks,otype)
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formorderparts.action = orderlinks;
	document.formorderparts.submit();
   }
}


function confirmJobDone(id,orderlinks) 
{

  if (confirm("Are you sure you want to confirm Job Done: " + id))
   {
	document.formorder2.action = orderlinks;
	document.formorder2.submit();
   }
}


function calctotal(unitprice,qty,discountamt,discounttype,subtotal) {

var discount = 0;
var temp = 0;

if (discounttype == "%")
{
	discount = unitprice * (discountamt/100);
	temp = (unitprice-discount) * qty;
	subtotal.value = temp.toFixed(2);
	}
else
{
	discount = discountamt;
	temp = (unitprice-discount) * qty;
	subtotal.value = temp.toFixed(2);
	}
}
// -->
</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Job Sheet</div></td>
                        <td width="23%" align="right" class="titleblue1"><a href="rm_jobsheet_new_print.asp?job_code=<%=job_code%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
                      </tr>
                    </table></td>
                </tr>
                
                <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td width="45%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Code *</strong></font></td>
                        <td align="left"><label for="job_cust_code"></label>
                       <%=job_cust_code%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
                        <td align="left"><%=job_cust_name%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><%=job_cust_address%>
                       </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left"><%=job_cust_postcode%>
                        </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left"><%=job_cust_state%>
                        </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left"><%=job_cust_city%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left"><%=job_cust_city_name%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><%=job_cust_email%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><%=job_cust_tel1%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><%=job_cust_tel2%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td valign="top"><%=job_remark%>
                        </td>
                      </tr>
                    <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                        <td valign="top"><%=job_createdby%> @ 
                          <%=chkdatetime(job_createddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>JS Received Ack</strong></font></td>
                        <td valign="top"><%=job_JS_receivedby%> @ <%=chkdatetime(job_JS_receiveddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Submitted by</strong></font></td>
                        <td valign="top"><%=job_submittedby%> @ <%=chkdatetime(job_submitteddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Done by</strong></font></td>
                        <td valign="top"><%=job_doneby%> @ <%=chkdatetime(job_donedate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Posted by</strong></font></td>
                        <td valign="top"><%=job_postedby%> @ <%=chkdatetime(job_posteddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cancelled  by</strong></font></td>
                        <td valign="top"><%=job_cancelledby%> @ <%=chkdatetime(job_cancelleddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="55%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="0" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Job Sheet Information</font></strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job No.</strong></td>
                        <td align="left" style="background-color: #cccccc;"><strong><%=job_code%>
                          <input type="hidden" name="job_code" id="job_code" value="<%=job_code%>" />
                        </strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job  
                          Date*</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(job_date)%>
                        </strong></font></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status*</strong></font></td>
                        <td align="left"><strong><%=job_status%></strong></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Purchase  Date</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(job_purchase_date)%>
                       </strong></font></td>
                      </tr>
                      <tr align="left" >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
                        <td> <%=job_onlineWrtyNo%></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                        <td><strong>
                        <%if job_onlineWrtyStatus="Over" then %>
                             <font color="#FF0000">Over</font>
                        <%else%>
                              <font color="#000000">Under</font>
                        <%end if%>
                        </strong></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type*</strong></font></td>
                        <td align="left"><%=job_type%></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Serial #</strong></font></td>
                        <td align="left"><%=job_SN_no%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Item code*</strong></font></td>
                        <td colspan="3" align="left"><%=job_Model%></td>
                        </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model Desc</strong></font></td>
                        <td colspan="3" align="left"><%=job_Model_desc%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
                        <td colspan="3" align="left"><%=job_faulty_reason_cs%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Desc*</strong></font></td>
                        <td colspan="3" align="left"><strong><%=job_faulty_desc%>
                        </strong></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Reported By*</strong></font></td>
                        <td colspan="2"><%=job_reportedby%></td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr >
                      <td bgcolor="#E8E8E8" colspan="4" align="center"><strong>Appoinment Details</strong></td>
                     </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Date*</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(job_appointment_date)%></strong></font></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Time </strong></font></td>
                        <td><%=job_appointment_time%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician*</strong></font></td>
                        <td colspan="3"><%=job_tech_code%>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Remark<br />
                        </strong></font></td>
                        <td colspan="3"><strong><%=job_appointment_remark%>
                        </strong></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                </form>
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
  
  <%
  if request.Cookies("GAPS")("slevel") = "technician" and job_status<>"Open" and job_status<>"Submitted"   then 
  %>              
	         <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                   
                    
                    <form name="formorder2" id="formorder2" method="post" action="action.asp?type=editjob_Technical" enctype="multipart/form-data">
                      <tr align="right">
                        <td colspan="2" valign="top" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                          <tbody>
                            <tr bgcolor="#E8E8E8">
                              <td colspan="4" align="left" bgcolor="#E8E8E8" scope="col"><strong>Technical Findings <a name="articletitle" id="articletitle"></a></strong></td>
                              </tr>
                            <tr >
                              <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Appointment Date*</strong></font></td>
                              <td align="left"><font color="#000000"><strong>
                                <input name="job_appointment_date" type="text" id="job_appointment_date" value="<%=chkdate(job_appointment_date)%>" size="12" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.formorder2.job_appointment_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                              <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Time </strong></font></td>
                              <td align="left"><input name="job_appointment_time" type="text" id="job_appointment_time" value="<%=job_appointment_time%>" />
                                e.g 3:30pm</td>
                            </tr>
                            <tr >
                              <td width="12%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type*</strong></font></td>
                              <td width="38%" align="left"><select name="job_tech_type" id="job_tech_type">
                                <option value=""></option>
                                <option value="CF" <%if job_tech_type="CF" then response.write " selected"%>>C-Ceiling Fan</option>
                                <option value="WH" <%if job_tech_type="WH" then response.write " selected"%>>W-Water Heater</option>
                              </select>                                <input type="hidden" name="job_code" id="job_code" value="<%=job_code%>" /></td>
                              <td width="29%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model*</strong></font></td>
                              <td width="21%" align="left">
                                <input name="job_tech_model" type="text" id="job_tech_model" value="<%=job_tech_model%>" size="20" maxlength="50" />
                                <%if job_status="Accepted" then %>
                                <label for="job_type2">[<a href="javascript:popup('rm_job_new_model_tech.asp?searchitem=md_code&searchvalue=<%=cust_code%>&formname=formorder2&fieldname=job_tech_model&md_type=' + formorder2.job_tech_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] 
                                <%end if%>
                                </label></td>
                              </tr>
                            <tr >
                              <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tax Invoice*</strong></font></td>
                              <td align="left"><select name="job_tech_tax_invoice" id="job_tech_tax_invoice">
                                <option value="No" <%if job_tech_tax_invoice="No" then response.write " selected"%>>No</option>
                                <option value="Yes" <%if job_tech_tax_invoice="Yes" then response.write " selected"%>>Yes</option>
                              </select></td>
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model Description </strong></font></td>
                              <td align="left"><label for="job_tech_status">
                                <input name="job_tech_model_desc" type="text" id="job_tech_model_desc" value='<%=job_tech_model_desc%>' size="40" maxlength="50" />
                              </label></td>
                              </tr>
                            <tr>
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>S/N*</strong></font></td>
                              <td align="left"><input name="job_tech_SN" type="text" id="job_tech_SN" value='<%=job_tech_SN%>' size="50" maxlength="50" /></td>
                              <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Service Status<strong>*</strong></font></strong></td>
                              <td align="left"><select name="job_tech_status" id="job_tech_status">
                                <option value="Pending" <%if job_tech_status="Pending" then response.write " selected"%>>Pending</option>
                                <option value="Done" <%if job_tech_status="Done" then response.write " selected"%>>Done</option>
                                <option value="Send Back to HQ" <%if job_tech_status="Send Back to HQ" then response.write " selected"%>>Send Back to HQ</option>
                              </select></td>
                              </tr>
                            <tr >
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
                              <td align="left"><select name="job_tech_faulty_code" id="job_tech_faulty_code" style="width:300px">
                                <option value=""></option>
                                <%			
				sql = "SELECT fr_id, fr_code, fr_description, fr_type FROM tblfaultyreason where fr_id is not null and fr_status='Y'" 
				
				if job_tech_type <> "" then 
				sql = sql & " and fr_type='" & job_tech_type & "'"	
				
				elseif job_type <> "" then 
				sql = sql & " and fr_type='" & job_type & "'"	 
				
				end if
				
				sql = sql & " order by fr_code"	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (job_tech_faulty_code) = (rs("fr_code")) then
					  response.write "<option value='" & rs("fr_code") & "' selected>" & rs("fr_code") & " - " & rs("fr_description")  & "</option>"
					  else
					  response.write "<option value='" & rs("fr_code") & "'>" & rs("fr_code") & " - " & rs("fr_description")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                              </select></td>
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Product Collection Date</strong></font></td>
                              <td align="left"><font color="#000000"><strong>
                                <input name="job_tech_product_collectdate" type="text" id="job_tech_product_collectdate" value="<%=chkdate(job_tech_product_collectdate)%>" size="12" maxlength="20" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.formorder2.job_tech_product_collectdate;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                            </tr>
                            <tr >
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Repair Action / <br />
Remark*<br />
                              </strong>(150 Chars) <strong><br />
                              <br />
                              <a href="javascript:popup('rm_job_new_repaitaction.asp?searchitem=ra_repairaction&amp;searchvalue=<%=cust_code%>&amp;md_type=' + formorder2.job_tech_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')"><font color="#FFFFFF">[Select]</font></a></strong></font></td>
                              <td align="left"><strong>
                                <textarea name="job_tech_faulty_action" cols="60" rows="6" wrap="virtual" id="job_tech_faulty_action"><%=job_tech_faulty_action%></textarea>
                              </strong></td>
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Actual Warranty Status*</strong></font></td>
                              <td align="left"><select name="job_actual_wrty_status" id="select15">
                                <option value="Under" <%if job_actual_wrty_status="Under" then response.write " selected"%>>Under</option>
                                <option value="Over" <%if job_actual_wrty_status="Over" then response.write " selected"%>>Over</option>
                              </select></td>
                            </tr>
                            <tr >
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Warranty Photo<br />
(Optional) </strong></font></td>
                              <td align="left"><input type="file" name="job_wrty_photo" id="job_wrty_photo" /><br/>
                                  <input type="file" name="job_wrty_photo2" id="job_wrty_photo2" />
                                  <input type="file" name="job_wrty_photo3" id="job_wrty_photo3" />
                                <a href="shared/<%=job_wrty_photo%>" target="_blank"><img src="shared/<%=job_wrty_photo%>" alt="Click on to Pop-up" width="100" border="0" /></a>
                                <a href="shared/<%=job_wrty_photo2%>" target="_blank"><img src="shared/<%=job_wrty_photo2%>" alt="Click on to Pop-up" width="100" border="0" /></a>
                                <a href="shared/<%=job_wrty_photo3%>" target="_blank"><img src="shared/<%=job_wrty_photo3%>" alt="Click on to Pop-up" width="100" border="0" /></a></td>
                              <td align="left"></td>
                              <td align="left"></td>
                              </tr>
                            </tbody>
                        </table></td>
                      </tr>
                    <tr align="right">
                      <td colspan="2" scope="col">
                     
                      <%if job_status="Accepted" then%>  
                      <input type="submit" name="button7" id="button10" value="Save" />
                      <input type="button" name="SubmitJob2" id="SubmitJob2" value="Update Job Status " onclick="javascript:confirmJobDone('<%=job_code%>','action.asp?type=DoneJob')" />
                      <%end if%>
                      </td>
                    </tr>
                    </form>
                    
                    
                    
                    <tr>
                      <td colspan="2" scope="col">&nbsp;</td>
                    </tr>
                    
    <%if Request.Cookies("GAPS")("slevel") = "cs" then %>                
                   <form name="formorder3" id="formorder3" method="post" action="action.asp?type=editjob_hq" >
                      </form>
                    
       <%end if%>
                    
                    <tr align="right">
                      <td colspan="2" bgcolor="#FFFFFF"></td>
                    </tr>
                    
                    <tr>
                      <td></tbody></td>
                    </tr>
                  </table></td>
                </tr>
                
<%end if %>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
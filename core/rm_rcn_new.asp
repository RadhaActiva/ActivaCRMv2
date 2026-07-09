<!-- #include file="header.asp" -->
<%

rcn_cust_cnty_id=request.querystring("rcn_cust_cnty_id")
rcn_cust_name=request.querystring("rcn_cust_name")
rcn_cust_tel1=request.querystring("rcn_cust_tel1")
rcn_cust_tel2=request.querystring("rcn_cust_tel2")
rcn_cust_address=request.querystring("rcn_cust_address")
rcn_cust_postcode = request.form("rcn_cust_postcode")

if rcn_cust_postcode = "" then
    rcn_cust_postcode = request.QueryString("rcn_cust_postcode")
end if

set rs = server.CreateObject("adodb.recordset")

if request("rcn_no") <> "" then	  
 sql="SELECT rcn_id, rcn_no, rcn_date, rcn_status, rcn_job_code, rcn_onlineWrtyNo, rcn_SN_no, rcn_onlinewrtyStatus, rcn_modelcode, rcn_modeltype, rcn_tech_code, rcn_cust_code, rcn_cust_name, rcn_cust_address, rcn_cust_postcode, " & _
			"rcn_cust_state, rcn_cust_state_id, rcn_cust_city, rcn_cust_city_id, rcn_cust_cnty_id, rcn_cust_email, rcn_cust_tel1, rcn_cust_tel2, rcn_remark,   " & _
			"rcn_createddate, rcn_createdby, rcn_submitteddate, rcn_submittedby,  " & _
			"rcn_posteddate, rcn_postedby, rcn_cancelleddate, rcn_cancelledby, rcn_totalqty, rcn_totalPartsAmt, rcn_labourAmt, rcn_transportAmt, rcn_gstAmt, rcn_gstRate,  " & _
			"rcn_gstCode, rcn_totalAmt, rcn_emailsentdate " & _
			"FROM tblrcn where rcn_no = '" & request("rcn_no") & "'"
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			rcn_id = rs("rcn_id") 
			rcn_no = rs("rcn_no") 
			rcn_date = rs("rcn_date") 
			rcn_status = rs("rcn_status") 
			rcn_job_code = rs("rcn_job_code") 
			rcn_onlineWrtyNo = rs("rcn_onlineWrtyNo") 
			rcn_SN_no = rs("rcn_SN_no") 
			rcn_onlinewrtyStatus = rs("rcn_onlinewrtyStatus") 
			rcn_modelcode = rs("rcn_modelcode") 
			rcn_modeltype = rs("rcn_modeltype") 
			rcn_tech_code = rs("rcn_tech_code") 
			rcn_cust_code = rs("rcn_cust_code")
			rcn_cust_name = rs("rcn_cust_name")
			rcn_cust_address = rs("rcn_cust_address") 
			rcn_cust_postcode = rs("rcn_cust_postcode") 
			rcn_cust_state = rs("rcn_cust_state") 
			rcn_cust_state_id = rs("rcn_cust_state_id") 
			rcn_cust_city = rs("rcn_cust_city") 
			rcn_cust_city_id = rs("rcn_cust_city_id") 
            rcn_cust_cnty_id = rs("rcn_cust_cnty_id")
			rcn_cust_email = rs("rcn_cust_email") 
			rcn_cust_tel1 = rs("rcn_cust_tel1") 
			rcn_cust_tel2 = rs("rcn_cust_tel2")  
			rcn_remark = rs("rcn_remark")
			rcn_createddate = rs("rcn_createddate") 
			rcn_createdby = rs("rcn_createdby") 
			rcn_submitteddate = rs("rcn_submitteddate") 
			rcn_submittedby = rs("rcn_submittedby") 
			rcn_posteddate = rs("rcn_posteddate") 
			rcn_postedby = rs("rcn_postedby")
			rcn_cancelleddate = rs("rcn_cancelleddate") 
			rcn_cancelledby = rs("rcn_cancelledby")
			rcn_totalqty = rs("rcn_totalqty")
			rcn_totalPartsAmt = rs("rcn_totalPartsAmt")
			rcn_labourAmt = rs("rcn_labourAmt")
			rcn_transportAmt = rs("rcn_transportAmt")
			rcn_gstAmt = rs("rcn_gstAmt")
			rcn_gstRate = rs("rcn_gstRate")
			rcn_gstCode = rs("rcn_gstCode")
			rcn_totalAmt = rs("rcn_totalAmt")
			rcn_emailsentdate = rs("rcn_emailsentdate")
		End If
		rs.Close
	  stype = "editRCN"	
	  actionname = "Save" 
 else    
	  stype = "addRCN"
	  actionname = "Save" 		
	  rcn_date = date()	 
	  rcn_status = "Open"   	
end if


if rcn_job_code <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_cnty_id, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_rcn_no, job_rcn_date, job_do_no, job_do_date, job_posteddate " & _
	    "FROM tbljob WHERE job_code = '" & rcn_job_code & "' "
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
			job_faulty_reason_cs = rs("job_faulty_reason_cs") 
			job_faulty_desc = rs("job_faulty_desc") 
			job_reportedby = rs("job_reportedby") 
			job_appointment_date = rs("job_appointment_date") 
			job_appointment_time = rs("job_appointment_time") 
			job_tech_code = rs("job_tech_code") 
			job_appointment_remark = rs("job_appointment_remark") 
			job_emailsentdate = rs("job_emailsentdate") 
			job_emailsent = rs("job_emailsent") 
			job_smssentdate = rs("job_smssentdate") 
			job_smssent = rs("job_smssent") 
			job_tech_type = rs("job_tech_type") 
			job_tech_model = rs("job_tech_model") 
			job_tech_tax_invoice = rs("job_tech_tax_invoice") 
			job_tech_SN = rs("job_tech_SN") 
			job_tech_faulty_reason = rs("job_tech_faulty_reason") 
			job_tech_faulty_action = rs("job_tech_faulty_action") 
			job_tech_status = rs("job_tech_status") 
			job_tech_product_collectdate = rs("job_tech_product_collectdate") 
			job_tech_returntoCustDate = rs("job_tech_returntoCustDate") 
			job_actual_wrty_status = rs("job_actual_wrty_status") 
			job_wrty_photo = rs("job_wrty_photo") 
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
			job_rcn_no = rs("job_rcn_no") 
			job_rcn_date = rs("job_rcn_date") 
			job_do_no = rs("job_do_no")
			job_do_date = rs("job_do_date") 
			job_posteddate = rs("job_posteddate")
		End If
		rs.Close
end if
 
if rcn_tech_code <> "" then	  
sql = "SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id " & _
	  "FROM tbltechnician WHERE tech_code = '" & rcn_tech_code & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			tech_name = rs("tech_name") 
			tech_tel1 = rs("tech_tel1") 
		End If
		rs.Close
end if

if rcn_cust_postcode <> "" and rcn_cust_cnty_id = "129" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & rcn_cust_postcode & "'"
		rs1.Open sql1,strconnect,0,1,&H0001
		If Not rs1.EOF Then
             rcn_cust_state_id = rs1("state_id") 'will auto populate state
             rcn_cust_state = rs1("state_name")
             rcn_cust_city_id = rs1("city_id") 'will auto populate city
             rcn_cust_city = rs1("post_office")
        end if
    rs1.close
end if

%>

<script language="javascript">

    function getPostcode(p) {
        //alert(s);
        //document.getElementById('cust_name').value = s;
        document.getElementById('rcn_cust_postcode').value = p;
        document.formorder.submit();
    }

    function getCountrycode(c) {
        document.getElementById('rcn_cust_cnty_id').value = c;
        document.formorder.submit();
    }

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formrcndetail.action = orderlinks;
	document.formrcndetail.submit();
   }
}


function calctotal(unitprice, qty, discountamt, discounttype, subtotal)
{
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
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Return Credit Note (RCN)</div></td>
                        <td width="23%" align="right" class="titleblue1"><a href="rm_rcn_new_print.asp?rcn_no=<%=rcn_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Code *</strong></font></td>
                        <td align="left"><label for="rcn_cust_code"></label>
                          <input name="rcn_cust_code" type="text" id="rcn_cust_code" style="background-color: #cccccc;" value="<%=rcn_cust_code%>" maxlength="50" onfocus="this.blur();" />
                          [<a href="javascript:popup('rm_rcn_new_customer.asp?searchitem=tblcustomer.cust_code&amp;searchvalue=<%=cust_code%>&amp;formname=&amp;fields=rcn_cust_code','cb17','scrollbars=yes,resizable=yes,width=900,height=650')")">Select Customer</a>] </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
                        <td align="left"><input name="rcn_cust_name" type="text" id="rcn_cust_name" value="<%=rcn_cust_name%>" size="50" maxlength="100" /></td>
                      </tr>
                         <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left"><select name="rcn_cust_cnty_id" id="rcn_cust_cnty_id" style="width:150px"  onblur="getCountrycode(this.value)">
                                    <option value="<%=rcn_cust_cnty_id%>"></option>                                       
                                    <%
                                          sql = "SELECT cnty_name,cnty_id from tblcountry"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((rcn_cust_cnty_id)) = cstr((rs1("cnty_id"))) then
					                               response.write "<option value='" & rs1("cnty_id") & "' selected>" & rs1("cnty_name") & "</option>"
					                            else
					                               response.write "<option value='" & rs1("cnty_id") & "'>" & rs1("cnty_name") & "</option>"
					                            end if 	                                
                                          rs1.movenext 
                                          wend                                                                       
                                          rs1.close 
                                    %>
                                        </select></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><strong>
                          <textarea name="rcn_cust_address" cols="50" rows="3" id="rcn_cust_address"><%=rcn_cust_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left"><strong>
                          <input name="rcn_cust_postcode" type="text" id="rcn_cust_postcode" value="<%=rcn_cust_postcode%>" onblur="getPostcode(this.value)" size="20" maxlength="20" />
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left">
                             <input name="rcn_cust_state" type="text" id="rcn_cust_state" value="<%=rcn_cust_state%>" size="30" readonly maxlength="50" />
                             <input name="rcn_cust_state_id" type="hidden" id="rcn_cust_state_id" value="<%=rcn_cust_state_id%>" size="30" maxlength="50" /> 
                         </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                             <input name="rcn_cust_city" type="text" id="rcn_cust_city" value="<%=rcn_cust_city%>" size="30" readonly maxlength="50" />
                             <input name="rcn_cust_city_id" type="hidden" id="rcn_cust_city_id" value="<%=rcn_cust_city_id%>" size="30" maxlength="50" /> 
                         </td>
                      </tr>                      
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="rcn_cust_email" type="text" id="rcn_cust_email" value="<%=rcn_cust_email%>" size="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><label for="rcn_cust_tel1"></label>
                          <input name="rcn_cust_tel1" type="text" id="rcn_cust_tel1" value="<%=rcn_cust_tel1%>" size="30" maxlength="50" />
                          e.g 0121234657</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="rcn_cust_tel2" type="text" id="rcn_cust_tel2" value="<%=rcn_cust_tel2%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td valign="top"><strong>
                          <textarea name="rcn_remark" cols="50" rows="3" id="rcn_remark"><%=rcn_remark%></textarea>
                        </strong></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> RCN Information</font></strong></td>
                      </tr>
                      <tr>
                        <td nowrap="nowrap"><font color="#FFFFFF"><strong>RCN  
                          No.<br />
                          <font size="1">(System Generate) </font></strong></font></td>
                        <td align="left"><strong><%=rcn_no%></strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>RCN  
                          Date</strong></font></td>
                        <td align="left"><%=chkdate(rcn_date)%></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job No.</strong></font></td>
                        <td align="left"><a href="rm_jobsheet.asp?job_code=<%=rcn_job_code%>" target="_blank"><strong><%=rcn_job_code%></strong></a></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Posted Date</strong></font></td>
                        <td><%=chkdate(job_posteddate)%></td>
                      </tr>
                      <tr align="left" >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status</strong></font></td>
                        <td><strong><%=rcn_status%></strong></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>SN No.</strong></font></td>
                        <td><%=rcn_SN_no%></td>
                      </tr>
                      <tr align="left" >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Wrty Card No.</strong></font></td>
                        <td>
                        <strong><%=rcn_onlineWrtyNo%></strong></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                        <td><strong><%=rcn_onlinewrtyStatus%></strong></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model</strong></font></td>
                        <td align="left"><strong><%=rcn_modelcode%></strong></td>
                        <!--<td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type</strong></font></td>-->
						<td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>RCN Type</strong></font></td>
                        <td align="left"><label for="select2"></label>
                        <!--<strong><%=rcn_modeltype%></strong>-->
						<select name="rcn_modeltype" id="rcn_modeltype">
                            <option value="Replacement" <%if rcn_modeltype="Replacement" then response.write " selected"%>>Replacement</option>
                            <option value="CN" <%if rcn_modeltype="CN" then response.write " selected"%>>CN</option>
                          </select>
						</td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                        <td colspan="3"><strong><%=rcn_tech_code%> - <%=tech_name%></strong></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                        <td colspan="3"><%=rcn_createdby%> @
                          <%=chkdatetime(rcn_createddate)%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Submitted By</strong></font></td>
                        <td colspan="3"><%=rcn_submittedby%> @
                          <%=chkdatetime(rcn_submitteddate)%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Posted By</strong></font></td>
                        <td colspan="3"><%=rcn_postedby%>@ <%=chkdatetime(rcn_posteddate)%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cancelled  By</strong></font></td>
                        <td colspan="3"><%=rcn_cancelledby%>@
                          <%=chkdatetime(rcn_cancelleddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF">
				  <input type="hidden" name="rcn_no" id="rcn_no" value="<%=rcn_no%>" />
				  <%if rcn_status="Open" then %>
                    <input type="submit" name="button" id="button" value="<%=actionname%>" />
                  <%end if%></td>
              </tr>
              </form>
              
              
              <%if rcn_no <> "" then %>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <form name="form1" id="form1" method="post" action="global_ma_repeating_new.asp?type=editRepeat">
                    </form>
                    <tr valign="top">
                      <td colspan="3" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Model Code<a name="spareparts" id="spareparts"></a></strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Model Description</strong></font></td>
                          <!--Added By sanjay on 23/Feb/2012-->
                          <td align="right"><font color="#FFFFFF"><strong>Unit Price (RM)</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Discount </strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Total 
                            Amt (RCP)</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Action</strong></font></td>
                        </tr>
                        <%

if request("rcnd_id") <> "" then
		sql = "SELECT rcnd_id, rcnd_rcn_no, rcnd_job_code, rcnd_partcode, rcnd_desc, rcnd_unitcost, rcnd_qty, rcnd_discountamt, " & _
		      "rcnd_discounttype, rcnd_netcost, rcnd_subtotal FROM tblrcn_detail where rcnd_id = '" & request("rcnd_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   rcnd_id = rs("rcnd_id")
		   rcnd_rcn_no = rs("rcnd_rcn_no")
		   rcnd_job_code = rs("rcnd_job_code")
		   rcnd_partcode = rs("rcnd_partcode")
		   rcnd_desc = rs("rcnd_desc")
		   rcnd_unitcost = rs("rcnd_unitcost")
		   rcnd_qty = rs("rcnd_qty")
		   rcnd_discountamt = rs("rcnd_discountamt")
		   rcnd_discounttype = rs("rcnd_discounttype")  
		   rcnd_netcost = rs("rcnd_netcost")   
		   rcnd_subtotal = rs("rcnd_subtotal")
        end if
		rs.close
		sbutton = "Update"
		stype="editRCNDetail"	
else
		sbutton = "Add"
		stype="addRNCDetail"
		rcnd_qty = "1"	
		rcnd_unitcost = "0.00"	
		rcnd_discountamt = "0.00"
		rcnd_netcost = "0.00"	
		rcnd_netcost = "0.00"	
end if

%>
                        <%if rcn_status="Open" then %>
                        <form name="formrcndetail" id="formrcndetail" method="post" action="rm_rcn_new.asp#spareparts" >
                          <tr>
                            <td bgcolor="#666666">&nbsp;</td>
                            <td align="left" nowrap="nowrap" bgcolor="#666666"><input name="rcnd_partcode" type="text" id="rcnd_partcode" value="<%=rcnd_partcode%>" size="20" maxlength="50" />
                              [<a href="javascript:popup('rm_rcn_new_model.asp?searchitem=md_code&amp;searchvalue=<%=cust_code%>&amp;formname=formrcndetail&amp;fieldname=rcnd_partcode','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                            <td align="left" bgcolor="#666666"><font color="#FFFFFF"> </font>
                              <input name="rcnd_desc" type="text" id="rcnd_desc" value="<%=rcnd_desc%>" size="30" maxlength="100" /></td>
                            <td align="right" bgcolor="#666666"><font color="#FFFFFF">
                              <input type="hidden" name="rcn_no" id="rcn_no" value="<%=rcn_no%>" />
                              <input type="hidden" name="rcnd_job_code" id="rcnd_job_code" value="<%=rcn_job_code%>" />
                              <input type="hidden" name="rcnd_id" id="rcnd_id" value="<%=rcnd_id%>" />
                              <input name="rcnd_unitcost" type="text" id="rcnd_unitcost" style="text-align:right; background-color: #cccccc;" onfocus="this.blur();" onkeydown="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);" onkeyup="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);" value="<%=rcnd_unitcost%>" size="5" maxlength="10" />
                            </font></td>
                            <td align="right" bgcolor="#666666"><input name="rcnd_qty" type="text" id="rcnd_qty" style="text-align:right" onkeydown="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);" onkeyup="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);" value="<%=rcnd_qty%>" size="5" maxlength="5" /></td>
                            <td align="right" nowrap="nowrap" bgcolor="#666666"><font color="#FFFFFF">
                              <input name="rcnd_discountamt" type="text" id="rcnd_discountamt" value="<%=rcnd_discountamt%>" size="5" onkeydown="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);" onkeyup="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);" style="text-align:right" />
                              <select name="rcnd_discounttype" id="rcnd_discounttype" onchange="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);" onkeyup="calctotal(document.formrcndetail.rcnd_unitcost.value, document.formrcndetail.rcnd_qty.value, document.formrcndetail.rcnd_discountamt.value, document.formrcndetail.rcnd_discounttype.value, document.formrcndetail.rcnd_subtotal);">
                                <option value="%" <%if rcnd_discounttype = "%" then response.write " selected"%>>%</option>
                                <option value="RM" <%if rcnd_discounttype = "RM" then response.write " selected"%>>RM</option>
                              </select>
                            </font></td>
                            <td align="right" bgcolor="#666666"><input name="rcnd_subtotal" type="text" id="rcnd_subtotal" style="text-align:right; background-color: #cccccc;" onfocus="this.blur();" value="<%=rcnd_subtotal%>" size="10" maxlength="10" /></td>
                            <td align="center" bgcolor="#666666"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("rcnd_id")%>','action.asp?type=<%=stype%>','<%=rcnd_subtotal%>');" /></td>
                          </tr>
                        </form>
                        <%end if%>
                        <%				i = 1
				sql1 = "SELECT rcnd_id, rcnd_rcn_no, rcnd_job_code, rcnd_partcode, rcnd_desc, rcnd_unitcost, rcnd_qty, rcnd_discountamt, rcnd_discounttype, rcnd_netcost, rcnd_subtotal " & _
	                   "FROM tblrcn_detail where rcnd_rcn_no = '" & rcn_no & "' order by rcnd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="left"><%=rs1("rcnd_partcode")%></td>
                          <td align="left"><%=rs1("rcnd_desc")%></td>
                          <td align="right"><%=chknumber2(rs1("rcnd_unitcost"))%></td>
                          <td align="right"><%=rs1("rcnd_qty")%></td>
                          <td align="right">- <%=chknumber2(rs1("rcnd_unitcost")-rs1("rcnd_netcost"))%></td>
                          <td align="right"><%=chknumber2(rs1("rcnd_subtotal"))%></td>
                          <td align="center" nowrap="nowrap"><%if rcn_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_rcn_new.asp?rcnd_id=<%=rs1("rcnd_id")%>&rcn_no=<%=rcn_no%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("rcnd_partcode")%>','action.asp?type=delRCNDetail&amp;rcnd_id=<%=rs1("rcnd_id")%>&amp;rcn_no=<%=rcn_no%>')" />
                            <%end if%></td>
                        </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="6" align="right"><strong>Total</strong></td>
                          <td align="right"><%=chknumber2(rcn_totalAmt)%></td>
                          <td>&nbsp;</td>
                        </tr>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="8" align="left">**GST 6% Inclusive, 
                            GST Amount: RM <%=chknumber2(rcn_gstAmt)%></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                     <!-- <td width="55%" align="left" bgcolor="#FFFFFF" 
          scope="col">
Email
  <input name="rcn_emailsentdate" type="text" id="rcn_emailsentdate" size="50" value="<%=rcn_emailsentdate%>" />
  <input type="submit" name="button4" id="button4" value="Email RCN" /></td>-->
                      <td width="22%" align="right" bgcolor="#FFFFFF" 
          scope="col"><%if rcn_status<>"Cancel" or rcn_status<>"Posted" then %>
                        <input type="button" name="PostedInvoice" id="PostedInvoice" value="Cancel RCN " onclick="javascript:confirmAction('<%=rcn_no%>','action.asp?type=CancelRCN&amp;rcn_no=<%=rcn_no%>')" />
                    <%end if%>
                      </td>
                      <td width="23%" align="right" bgcolor="#FFFFFF" 
          scope="col"><%if rcn_status="Open" then %>
                        <input type="button" name="SubmitJob2" id="SubmitJob2" value="Submit RCN " onclick="javascript:confirmAction('<%=rcn_no%>','action.asp?type=submitRCN&amp;rcn_no=<%=rcn_no%>')" />
                        <%end if%>
                        <br />
                        <%if rcn_status="Submitted" then %>
                        <input type="button" name="PostedInvoice2" id="PostedInvoice2" value="Posted RCN " onclick="javascript:confirmAction('<%=rcn_no%>','action.asp?type=PostedRCN&amp;rcn_no=<%=rcn_no%>')" />
                      <%end if%></td>
                    </tr>
                    <tr align="right">
                      <td colspan="3" bgcolor="#FFFFFF"></td>
                    </tr>
                    <tr>
                      <td></tbody></td>
                    </tr>
                  </table></td>
                </tr>
                <%end if%>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
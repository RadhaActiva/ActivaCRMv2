<!-- #include file="header.asp" -->
<%
do_cust_name=request.querystring("do_cust_name")
do_cust_tel1=request.querystring("do_cust_tel1")
do_cust_tel2=request.querystring("do_cust_tel2")
do_cust_address=request.querystring("do_cust_address")
do_cust_postcode = request.form("do_cust_postcode")

if do_cust_postcode = "" then
    do_cust_postcode = request.QueryString("do_cust_postcode")
end if

set rs = server.CreateObject("adodb.recordset")

if request("do_no") <> "" then	  
sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, do_cust_state, " & _
		"do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_cnty_id,do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby, do_job_code, do_tech_code,  " & _
		"do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, do_emailsentdate, do_deliveredby,  " & _
		"do_delivereddate, do_submittedby, do_submitteddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate,  " & _
		"do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time, do_appointment_remark " & _
		"FROM tbldo WHERE do_no = '" & request("do_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			do_id = rs("do_id") 
			do_no = rs("do_no") 
			do_status = rs("do_status")
			do_date = rs("do_date") 
			do_inv_no = rs("do_inv_no")  
			do_inv_date = rs("do_inv_date")  
			do_cust_code = rs("do_cust_code") 
			do_cust_name = rs("do_cust_name")  
			do_cust_address = rs("do_cust_address")  
			do_cust_postcode = rs("do_cust_postcode")  
			do_cust_state = rs("do_cust_state")  
			do_cust_state_id = rs("do_cust_state_id")  
			do_cust_city = rs("do_cust_city")  
			do_cust_city_id = rs("do_cust_city_id") 
            do_cust_cnty_id = rs("do_cust_cnty_id") 
			do_cust_email = rs("do_cust_email")  
			do_cust_tel1 = rs("do_cust_tel1")  
			do_cust_tel2 = rs("do_cust_tel2")  
			do_createddate = rs("do_createddate")  
			do_createdby = rs("do_createdby")  
			do_job_code = rs("do_job_code")  
			do_tech_code = rs("do_tech_code")  
			do_totalqty = rs("do_totalqty")  
			do_totalPartsAmt = rs("do_totalPartsAmt")  
			do_remark = rs("do_remark")  
			do_labourAmt = rs("do_labourAmt")  
			do_transportAmt = rs("do_transportAmt")  
			do_gstAmt = rs("do_gstAmt")  
			do_totalAmt = rs("do_totalAmt")  
			do_emailsent = rs("do_emailsent")  
			do_emailsentdate = rs("do_emailsentdate")  
			do_deliveredby = rs("do_deliveredby")  
			do_delivereddate = rs("do_delivereddate")  
			do_submittedby = rs("do_submittedby")  
			do_submitteddate = rs("do_submitteddate")  
			do_doneby = rs("do_doneby")  
			do_donedate = rs("do_donedate")  
			do_postedby = rs("do_postedby")  
			do_posteddate = rs("do_posteddate")  
			do_cancelledby = rs("do_cancelledby")  
			do_cancelleddate = rs("do_cancelleddate")  
			do_purchase_date = rs("do_purchase_date")  
			do_onlineWrtyNo = rs("do_onlineWrtyNo")  
			do_onlineWrtyStatus = rs("do_onlineWrtyStatus")  
			do_SN_no = rs("do_SN_no")  
			do_type = rs("do_type")  
			do_Model = rs("do_Model")  
			do_model_desc = rs("do_model_desc")  
			do_appointment_date = rs("do_appointment_date")  
			do_appointment_time = rs("do_appointment_time")  
			do_appointment_remark = rs("do_appointment_remark")
		End If
		rs.Close
	  stype = "editDO"	
	  actionname = "Save" 
 else    
	  stype = "addDO"
	  actionname = "Save" 		
	  do_date = date()	 
	  do_status = "Open"   	
	  do_tech_code = "walk-in"
end if


if do_job_code <> "" and do_cust_cnty_id="129" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_cnty_id, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date " & _
	    "FROM tbljob WHERE job_code = '" & do_job_code & "' "
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
            job_cust_cnty_id= rs("do_cust_cnty_id")
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
			job_inv_no = rs("job_inv_no") 
			job_inv_date = rs("job_inv_date") 
			job_do_no = rs("job_do_no")
			job_do_date = rs("job_do_date") 
		End If
		rs.Close
end if

if do_cust_postcode <> "" and do_cust_cnty_id="129" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & inv_cust_postcode & "' "
		rs1.Open sql1,strconnect,0,1,&H0001   
		If Not rs1.EOF Then
             do_cust_state_id = rs1("state_id") 'will auto populate state
             do_cust_state=  rs1("state_name")
             do_cust_city_id = rs1("city_id") 'will auto populate city
             do_cust_city = rs1("post_office")    
        end if
    rs1.close
end if
%>

<script language="javascript">

 function getPostcode(p)
    {    
        document.getElementById('do_cust_postcode').value = p;
        document.formorder.submit();
    }

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formdodetail.action = orderlinks;
	document.formdodetail.submit();
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
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Update </font>DO</div></td>
                        <td width="23%" align="right" class="titleblue1">
                        <%if do_status="Posted" then %>
                        <a href="rm_do_new_print.asp?do_no=<%=do_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print A4 format" border="0" style="border:0"/></a>&nbsp;&nbsp; <a href="/CRMone/pdfside/default.aspx?do_no=<%=do_no%>" target="_blank"><img src="images/pdf.png" height="35" width="30" alt="Pre-Printed Dot Matrix" border="0" style="border:0"/></a> 
                        <%end if%>
                        </td>
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
                        <td align="left"><label for="inv_cust_code"></label>
                          <input name="do_cust_code" type="text" id="do_cust_code" style="background-color: #cccccc;" value="<%=do_cust_code%>" maxlength="50" />
                          [<a href="javascript:popup('rm_do_new_customer.asp?searchitem=tblcustomer.cust_code&amp;searchvalue=<%=cust_code%>&amp;formname=&amp;fields=do_cust_code','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select Customer</a>] </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
                        <td align="left"><input name="do_cust_name" type="text" id="do_cust_name" value="<%=do_cust_name%>" size="50" maxlength="100" /></td>
                      </tr>
                         <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left">
                                    <select name="do_cust_cnty_id" id="do_cust_cnty_id" style="width:150px">
                                    <option value="<%=do_cust_cnty_id%>"></option>                                       
                                    <%
                                          sql = "SELECT cnty_name,cnty_id from tblcountry"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((do_cust_cnty_id)) = cstr((rs1("cnty_id"))) then
					                               response.write "<option value='" & rs1("cnty_id") & "' selected>" & rs1("cnty_name") & "</option>"
					                            else
					                               response.write "<option value='" & rs1("cnty_id") & "'>" & rs1("cnty_name") & "</option>"
					                            end if 	                                
                                          rs1.movenext 
                                          wend                                                                       
                                          rs1.close 
                                    %>
                                    </select>       
                        </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><strong>
                          <textarea name="do_cust_address" cols="50" rows="3" id="do_cust_address"><%=do_cust_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left"><strong>
                          <input name="do_cust_postcode" type="text" id="do_cust_postcode" value="<%=do_cust_postcode%>" onblur="getPostcode(this.value)" size="20" maxlength="20" />
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left">
                             <input name="do_cust_state" type="text" id="do_cust_state" value="<%=do_cust_state%>" size="30" readonly maxlength="50" />
                             <input name="do_cust_state_id" type="hidden" id="do_cust_state_id" value="<%=do_cust_state_id%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                            <input name="do_cust_city" type="text" id="do_cust_city" value="<%=do_cust_city%>" size="30" readonly maxlength="50" />
                            <input name="do_cust_city_id" type="hidden" id="ddo_cust_city_id" value="<%=do_cust_city_id%>" size="30" maxlength="50" />
                          </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="do_cust_email" type="text" id="do_cust_email" value="<%=do_cust_email%>" size="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><label for="do_cust_tel1"></label>
                          <input name="do_cust_tel1" type="text" id="do_cust_tel1" value="<%=do_cust_tel1%>" size="30" maxlength="50" />
                          e.g 0121234657</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="do_cust_tel2" type="text" id="do_cust_tel2" value="<%=do_cust_tel2%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td valign="top"><strong>
                          <textarea name="do_remark" cols="50" rows="3" id="do_remark"><%=do_remark%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Created by</strong></font></td>
                        <td valign="top"><%=do_createdby%>@
                        <%=chkdatetime(do_createddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Delivered by</strong></font></td>
                        <td valign="top"><%=do_deliveredby%> @ <%=chkdatetime(do_delivereddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Done by</strong></font></td>
                        <td valign="top"><%=do_doneby%> @ <%=chkdatetime(do_donedate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Posted by</strong></font></td>
                        <td valign="top"><%=do_postedby%> @ <%=chkdatetime(do_posteddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cancelled  by</strong></font></td>
                        <td valign="top"><%=do_cancelledby%> @ <%=chkdatetime(do_cancelleddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> DO Information</font></strong></td>
                      </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>DO  
                          No.<br />
                          <font size="1">(System Generate) </font></strong></font></td>
                        <td align="left"><strong><%=do_no%></strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>DO  
                          Date</strong></font></td>
                        <td align="left"><%=chkdate(do_date)%></td>
                      </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status</strong></font></td>
                        <td align="left"><strong><%=do_status%></strong></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Purchase  Date</strong></font></td>
                        <td align="left"><font color="#000000"><strong>
                          <input name="do_purchase_date" type="text" id="do_purchase_date" value="<%=chkdate(do_purchase_date)%>" size="12" maxlength="20" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.formorder.do_purchase_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                      </tr>
                      <tr align="left" >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job Code</strong></font></td>
                        <td align="left"><strong><%=do_job_code%></strong></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job Date</strong></font></td>
                        <td align="left"><%=chkdate(job_date)%></td>
                      </tr>
                      <tr align="left" >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
                        <td align="left"><label for="textfield5"></label>
                        <input name="do_onlineWrtyNo" type="text" id="do_onlineWrtyNo" value="<%=do_onlineWrtyNo%>" maxlength="50" /></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                        <td align="left"><strong>
                          <%if do_onlineWrtyStatus="Over" then %>
                          <font color="#FF0000">Over</font>
                          <%else%>
                          <font color="#000000">Under</font>
                          <%end if%>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>S/N</strong></font></td>
                        <td colspan="3" align="left"><input name="do_SN_no" type="text" id="do_SN_no" value='<%=do_SN_no%>' maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type</strong></font></td>
                        <td colspan="3" align="left"><select name="do_type" id="do_type">
                          <option value=""></option>
                          <option value="CF" <%if do_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                          <option value="WH" <%if do_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                        </select>
                          <label for="select3"></label></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model</strong></font></td>
                        <td colspan="3" align="left"><input name="do_Model" type="text" id="do_Model" value="<%=do_Model%>" size="20" maxlength="50" />
                          <%if Request.Cookies("GAPS")("slevel") = "cs" or Request.Cookies("GAPS")("slevel") = "sc" then %>
                          <label for="do_type">[<a href="javascript:popup('rm_do_new_model.asp?searchitem=md_code&amp;searchvalue=<%=cust_code%>&amp;formname=formorder&amp;fieldname=do_Model&amp;fieldname1=do_Model_desc&amp;md_type=<%=do_type%>','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </label>
                        <%end if%></td>
                      </tr>
                      <tr >
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model Description</strong></font></td>
                        <td colspan="3" align="left"><input name="do_Model_desc" type="text" id="do_Model_desc" value="<%=do_Model_desc%>" size="60" maxlength="100" /></td>
                      </tr>
                      <tr >
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Appointment Date</strong></font></td>
                        <td align="left"><font color="#000000"><strong>
                          <input name="do_appointment_date" type="text" id="do_appointment_date" value="<%=chkdate(do_appointment_date)%>" size="12" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.formorder.do_appointment_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Time</strong></font></td>
                        <td align="left"><label for="textfield">
                          <input type="text" name="textfield" id="textfield" />
                        </label></td>
                      </tr>
                      <tr >
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                        <td colspan="3" align="left"><select name="do_tech_code" id="do_tech_code">
                          <option value=""></option>
                          <option value="resolved_no_appt">Issue resolve without appt</option>
                          <%			
				sql = "SELECT tech_id, tech_code, tech_name FROM tbltechnician where tech_status = 'Y' "	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (do_tech_code) = (rs("tech_code")) then
					  response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                        </select>
                          <font color="#000000"><strong>
                          <%if Request.Cookies("GAPS")("slevel") = "cs" or Request.Cookies("GAPS")("slevel") = "sc" then %>
                          <a href="javascript:popup('rm_do_new_schedule.asp?searchitem=do_date&amp;do_date=<%=chkdate(do_date)%>&amp;do_cust_state=' + formorder.do_cust_state.options[formorder.do_cust_state.selectedIndex].text,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Schedule<img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>
                          <%end if%>
                        </strong></font></td>
                      </tr>
                      <tr >
                        <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Invoice No.</font></strong></td>
                        <td align="left"><label for="select13"><strong>
                          <input name="do_inv_no" type="text" id="do_inv_no" value='<%=do_inv_no%>' maxlength="50" />
                         </strong></label></td>
                        <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Invoice Date</font></strong></td>
                        <td align="left"><font color="#000000"><strong>
                          <input name="do_inv_date" type="text" id="do_inv_date" value="<%=chkdate(do_inv_date)%>" size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.formorder.do_inv_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF"><label for="do_no"></label>
                    <input type="hidden" name="do_no" id="do_no" value="<%=do_no%>" />
                    <%if do_status<>"Cancel" then %>
                    <input type="submit" name="button" id="button" value="<%=actionname%>" />
                  <%end if%></td>
              </tr>
              </form>
              
          <%if do_no <> "" then %>    
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><a name="spareparts" id="spareparts"></a></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Item Code</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Model Description</strong></font></td>
                          <!--Added By sanjay on 23/Feb/2012-->
                          <td align="right"><font color="#FFFFFF"><strong>Unit Price (RM)</strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Action</strong></font></td>
                        </tr>
                        <%

if request("dod_id") <> "" then
		sql = "SELECT dod_id, dod_do_no, dod_inv_no, dod_job_code, dod_partcode, dod_desc, dod_unitcost, dod_qty, dod_discountamt, " & _
		      "dod_discounttype, dod_netcost, dod_subtotal " & _
	          "FROM tbldo_detail where dod_id = '" & request("dod_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   dod_id = rs("dod_id")
		   dod_do_no = rs("dod_do_no")
		   dod_inv_no = rs("dod_inv_no")
		   dod_job_code = rs("dod_job_code")
		   dod_partcode = rs("dod_partcode")
		   dod_desc = rs("dod_desc")
		   dod_unitcost = rs("dod_unitcost")
		   dod_qty = rs("dod_qty")
		   dod_discountamt = rs("dod_discountamt")
		   dod_discounttype = rs("dod_discounttype")  
		   dod_netcost = rs("dod_netcost")   
		   dod_subtotal = rs("dod_subtotal")
        end if
		rs.close
		sbutton = "Update"
		stype="editDODetail"	
else
		sbutton = "Add"
		stype="addDODetail"
		invd_qty = "1"	
		dod_unitcost = "0.00"	
		dod_discountamt = "0.00"
		dod_netcost = "0.00"	
		dod_subtotal = "0.00"	
end if

%>
                        <%if do_status="Open" then %>
                        <form name="formdodetail" id="formdodetail" method="post" action="rm_jobsheet.asp#spareparts" >
                          <tr>
                            <td bgcolor="#666666">&nbsp;</td>
                            <td align="left" bgcolor="#666666"><input name="dod_partcode" type="text" id="dod_partcode" value="<%=dod_partcode%>" maxlength="50" />
                              [<a href="javascript:popup('rm_do_new_model.asp?searchitem=md_code&amp;searchvalue=<%=cust_code%>&amp;formname=formdodetail&amp;fieldname=dod_partcode&fieldname1=dod_desc&md_type=<%=do_type%>','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                            <td align="left" bgcolor="#666666"><textarea name="dod_desc" cols="40" rows="3" id="dod_desc"><%=dod_desc%></textarea></td>
                            <td align="right" bgcolor="#666666"><font color="#FFFFFF">
                              <input type="hidden" name="do_no" id="do_no" value="<%=do_no%>" />
                              <input type="hidden" name="dod_id" id="dod_id" value="<%=dod_id%>" />
                              <input name="dod_unitcost" type="text" id="dod_unitcost" style="text-align:right; background-color: #cccccc;"  value="<%=dod_unitcost%>" size="5" maxlength="10" />
                            </font></td>
                            <td align="right" bgcolor="#666666"><input name="dod_qty" type="text" id="dod_qty" style="text-align:right"  value="<%=dod_qty%>" size="5" maxlength="5" /></td>
                            <td align="center" bgcolor="#666666"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("dod_id")%>','action.asp?type=<%=stype%>','<%=dod_subtotal%>');" /></td>
                          </tr>
                        </form>
                        <%end if%>
                        <%				i = 1
				sql1 = "SELECT dod_id, dod_do_no, dod_inv_no, dod_job_code, dod_partcode, dod_desc, dod_unitcost, dod_qty, dod_discountamt, " & _
						"dod_discounttype, dod_netcost, dod_subtotal " & _
						"FROM tbldo_detail where dod_do_no = '" & do_no & "' order by dod_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="left"><%=rs1("dod_partcode")%></td>
                          <td align="left"><%=rs1("dod_desc")%></td>
                          <td align="right"><%=chknumber2(rs1("dod_unitcost"))%></td>
                          <td align="right"><%=rs1("dod_qty")%></td>
                          <td align="center"><%if do_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_do_new.asp?dod_id=<%=rs1("dod_id")%>&do_no=<%=rs1("dod_do_no")%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("dod_partcode")%>','action.asp?type=delDODetail&amp;dod_id=<%=rs1("dod_id")%>&amp;do_no=<%=rs1("dod_do_no")%>')" />
                            <%end if%></td>
                        </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="4" align="right"><strong>Total</strong></td>
                          <td align="right"><%=chknumber0(do_totalqty)%></td>
                          <td>&nbsp;</td>
                        </tr>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="6" align="left">&nbsp;</td>
                        </tr>
                      </table></td>
                    </tr>
                    <form name="formDOemail" id="formDOemail" method="post" action="action.asp?type=submitInvoice&amp;inv_no=<%=inv_no%>&amp;#spareparts" >
                      <tr>
                       <td width="55%" align="left" bgcolor="#FFFFFF" scope="col"><!--  <strong>Email</strong>
                          <input name="emailto_DO" type="text" id="emailto_DO" value="<%=do_cust_email%>" size="50" maxlength="150" />
                          <input type="button" name="Submit523" value="Email DO" style="{width:200px}" onclick="javascript:popup('rm_do_new_email.asp?do_no=<%=do_no%>&emailto=' + formDOemail.emailto_DO.value + '&email_remark=' + formDOemail.email_remark.value,'cb17','scrollbars=yes,resizable=yes,width=600px,height=600px')" />
                          <br />
<strong><font color="#000000">Email Message: </font></strong>:
                                <input name="email_remark" type="text" id="email_remark" size="50" maxlength="200" />
                          <br />
                          <br />-->
                          <%if do_status<>"Posted" or request.Cookies("GAPS")("sloginid") = "davidhui" or request.Cookies("GAPS")("sloginid")="ERICLOH" then %>
                          <input type="button" name="CancelJob" id="CancelJob" value="Cancel DO " onclick="javascript:confirmAction('<%=do_no%>','action.asp?type=CancelDO&amp;do_no=<%=do_no%>')" />
                        <%end if%></td>
                        <td width="45%" align="right" bgcolor="#FFFFFF" 
          scope="col"><%if do_status="Open" then %>
                          <input type="button" name="SubmitJob2" id="SubmitJob2" value="Delivered DO " onclick="javascript:confirmAction('<%=do_no%>','action.asp?type=DeliveredDO&do_no=<%=do_no%>')" />
                          <%end if%>
                          <br />
                          <%if do_status="Delivered" then %>
                          <input type="button" name="DoneInvoice" id="DoneInvoice" value="Done DO " onclick="javascript:confirmAction('<%=do_no%>','action.asp?type=DoneDO&do_no=<%=do_no%>')" />
                          <%end if%>
                          <br />
                          <%if do_status="Done" then %>
                          <input type="button" name="PostedInvoice" id="PostedInvoice" value="Posted DO " onclick="javascript:confirmAction('<%=do_no%>','action.asp?type=PostedDO&do_no=<%=do_no%>')" />
                          <%end if%>
                        </td>
                      </tr>
                    </form>
                    <tr align="right">
                      <td colspan="2" bgcolor="#FFFFFF"></td>
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
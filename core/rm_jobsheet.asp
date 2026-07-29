<!-- #include file="header.asp" -->

<head>
    <style type="text/css">
        .auto-style1 {
            height: 28px;
        }
        .auto-style2 {
            height: 35px;
        }
        .auto-style3 {
            height: 40px;
        }
        .auto-style4 {
            width: 28%;
            white-space: nowrap;
        }
        .auto-style5 {
            width: 45%;
            height: 32px;
        }
        .auto-style6 {
            height: 32px;
        }
        .auto-style7 {
            width: 100%;
        }
       
.btn-select-customer {
    background: linear-gradient(135deg, #3bb78f, #0bab64);
    color: #ffffff;
    border: none;
    padding: 8px 16px;
    border-radius: 20px;
    font-size: 13px;
    font-family: Arial, sans-serif;
    cursor: pointer;
    box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    transition: all 0.2s ease;
}

.btn-select-customer:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 14px rgba(0,0,0,0.2);
    background: linear-gradient(135deg, #46c79a, #12b86b);
}

.btn-select-customer:active {
    transform: translateY(0);
    box-shadow: 0 3px 6px rgba(0,0,0,0.15);
}
    </style>
</head>

<% 
job_cust_postcode=request.querystring("job_cust_postcode")
job_cust_name=request.querystring("job_cust_name")
job_cust_code=request.querystring("job_cust_code")
job_cust_tel1=request.querystring("job_cust_tel1")
job_cust_tel2=request.querystring("job_cust_tel2")
job_cust_address=request.querystring("job_cust_address")
job_cust_cnty_id=request.querystring("job_cust_cnty_id")
job_cust_city_code=request.querystring("job_cust_city_code") 'used when pre-existing customer selected
    
set rs = server.CreateObject("adodb.recordset")

if request("job_code") <> "" or request("job_id") <> "" then	  
sql = "SELECT top 1 job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, trim(job_cust_city) as job_cust_city, job_cust_city_id,job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_model_desc, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_code, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_service_date, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_wrty_photo2,job_wrty_photo3, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_overwty_allowance,job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date, job_submittedby, job_submitteddate, " & _
		"job_doneby, job_donedate, job_postedby, job_posteddate, job_cancelledby, job_cancelleddate,job_cust_cnty_id,job_submitforclaims, job_dealer,job_payee, job_dealer_inv, " & _
		"job_total_job = ( select count(job_id) from tbljob AS a where a.job_cust_code = b.job_cust_code)  " & _
	    "FROM tbljob b WHERE job_id is not null "
		
	    if request("job_code") <> "" then 
		sql = sql & " and job_code = '" & request("job_code") & "' "
		elseif request("job_id") <> "" then 
		sql = sql & " and job_id = " & request("job_id") & " "
		end if
 
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
            job_cust_city_id= rs("job_cust_city_id")
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
            job_cust_cnty_id = rs("job_cust_cnty_id")

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
			job_tech_service_date = rs("job_tech_service_date") 
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
            job_overwty_allowance =  rs("job_overwty_allowance")
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
			job_submittedby = rs("job_submittedby")
			job_submitteddate = rs("job_submitteddate")
			job_doneby = rs("job_doneby")
			job_donedate = rs("job_donedate")
			job_postedby = rs("job_postedby")
			job_posteddate = rs("job_posteddate")
			job_cancelledby = rs("job_cancelledby")
			job_cancelleddate = rs("job_cancelleddate")
            job_submitforclaims = rs("job_submitforclaims")
			job_total_job = rs("job_total_job")
            job_dealer = rs("job_dealer")
            job_payee = rs("job_payee")
            job_dealer_inv = rs("job_dealer_inv")
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

if job_tech_code <> "" then	  
sql = "SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id,tech_wh_code " & _
	  "FROM tbltechnician WHERE tech_code = '" & job_tech_code & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			job_emailsent = rs("tech_email") 
			job_smssent = rs("tech_tel1") 
            tech_wh = rs("tech_wh_code")
		End If
		rs.Close
end if

if job_cust_postcode <> "" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & job_cust_postcode & "' "
		rs1.Open sql1,strconnect,0,1,&H0001   
		If Not rs1.EOF Then
             job_cust_state_id = rs1("state_id") 'will auto populate state
             job_cust_state=  rs1("state_name")     
        end if
    rs1.close    
end if

custlabel=""
set rs3 = server.CreateObject("adodb.recordset")
sql3 = "SELECT cust_type FROM tblcustomer WHERE cust_code = '" & job_cust_code & "' "
rs3.Open sql3,strconnect,0,1,&H0001
If Not rs3.EOF Then
    custlabel=rs3("cust_type")
    rs3.close
end if

%>

<script language="javascript">

function getPostcode(p)
{
   document.getElementById('job_cust_postcode').value = p;
   document.formorder.submit();
}

function getCountrycode(c)
{
   document.getElementById('job_cust_cnty_id').value = c;   
   document.formorder.submit();
}

function isEmpty(s) {
  return ((s == null) || (s.length == 0));
}

function confirmForm(id,orderlinks,otype) 
{
	
  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formorderparts.action = orderlinks;
	document.formorderparts.submit();
   }
}


function calctotal(unitprice,qty,discounttype,discountamt,subtotal) {

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
	}a
}

function confirmform(id,orderlinks,otype) 
{

    var tel = document.forms["formorder"]["job_cust_tel1"].value;
    var caddr = document.forms["formorder"]["job_cust_address"].value;
    caddr = caddr.trim();

    if (isEmpty(document.forms["formorder"].job_cust_name.value)) {
        alert("Please enter Cust Name.");
        document.forms["formorder"].job_cust_name.focus();
        return false;
    }

    if (isEmpty(document.forms["formorder"].job_cust_cnty_id)) {
        alert("Please select Country");
        document.forms["formorder"].job_cust_cnty_id.focus();
        return false;
    }

    if (caddr.length < 8) {
        alert("Customer address must be at least 8 characters.");
        document.forms["job_cust_address"].focus();
        return false;
    }

    if (isEmpty(document.forms["formorder"].job_cust_postcode.value)) {
        alert("Please enter Cust Postcode");
        document.forms["formorder"].job_cust_postcode.focus();
        return false;
    }    

   // var cityCode = document.getElementById("job_cust_city_code").value;
    //if (cityCode === "" || cityCode === "Select") {
      //  alert("Please select a city code.");
      //  return false;
    //}

    if (!tel || tel.trim() === "") {
        alert("Customer phone number is required.");
        document.forms["formorder"]["job_cust_tel1"].focus();
        return false;   // stop submit
    }

    var telPattern = /^[0-9+\-]+$/;

    if (tel !== "" && !telPattern.test(tel)) {
        alert("Customer phone number can only contain numbers and '+ -'");
        document.forms["formorder"]["job_cust_tel1"].focus();
        return false;
    }

    //  no consecutive dash
    if (tel.indexOf("--") !== -1) {
        alert("Customer phone number cannot contain consecutive '-'");
        document.forms["formorder"]["job_cust_tel1"].focus();
        return false;
    }

    if (document.forms["formorder"].job_type.selectedIndex == 0) {
        alert("Please Select Product Type.");
        document.forms["formorder"].job_type.focus();
        return false;
    }

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formorder.action = orderlinks;
	document.formorder.submit();
   }
}

function confirmformorder(id,orderlinks,otype) 
{
    var tel = document.forms["formorder"]["job_cust_tel1"].value;
    var caddr = document.forms["formorder"]["job_cust_address"].value;
    caddr = caddr.trim();

  if (isEmpty(document.forms["formorder"].job_cust_name.value)) {	  
   alert("Please enter Cust Name.");
   document.forms["formorder"].job_cust_name.focus();
   return false;     
   }

  if (isEmpty(document.forms["formorder"].job_cust_cnty_id)) {
        alert("Please select Country");
        document.forms["formorder"].job_cust_cnty_id.focus();
        return false;
    }

    if (caddr.length < 8) {
        alert("Customer address must be at least 8 characters.");
        document.forms["job_cust_address"].focus();
        return false;
    }

    
  if (isEmpty(document.forms["formorder"].job_cust_postcode.value)) {	  
   alert("Please enter Cust Postcode.");
   document.forms["formorder"].job_cust_postcode.focus();
   return false;     
   }     

    if (!tel || tel.trim() === "") {
        alert("Customer phone number is required.");
        document.forms["formorder"]["job_cust_tel1"].focus();
        return false;   // stop submit
    }

    var telPattern = /^[0-9+\-]+$/;

    if (tel !== "" && !telPattern.test(tel)) {
        alert("Customer phone number can only contain numbers and '+ -'");
        document.forms["formorder"]["job_cust_tel1"].focus();
        return false;
    }

    //  no consecutive dash
    if (tel.indexOf("--") !== -1) {
        alert("Customer phone number cannot contain consecutive '-'");
        document.forms["formorder"]["job_cust_tel1"].focus();
        return false;
    }
   
  if (document.forms["formorder"].job_type.selectedIndex == 0) {
	alert("Please Select Product Type.");        
	document.forms["formorder"].job_type.focus();    
	return false;     
   }
   
  if (isEmpty(document.forms["formorder"].job_Model.value)) {	  
   alert("Please enter Item Code.");
   document.forms["formorder"].job_Model.focus();
   return false;     
   } 
  
  if (document.forms["formorder"].job_tech_code.selectedIndex == 0) {
	alert("Please Select Technician.");        
	document.forms["formorder"].job_tech_code.focus();    
	return false;     
   }
   
  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formorder.action = orderlinks;
	document.formorder.submit();
   }
}

function confirmformorder2(id,orderlinks,otype) 
{
		
  if (document.forms["formorder2"].job_tech_type.selectedIndex == 0) {
	alert("Please Select type.");        
	document.forms["formorder2"].job_tech_type.focus();    
	return false;     
   }  		

  if (isEmpty(document.forms["formorder2"].job_tech_model.value)) {	  
   alert("Please enter Actual Model Number.");
   document.forms["formorder2"].job_tech_model.focus();
   return false;     
   }
   
  if (isEmpty(document.forms["formorder2"].job_tech_SN.value)) {	  
   alert("Please enter S/N.");
   document.forms["formorder2"].job_tech_SN.focus();
   return false;     
   }  

  if (document.forms["formorder2"].job_tech_faulty_code.selectedIndex == 0) {
	alert("Please Select Faulty Reason.");        
	document.forms["formorder2"].job_tech_faulty_code.focus();    
	return false;     
   }  
   	
  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.formorder2.action = orderlinks;
	document.formorder2.submit();
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
                        <td width="23%" align="right" class="titleblue1"><a href="rm_jobsheet_new_print.asp?job_code=<%=job_code%>" target="_blank">
                            <!--<img src="images/im_icon_print.gif" alt="Print | Email this page" border="0" style="border:0"/></a>-->
                            <img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/>

                        </td>
                      </tr>
                    </table></td>
                </tr>
                
                <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
                </tr>
                <tr>
                  <td width="40%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                         <td colspan="4" bgcolor="#E8E8E8" align="left" scope="col"><strong><font size="2">Customer Information <font color="#006400">(<%=custlabel%>)</font></strong></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Cust Code*</strong></font></td>
                        <td align="left">
                        <input name="job_cust_code" type="text" id="job_cust_code" style="background-color: #cccccc;" value="<%=job_cust_code%>" maxlength="50" onfocus="this.blur();" />
                        <%if Request.Cookies("GAPS")("slevel") = "cs" or Request.Cookies("GAPS")("slevel") = "sc" then %>     
                        [<a href="javascript:popup('rm_job_new_customer.asp?searchitem=tblcustomer.cust_name&searchvalue=<%=cust_code%>','cb17','scrollbars=yes,resizable=yes,width=900,height=650')">Select Customer</a>]
                        <%end if%>
                         </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Cust Name*</strong></font></td>
                        <td align="left"><input name="job_cust_name" type="text" id="job_cust_name" value="<%=job_cust_name%>" size="50" maxlength="100" /></td>
                      </tr>
                            <tr>
                         <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left">
                            <!--<input name="job_cust_cnty_code" type="hidden" id="job_cust_cnty_code" value="<%=job_cust_cnty_id%>" size="6" maxlength="50" />  -->
                            <!--<input name="job_cust_cnty" type="hidden" id="job_cust_cnty" value="<%=job_cust_cnty%>" size="30" maxlength="50" />-->                            
                                    <select name="job_cust_cnty_id" id="job_cust_cnty_id" style="width:150px"  onblur="getCountrycode(this.value)">
                                    <option value="<%=job_cust_cnty_id%>"></option>                                       
                                    <%
                                          sql = "SELECT cnty_name,cnty_id from tblcountry"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((job_cust_cnty_id)) = cstr((rs1("cnty_id"))) then					                               
                                                    response.write  "<option value='" & rs1("cnty_id") & "' selected>" & rs1("cnty_name") & "</option>"
					                            else
					                               response.write "<option value='" & rs1("cnty_id") & "'>" & rs1("cnty_name") & "</option>"
					                            end if 	                                
                                          rs1.movenext 
                                          wend                                                                       
                                          rs1.close 
                                    %>
                                        </select>            
                        </tr>
                        </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Address*</strong></font></td>
                        <td align="left"><strong>
                          <textarea name="job_cust_address" cols="50" rows="3" id="job_cust_address"><%=job_cust_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style5"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left" class="auto-style6">
                       <!-- <input name="job_cust_postcode" type="text" id="job_cust_postcode" value="<%=job_cust_postcode%>"  onblur="getPostcode(this.value)" size="20" maxlength="20" />-->                    
                            <%if job_cust_cnty_id ="129" then' %> 
                                <input name="job_cust_postcode" type="text" id="job_cust_postcode" value="<%=job_cust_postcode%>"  onblur="getPostcode(this.value)" size="10" maxlength="10" />
                                [<a href="javascript:popup('rm_new_address.asp?stateid=<%=job_cust_state_id%>&cityid=<%=job_cust_city_id%>&postcode=<%=job_cust_postcode%>','cb17','scrollbars=yes,resizable=yes,width=600,height=500')">Select Area</a>]</td>
                            <%else%> 
                                <input name="job_cust_postcode" type="text" id="job_cust_postcode" value="<%=job_cust_postcode%>" size="10" maxlength="10" /></td>
                            <%end if%>                        
                      </tr>
              <tr><td></td><td>[5 Digit-M'sia,6 Digit-S'pore]</td>  </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left">
                        <%if job_cust_cnty_id ="129" then' %> 
                            <input name="job_cust_state" type="text" id="job_cust_state" value="<%=job_cust_state%>" size="30" readonly maxlength="50" />
                            <input name="job_cust_state_id" type="hidden" id="job_cust_state_id" value="<%=job_cust_state_id%>" size="30" "readonly" maxlength="50" />
                        <%else%> 
                            <input name="job_cust_state" type="text" id="job_cust_state" value="<%=job_cust_state%>" size="30" style="background-color: #cccccc;" readonly maxlength="50" />
                            <input name="job_cust_state_id" type="hidden" id="job_cust_state_id" value="<%=job_cust_state_id%>" size="30" "readonly" maxlength="50" />
                        <%end if%> 
                      </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                        <%if job_cust_cnty_id ="129" then' %> 
                             <input name="job_cust_city_code" type="hidden" id="job_cust_city_code" value="<%=job_cust_city_id%>" size="6" maxlength="50" />  
                             <input name="job_cust_city" type="hidden" id="job_cust_city" value="<%=job_cust_city%>" size="30" maxlength="50" />
                             <select name="job_cust_city_id" id="job_cust_city_id" style="width:150px">
                             <option value="<%=job_cust_city_id%>"></option>
                                    <%
                                          sql = "SELECT distinct city_id, post_office FROM tblpostcode where postcode = '" & job_cust_postcode & "'"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((job_cust_city_id)) = cstr((rs1("city_id"))) then
					                               response.write "<option value='" & rs1("city_id") & "' selected>" & rs1("post_office") & "</option>"
					                            else
					                               response.write "<option value='" & rs1("city_id") & "'>" & rs1("post_office") & "</option>"
					                            end if 	                                
                                          rs1.movenext 
                                          wend                                                                       
                                          rs1.close 
                                    %>
                                        </select> 
                            <%else%>
                                    <input name="job_cust_city_code" type="hidden" id="job_cust_city_code" value="<%=job_cust_city_id%>" size="6" maxlength="50" /> 
                                    <input name="job_cust_city" type="text" id="job_cust_city" value="<%=job_cust_city%>" size="30" maxlength="50" />
                            <%end if%>
                        </td>
                      </tr>
                        
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="job_cust_email" type="text" id="job_cust_email" value="<%=job_cust_email%>" size="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><label for="job_cust_tel1"></label>
                        <input name="job_cust_tel1" type="text" id="job_cust_tel1" value="<%=job_cust_tel1%>" size="30" maxlength="50" />                         
                        <input type="button" name="button19" id="button19" class="btn-select-customer" value="Check Record" onclick="javascript:popup('rm_cust_phone_check.asp?phone1=' + formorder.job_cust_tel1.value+ '&phone2=' + formorder.job_cust_tel2.value,'cb18','scrollbars=yes,menubar=no,location=no,resizable=no,width=1200px,height=600px')" />
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="job_cust_tel2" type="text" id="job_cust_tel2" value="<%=job_cust_tel2%>" size="30" maxlength="50" /></td>
                      </tr>

                        <tr >
                             <td bgcolor="#E8E8E8" colspan="4" align="left"><strong>Dealer information</strong></td>
                         </tr>
 <tr>
      <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Payee (Received From)</strong></font></td>
      <td><strong>
 <select name="job_payee" id="job_payee" style="width:300px;">
   <option value=""></option>
   <option value="Dealer" <%if job_payee="Dealer" then response.write " selected"%>>Dealer</option>
   <option value="Customer" <%if job_payee="Customer" then response.write " selected"%>>Customer</option>
 </select>
 </strong></td>
     </tr>
 <tr>
   <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong> Dealer DO/INV No:</strong></font></td>
   <td valign="top"><input name="job_dealer_inv" type="text" id="job_dealer_inv" value="<%=job_dealer_inv%>"  style="width:300px; box-sizing:border-box;" maxlength="45" /></td>
 </tr>


<tr>

                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Dealer</strong></font></td>
                        <td valign="top">
                            <select name="job_dealer" id="job_dealer" style="width:300px">
                             <option value="<%=job_dealer%>"></option>
                                    <%
                                          sql = "SELECT cust_code, cust_name, cust_address,cust_city FROM tblcustomer where cust_type='Dealer' order by cust_name"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		  
                                        
                                            If Trim("" & job_dealer) = Trim("" & rs1("cust_code")) Then
                                                        Response.Write "<option value='" & rs1("cust_code") & "' selected>" & _
                                                        Server.HTMLEncode(rs1("cust_name") & " - " & rs1("cust_address") & " " & rs1("cust_city")) & _
                                                        "</option>"
                                           Else
                                                        Response.Write "<option value='" & rs1("cust_code") & "'>" & _
                                                        Server.HTMLEncode(rs1("cust_name") & " - " & rs1("cust_address") & " " & rs1("cust_city")) & _
                                                       "</option>"
                                           End If                                
                                          rs1.movenext 
                                          wend                                                                       
                                          rs1.close 
                                    %>
                                        </select>
                      </tr>
                <!--080326  create a more user friendly window to select dealers, as the list can get larger-->

                        <!--<tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Dealer</strong></font></td>
                        <td valign="top">
                         <input name="job_dealer" type="text" id="job_dealer" value="<%=job_dealer%>" size="30" maxlength="50" />
                           [<a href="javascript:popup('rm_job_dealer_list.asp?searchitem=md_type&amp;tech_wh=<%=tech_wh%>&amp;searchvalue=<%=cust_code%>&amp;formname=formorderparts&fieldname=jobp_partcode&fieldname1=jobp_desc&md_type=' + formorder2.job_tech_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                      </tr>-->

					 <!-- <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Branch Code</strong></font></td>
                        <td valign="top"><strong>
                          <input name="job_branch_code" type="text" id="job_branch_code" value="<%=job_branch_code%>" size="30" maxlength="50" />                          
                        </strong></td>
                      </tr>-->
                    <tr >
                         <td bgcolor="#E8E8E8" colspan="4" align="left"><strong>Job Information</strong></td>
                     </tr>

                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Jobs Count</strong></font></td>
                        <td><input name="job_total_job" id="job_total_job" style="background-color: #cccccc; text-align:center;" value="<%=job_total_job%>" size="5" /> </td>                       
                      </tr>

					  <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td valign="top"><strong>
                          <textarea name="job_remark" cols="50" rows="3" id="job_remark"><%=job_remark%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                         <td valign="top"><%=job_createdby%> @ <%=chkdatetime(job_createddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>JS Received Ack</strong></font></td>
                        <td valign="top"><%=job_JS_receivedby%> @ <%=chkdatetime(job_JS_receiveddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Submitted by</strong></font></td>
                        <td valign="top"><%=job_submittedby%> @ <%=chkdatetime(job_submitteddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Done by</strong></font></td>
                        <td valign="top"><%=job_doneby%> @ <%=chkdatetime(job_donedate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Posted by</strong></font></td>
                        <td valign="top"><%=job_postedby%> @ <%=chkdatetime(job_posteddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Cancelled  by</strong></font></td>
                        <td valign="top"><%=job_cancelledby%> @ <%=chkdatetime(job_cancelleddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="52%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="0" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Job Sheet Information</font></strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job  
                          No.</strong> </td>
                        <td align="left" style="background-color: #cccccc;" ><strong><%=job_code%> 
                          <input type="hidden" name="job_code" id="job_code" value="<%=job_code%>" />
                          <input type="hidden" name="job_id" id="job_id" value="<%=job_id%>" />
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
                        <td><font color="#000000"><strong>
                          <input name="job_purchase_date" type="text" id="job_purchase_date" value="<%=chkdate(job_purchase_date)%>" size="12" maxlength="20" onfocus="this.blur()" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.formorder.job_purchase_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                      </tr>
                      <tr align="left" >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
                        <td><label for="job_onlineWrtyNo"></label> <input name="job_onlineWrtyNo" type="text" id="job_onlineWrtyNo" value="<%=job_onlineWrtyNo%>" maxlength="50" /></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                        <td><strong>
                        <select name="job_onlineWrtyStatus" id="job_onlineWrtyStatus">
                          <option value=""></option>
                          <option value="Over" <%if job_onlineWrtyStatus="Over" then response.write " selected"%>>Over</option>
                          <option value="Under" <%if job_onlineWrtyStatus="Under" then response.write " selected"%>>Under</option>
                        </select>
                        </strong></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Product Type*</strong></font></td>
                        <td align="left"><select name="job_type" id="job_type">
                          <option value=""></option>
                          <option value="CF" <%if job_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                          <option value="WH" <%if job_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                        </select></td>
                       <!-- <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Serial #</strong></font></td>
                        <td align="left"><input name="job_SN_no" type="text" id="job_SN_no" value='<%=job_SN_no%>' maxlength="50" /></td>-->
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Item Code*</strong></font></td>
                        <td colspan="3" align="left"><label for="job_cust_city"></label>
                          <label for="job_Model"></label>
                          <input name="job_Model" type="text" id="job_Model" value="<%=job_Model%>" size="20" maxlength="50" />                          
                           <%if Request.Cookies("GAPS")("slevel") = "cs" or Request.Cookies("GAPS")("slevel") = "sc" then %>     
                          <label for="job_type">[<a href="javascript:popup('rm_job_new_model.asp?searchitem=md_code&searchvalue=<%=cust_code%>&md_type=' + formorder.job_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </label>
                          <%end if%>
                          </td>
                        </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model Desc</strong></font></td>
                        <td colspan="3" align="left"><input name="job_Model_desc" type="text" id="job_Model_desc" value="<%=job_Model_desc%>" size="60" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
                        <td colspan="3" align="left">
                          <select name="job_faulty_reason_cs" id="job_faulty_reason_cs" style="width:250px">
                          <option value="" <%if job_faulty_reason_cs="" then response.write " selected"%>>Common</option>
                          <option value="">--------------------------------------------</option>
                          <option value="Installation" <%if job_faulty_reason_cs="Installation" then response.write " selected"%>>Installation</option>
                          <option value="" <%if job_faulty_reason_cs="" then response.write " selected"%>>Description of Ceiling Fan Faults</option>
                          <option value="">--------------------------------------------</option>
                          <option value="Ceiling Fan Not Functioning" <%if job_faulty_reason_cs="Ceiling Fan Not Functioning" then response.write " selected"%>>Ceiling Fan Not Functioning</option>
                          <option value="No Speed control /can`t control /can`t do paring" <%if job_faulty_reason_cs="No Speed control /can`t control /can`t do paring" then response.write " selected"%>>No Speed control /can't control /can't do paring</option>
                          <option value="Fan wobbling" <%if job_faulty_reason_cs="Fan wobbling" then response.write " selected"%>>Fan wobbling</option>
                          <option value="Fan slow" <%if job_faulty_reason_cs="Fan slow" then response.write " selected"%>>Fan slow</option>
                          <option value="Fan noisy" <%if job_faulty_reason_cs="Fan noisy" then response.write " selected"%>>Fan noisy</option>
                          <option value="No light" <%if job_faulty_reason_cs="No light" then response.write " selected"%>>No light</option>
                          <option value="No wind /air" <%if job_faulty_reason_cs="No wind /air" then response.write " selected"%>>No wind /air</option>
                          <option value="Fan rusty" <%if job_faulty_reason_cs="Fan rusty" then response.write " selected"%>>Fan rusty</option>
                          <option value="Apps no function" <%if job_faulty_reason_cs="Apps no function" then response.write " selected"%>>Apps no function</option>
                          <option value="Info only(no job created)Short spare parts,refer respective dealer /salesman" <%if job_faulty_reason_cs="Info only(no job created)Short spare parts,refer respective dealer /salesman" then response.write " selected"%>>Info only(no job created)Short spare parts,refer respective dealer /salesman</option>
                          <option value="" <%if job_faulty_reason_cs="" then response.write " selected"%>>Description of Water Heater Faults </option>
                          <option value="">--------------------------------------------</option>
                          <option value="Water Heater Not Functioning" <%if job_faulty_reason_cs="Water Heater Not Functioning" then response.write " selected"%>>Water Heater Not Functioning</option>
                          <option value="Water Heater No Power" <%if job_faulty_reason_cs="Water Heater No Power" then response.write " selected"%>>Water Heater No Power</option>
                          <option value="Pump noisy" <%if job_faulty_reason_cs="Pump noisy" then response.write " selected"%>>Pump noisy</option>
                          <option value="Burning smell" <%if job_faulty_reason_cs="Burning smell" then response.write " selected"%>>Burning smell</option>
                          <option value="Pump not working" <%if job_faulty_reason_cs="Pump not working" then response.write " selected"%>>Pump not working</option>
                          <option value="No hot water" <%if job_faulty_reason_cs="No hot water" then response.write " selected"%>>No hot water</option>
                          <option value="Water heater tripped/RCD tripped" <%if job_faulty_reason_cs="Water heater tripped/RCD tripped" then response.write " selected"%>>Water heater tripped/RCD tripped</option>
                          <option value="Water heater internal leaking" <%if job_faulty_reason_cs="Water heater internal leaking" then response.write " selected"%>>Water heater internal leaking</option>
                          <option value="Water not hot enough" <%if job_faulty_reason_cs="Water not hot enough" then response.write " selected"%>>Water not hot enough</option>
                          <option value="Alarm Beeping" <%if job_faulty_reason_cs="Alarm Beeping" then response.write " selected"%>>Alarm Beeping</option>
                          <option value="Wrong Installation" <%if job_faulty_reason_cs="Wrong Installation" then response.write " selected"%>>Wrong Installation</option>
                          <option value="Info only(no job created)Short spare parts,refer respective dealer /salesman" <%if job_faulty_reason_cs="Info only(no job created)Short spare parts,refer respective dealer /salesman" then response.write " selected"%>>Info only(no job created)Short spare parts,refer respective dealer /salesman</option>
                          </select>
						</td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Desc*</strong></font></td>
                        <td colspan="3" align="left"><strong>
                          <textarea name="job_faulty_desc" cols="60" rows="3" wrap="virtual" id="job_faulty_desc"><%=job_faulty_desc%></textarea>
                        </strong></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Reported By *</strong></font></td>
                        <td colspan="2"><label for="job_reportedby"></label>
                          <select name="job_reportedby" id="job_reportedby">
                            <option value="Call-in" <%if job_reportedby="Call-in" then response.write " selected"%>>Call-in</option>
                            <option value="Walk-In" <%if job_reportedby="Walk-In" then response.write " selected"%>>Walk-In</option>
                            <option value="Dealer" <%if job_reportedby="Dealer" then response.write " selected"%>>Dealer</option>
                            <option value="Sales" <%if job_reportedby="Sales" then response.write " selected"%>>Sales</option>
                            <option value="Website" <%if job_reportedby="Website" then response.write " selected"%>>Website</option>
                            <option value="Technician" <%if job_reportedby="Technician" then response.write " selected"%>>Technician</option>
                            <option value="App" <%if job_reportedby="App" then response.write " selected"%>>App</option>							
                          </select></td>
                        <td><font color="#000000"><strong>
                         <%if Request.Cookies("GAPS")("slevel") = "cs" or Request.Cookies("GAPS")("slevel") = "sc" then %>
                        <a href="javascript:popup('rm_job_new_schedule.asp?searchitem=job_date&job_date=<%=chkdate(job_date)%>&job_cust_state=' + formorder.job_cust_state.options[formorder.job_cust_state.selectedIndex].text,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Schedule<img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a>
                        <%end if%>
                        </strong></font></td>
                      </tr>
                      <tr >
                           <td bgcolor="#E8E8E8" colspan="4" align="left"><strong>Appointment Details</strong></td>
                        </tr>
                        <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Date*</strong></font></td>
                        <td><font color="#000000"><strong>
                          <input name="job_appointment_date" type="text" id="job_appointment_date" value="<%=chkdate(job_appointment_date)%>" size="12" onfocus="this.blur()" />
                        <a href="javascript:void(null)" onclick="window.dateField = document.formorder.job_appointment_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Time </strong></font></td>
                        <td><label for="job_appointment_time"></label>
                        <input name="job_appointment_time" type="text" id="job_appointment_time" value="<%=job_appointment_time%>" /></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician*</strong></font></td>
                        <td colspan="3">
                        <select name="job_tech_code" id="job_tech_code">
                        <option value=""></option>
                        <option value="resolved_no_appt" <%if job_tech_code="resolved_no_appt" then response.write " selected"%>>Issue resolve without appt</option>
                                            <%			
				sql = "SELECT tech_id, tech_code, tech_name FROM tbltechnician where tech_status = 'Y' and (tech_type='TPC' or tech_type='IHT' or tech_type='IHC' or tech_type='IC' or tech_type='SGT') "	
				
				if Request.Cookies("GAPS")("slevel") = "technician" then 
				sql = sql & "and tech_code='" & request.Cookies("GAPS")("job_tech_code") & "'"
				end if
				
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (job_tech_code) = (rs("tech_code")) then
					  response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                                          </select>
                        </td>
                      </tr>
                      <tr >
                        <td valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Remark<br />
                        </strong></font></td>
                        <td colspan="3"><strong>
                          <textarea name="job_appointment_remark" cols="60" rows="3" wrap="virtual" id="job_appointment_remark"><%=job_appointment_remark%></textarea>
                        </strong></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                
                <tr>
                  <td align="left" valign="top" bgcolor="#FFFFFF"><strong>Technician Alert</strong><br />
                   <!-- Email to
                    <input name="emailto" type="text" id="emailto" value="<%=job_emailsent%>" size="40" maxlength="100" />
                    <input type="button" name="Submit523" value="Email Job" style="{width:200px}" />
                    <br />
                    <strong><font color="#000000">Email Message: </font></strong>:
                    <input name="email_remark" type="text" id="email_remark" size="40" maxlength="200" />
                    <br />
                    <br />-->
                    SMS to Technician
                    <input name="job_smssent" type="text" id="job_smssent" value="<%=job_smssent%>" size="15" maxlength="15" />
                    <input type="button" name="button3" id="button7" value="Resend SMS Now" onclick="javascript:popup('rm_jobsheet_sendsms.asp?job_code=<%=job_code%>&smsnumber=' + formorder.job_smssent.value,'cb18','scrollbars=yes,resizable=yes,width=600px,height=600px')" />
                  </td>
                  <td align="right" valign="top" bgcolor="#FFFFFF"><input type="button" name="button1" id="button1" value="<%=actionname%>" onclick="javascript:confirmform('<%=job_code%>','action.asp?type=<%=stype%>','dummy')" />
                  <%if job_status="Open" and job_code <> "" then %>
                  <input type="button" name="SubmitJob" id="SubmitJob" value="Submit Job" onclick="javascript:confirmformorder('<%=job_code%>','action.asp?type=submitJob&job_code=<%=job_code%>')" />
                   <%end if%>
                  </td>
                </tr>
                <tr>
                  <td align="left" valign="top" bgcolor="#FFFFFF">
                  <%if job_status="Open" then %>
                  <input type="button" name="CancelJob" id="CancelJob" value="Cancel Job" onclick="javascript:confirmform('<%=job_code%>','action.asp?type=cancelJob&job_code=<%=job_code%>')" />
                  <%end if%>
                  </td>
                  <td align="right" valign="top" bgcolor="#FFFFFF">
                    
                    <%if job_status <> "Open" and job_status <> "Cancel"  and job_status <> "Posted" then %>
                    <input type="button" name="button2" id="button2" value="Reset Status to Open" onclick="javascript:confirmAction('<%=job_code%>','action.asp?type=resetJobOpen&job_code=<%=job_code%>')" />
                    <%end if%>
                    
                    <%if job_code <> "" then %>
                    <input type="button" name="button10" id="button12" value="Duplicate Job" onclick="javascript:confirmAction('<%=job_code%>','action.asp?type=duplicateJob&job_code=<%=job_code%>')" />
                    <%end if%>
                  </td>                
                </tr>
                </form>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
  
              
	<%
    if job_code <> "" and job_status<>"Open" then 
    %>  
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <form name="formorder2" id="formorder2" method="post" action="action.asp?type=editjob_Technical" enctype="multipart/form-data">
                      <tr align="right">
                        <td colspan="2" valign="top" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                          <tbody>
                            <tr bgcolor="#E8E8E8">
                              <td colspan="4" align="left" bgcolor="#E8E8E8" scope="col"><strong>Technical Findings</strong></td>
                              </tr>
                            <tr >
                              <td width="12%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type*</strong></font></td>
                              <td width="38%" align="left"><select name="job_tech_type" id="job_tech_type">
                                <option value=""></option>
                                <option value="CF" <%if job_tech_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                                <option value="WH" <%if job_tech_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                              </select>
							  <input type="text" name="job_code" id="job_code" value="<%=job_code%>" /></td>
                              <td width="22%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Actual Model Number*</strong></font></td>
                              <td width="28%" align="left">
                                <input name="job_tech_model" type="text" id="job_tech_model" value="<%=job_tech_model%>" size="20" maxlength="50" />
                                <label for="job_type2">[<a href="javascript:popup('rm_job_new_model_tech.asp?searchitem=md_code&searchvalue=<%=cust_code%>&md_type=' + formorder2.job_tech_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </label></td>
                              </tr>
                            <tr >
                              <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tax Invoice*</strong></font></td>
                              <td align="left"><select name="job_tech_tax_invoice" id="job_tech_tax_invoice">
                                <option value="No" <%if job_tech_tax_invoice="No" then response.write " selected"%>>No</option>
                                <option value="Yes" <%if job_tech_tax_invoice="Yes" then response.write " selected"%>>Yes</option>

                              </select></td>
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong> Model Desc</strong></font></td>
                              <td align="left"><input name="job_tech_model_desc" type="text" id="job_tech_model_desc" value="<%=job_tech_model_desc%>" size="40" maxlength="200" /></td>
                            </tr>
                            <tr >
                              <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Part S/N*</strong></font></td>
                              <td align="left"><input name="job_tech_SN" type="text" id="job_tech_SN" value='<%=job_tech_SN%>' size="50" maxlength="50" /></td>
                              <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Service Status<strong>*</strong></font></strong></td>
                              <td align="left"><label for="job_tech_status"></label>
                                <select name="job_tech_status" id="job_tech_status">
                                  <option value="Pending" <%if job_tech_status="Pending" then response.write " selected"%>>Pending</option>
                                  <option value="Done" <%if job_tech_status="Done" then response.write " selected"%>>Done</option>
                                  <option value="Send Back to HQ" <%if job_tech_status="Send Back to HQ" then response.write " selected"%>>Send Back to HQ</option>
                                </select></td>
                              </tr>
                            <tr>
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
                              <td align="left"><select name="job_tech_faulty_code" id="job_tech_faulty_code" style="width:300px">
                                <option value=""></option>
                                <%			
				sql = "SELECT fr_id, fr_code, fr_description, fr_type FROM tblfaultyreason where fr_id is not null and fr_status='Y' " 
				
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
                              <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Service Date</font></strong></td>
                              <td align="left"><font color="#000000"><strong>
                                <input name="job_tech_service_date" type="text" id="job_tech_service_date" value="<%=chkdate(job_tech_service_date)%>" size="12" maxlength="20" onfocus="this.blur()" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.formorder2.job_tech_service_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                              </tr>
                            <tr >
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Repair Action / <br />
Remark*<br />
                              </strong>(150 Chars)<br />
<strong></strong></font><br />
                              <strong><a href="javascript:popup('rm_job_new_repaitaction.asp?searchitem=ra_repairaction&amp;searchvalue=<%=cust_code%>&amp;md_type=' + formorder.job_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')"><font color="#FFFFFF">[Select]</font></a></strong> <br />
                              <br /></td>
                              <td align="left"><strong>
                                <textarea name="job_tech_faulty_action" cols="60" rows="6" id="job_tech_faulty_action"><%=job_tech_faulty_action%></textarea>
                              </strong></td>
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Product Collection Date</strong></font></td>
                              <td align="left"><font color="#000000"><strong>
                                <input name="job_tech_product_collectdate" type="text" id="job_tech_product_collectdate" value="<%=chkdate(job_tech_product_collectdate)%>" size="12" maxlength="20" onfocus="this.blur()" />
                                <a href="javascript:void(null)" onclick="window.dateField = document.formorder2.job_tech_product_collectdate;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                            </tr>
                            <tr >
                              <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Actual Warranty Status*</strong></font></td>
                              <td align="left"><select name="job_actual_wrty_status" id="job_actual_wrty_status">
                                <option value="Under" <%if job_actual_wrty_status="Under" then response.write " selected"%>>Under</option>
                                <option value="Over" <%if job_actual_wrty_status="Over" then response.write " selected"%>>Over</option>
                              </select></td>
                              <td align="left"></td>
                              <td align="left"></td>
                            </tr>
                            <tr >
                              <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Warranty Photo<br />
(Optional) </strong></font></td>
                              <td align="left" valign="top"><input type="file" name="job_wrty_photo" id="job_wrty_photo" /><br/>
                              <input type="file" name="job_wrty_photo2" id="job_wrty_photo2" /><br/>
                              <input type="file" name="job_wrty_photo3" id="job_wrty_photo3" />
                        
                                <a href="shared/<%=job_wrty_photo%>" target="_blank"><img src="shared/<%=job_wrty_photo%>" alt="Click on to Pop-up" width="100" border="0" /></a>
                                  <a href="shared/<%=job_wrty_photo2%>" target="_blank"><img src="shared/<%=job_wrty_photo2%>" alt="Click on to Pop-up" width="100" border="0" /></a>
                                <a href="shared/<%=job_wrty_photo3%>" target="_blank"><img src="shared/<%=job_wrty_photo3%>" alt="Click on to Pop-up" width="100" border="0" /></a></td>
                              <td align="left" valign="top"></td>
                              <td align="left" valign="top"></td>
                              </tr>
                            </tbody>
                        </table></td>
                      </tr>
                    <tr align="right">
                      <td colspan="2" scope="col" class="auto-style3">
                       <input type="button" name="button7" id="button10" value="Save" onclick="javascript:confirmformorder2('<%=job_code%>','action.asp?type=editjob_Technical','Save')" />
                        <%if job_status <> "Posted" then ' 28/02/2024 %>
                                <input type="button" name="SubmitJob2" id="SubmitJob2" value="Update Job Status " onclick="javascript:confirmformorder2('<%=job_code%>','action.asp?type=DoneJob')" />
                        <%end if%>
                          </td>
                        </tr>
                    </form>
                   <tr>
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                      <tr>
                       <td colspan="2" scope="col">
                        <table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" class="auto-style7">                     
                        <tr bgcolor="#2f8e88"><td colspan="4" align="center"><font color="#d8e3e2" SIZE="3"><strong>CUSTOMER SUPPORT (OFFICE) SECTION</strong></font></td></tr>
                    </table>
                       </td>
                    </tr>

                      <form name="formorder3" id="formorder3" method="post" action="action.asp?type=editjob_hq" >
                    <tr>
                      <td colspan="2" scope="col">
                          <table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV8">
                        <tbody><tr bgcolor="#E8E8E8">
                            <td colspan="4" bgcolor="#E8E8E8" scope="col"><font size="2"><strong>Office Remarks</strong></font></td>
                          </tr>
                          <tr >
                            <td width="12%" rowspan="2" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Remark<br />spare-parts
                            </strong>(150 Chars) </font></td>
                            <td width="38%" rowspan="2" align="left"><strong>
                              <textarea name="job_hq_remark" cols="60" rows="3" wrap="virtual" id="job_hq_remark"><%=job_hq_remark%></textarea>
                            </strong></td>
                            <td width="29%" align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>HQ Received Date</strong></font></td>
                            <td width="21%" align="left"><label for="select3"><font color="#000000"><strong>
                              <input name="job_hq_received_date" type="text" id="job_hq_received_date" value="<%=chkdate(job_hq_received_date)%>" size="12" maxlength="20" onfocus="this.blur()" />
                              <a href="javascript:void(null)" onclick="window.dateField = document.formorder3.job_hq_received_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></label></td>
                          </tr>
                          <tr >
                            <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Repair Date</strong></font></td>
                            <td align="left"><label for="job_hq_category_code"><font color="#000000"><strong>
                              <input name="job_repair_date" type="text" id="OrderCancellationDate11" value="<%=chkdate(job_repair_date)%>" size="12" maxlength="20" onfocus="this.blur()" />
                              <a href="javascript:void(null)" onclick="window.dateField = document.formorder3.job_repair_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></label></td>
                          </tr>
                          <tr >
                            <td nowrap="nowrap" bgcolor="#475387">&nbsp;</td>
                            <td align="left">&nbsp;</td>
                            <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Return to Technician Date</strong></font></td>
                            <td align="left"><font color="#000000"><strong>
                              <input name="job_return_tech_date" type="text" id="job_return_tech_date" value="<%=chkdate(job_return_tech_date)%>" size="12" maxlength="20" onfocus="this.blur()" />
                              <a href="javascript:void(null)" onclick="window.dateField = document.formorder3.job_return_tech_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                          </tr>
                          <tr >
                            <td width="12%" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Category Code</strong></font></td>
                            <td width="38%" align="left"><select name="job_hq_category_code" id="job_hq_category_code">
                              <option value="MD" <%if job_hq_category_code="MD" then response.write " selected"%>>MD</option>
                              <option value="DS" <%if job_hq_category_code="DS" then response.write " selected"%>>DS</option>
                              <option value="WI" <%if job_hq_category_code="WI" then response.write " selected"%>>WI</option>
                              <option value="CF" <%if job_hq_category_code="CF" then response.write " selected"%>>CF</option>
                              <option value="IT" <%if job_hq_category_code="IT" then response.write " selected"%>>IT</option>
                            </select>
                              <input type="hidden" name="job_code" id="job_code" value="<%=job_code%>" /></td>
                            <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Return to Customer Date</strong></font></td>
                            <td align="left"><font color="#000000"><strong>
                              <input name="job_tech_returntoCustDate" type="text" id="job_tech_returntoCustDate" value="<%=chkdate(job_tech_returntoCustDate)%>" size="12" maxlength="20" onfocus="this.blur()" />
                              <a href="javascript:void(null)" onclick="window.dateField = document.formorder3.job_tech_returntoCustDate;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                          </tr>
                        </tbody>
                      </table></td>
                    </tr>
                    <tr>
                      <td align="left" 
          scope="col" class="auto-style2"></td>
                      <td align="right" 
          scope="col" class="auto-style2"><input type="submit" name="button8" id="button8" value="Save" /></td>
                    </tr>
                      </form>
          
                    
                    <tr>
                      <td colspan="2" 
          scope="col">&nbsp;</td>
                    </tr>
                    
                    <tr>
                      <td colspan="2" bgcolor="#E8E8E8" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td><strong><font size="2">Spare-Part & Services Charges<a name="spareparts" id="spareparts"></a></font></strong></td>
                          </tr>
                      </table></td>
                    </tr>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Spare Part 
                            Code</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Description</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Faulty Code</strong></font></td>                        
                          <td align="right"><font color="#FFFFFF"><strong>RCP Price (RM)</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Discount </strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Total 
                            Amt (RCP)</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Action</strong></font></td>
                        </tr>
                        
                      
                        
<%

if request("jobp_id") <> "" then
		sql = "SELECT jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, jobp_netcost, jobp_qty, jobp_subtotal, jobp_faultycode " & _
	          "FROM tbljob_parts where jobp_id = '" & request("jobp_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   jobp_id = rs("jobp_id")
		   job_code = rs("job_code")
		   jobp_partcode = rs("jobp_partcode")
		   jobp_desc = rs("jobp_desc")
		   jobp_unitcost = rs("jobp_unitcost")
		   jobp_discountamt = rs("jobp_discountamt")
		   jobp_discounttype = rs("jobp_discounttype")  
		   jobp_netcost = rs("jobp_netcost")   
		   jobp_qty = rs("jobp_qty")
		   jobp_subtotal = rs("jobp_subtotal")
		   jobp_faultycode = rs("jobp_faultycode")
        end if
		rs.close
		sbutton = "Update"
		stype="editJobDetail"	
else
		sbutton = "Add"
		stype="addJobDetail"
		jobp_qty = "1"	
		jobp_unitcost = "0.00"	
		jobp_discountamt = "0.00"
		jobp_netcost = "0.00"	
		jobp_subtotal = "0.00"	
		jobp_faultycode = ""
		
		if job_actual_wrty_status="Over" and job_totallabourAmt = "" then
		job_totallabourAmt = 53
		end if
end if

%>

                      <%if job_status<>"Posted" and (job_inv_no="" or isnull(job_inv_no)) then %> 
                        <form name="formorderparts" id="formorderparts" method="post" action="action.asp?type=<%=stype%>" >
                        <tr>
                          <td colspan="2" bgcolor="#666666">
                            <input name="jobp_partcode" type="text" id="jobp_partcode" value="<%=jobp_partcode%>" size="20" maxlength="50" />
                              <%if job_faulty_reason_cs = "Installation" then%>
                                [<a href="javascript:popup('rm_job_new_spareparts.asp?searchitem=md_type&amp;tech_wh=<%=tech_wh%>&amp;faulty=<%="1"%>&amp;searchvalue=<%=cust_code%>&amp;formname=formorderparts&fieldname=jobp_partcode&fieldname1=jobp_desc&md_type=' + formorder2.job_tech_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                            <%else %>
                                [<a href="javascript:popup('rm_job_new_spareparts.asp?searchitem=md_type&amp;tech_wh=<%=tech_wh%>&amp;searchvalue=<%=cust_code%>&amp;formname=formorderparts&fieldname=jobp_partcode&fieldname1=jobp_desc&md_type=' + formorder2.job_tech_type.value,'cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                            <%end if%> 
                          <td align="left" bgcolor="#666666">
                            <input name="jobp_desc" type="text" id="jobp_desc" value="<%=jobp_desc%>" size="30" maxlength="100" /></td>
                          <td align="left" bgcolor="#666666"><select name="jobp_faultycode" id="jobp_faultycode" style="width:80px">
                            <option value=""></option>
                            <%			
				sql = "SELECT fr_id, fr_code, fr_description, fr_type FROM tblfaultyreason where fr_id is not null and fr_status='Y' " 
				
				if job_tech_type <> "" then 
				sql = sql & " and fr_type='" & job_tech_type & "'"	
				
				elseif job_type <> "" then 
				sql = sql & " and fr_type='" & job_type & "'"	 
				
				end if
				
				sql = sql & " order by fr_code"	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (jobp_faultycode) = (rs("fr_code")) then
					  response.write "<option value='" & rs("fr_code") & "' selected>" & rs("fr_code") & " - " & rs("fr_description")  & "</option>"
					  else
					  response.write "<option value='" & rs("fr_code") & "'>" & rs("fr_code") & " - " & rs("fr_description")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                          </select></td>
                          <td align="right" bgcolor="#666666"><font color="#FFFFFF">
                            <input type="hidden" name="jobp_id" id="jobp_id" value="<%=jobp_id%>" />
                            <input type="hidden" name="job_code" id="job_code" value="<%=job_code%>" />
                            <input type="hidden" name="job_id" id="job_id" value="<%=job_id%>" />
                            <input name="jobp_unitcost" type="text" id="jobp_unitcost" style="text-align:right; background-color: #cccccc;" onkeydown="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);" onkeyup="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);" value="<%=jobp_unitcost%>" size="5" maxlength="10" />
                          </font></td>
                          <td align="right" bgcolor="#666666"><input name="jobp_qty" type="text" id="jobp_qty" style="text-align:right" onkeydown="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);" onkeyup="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);" value="<%=jobp_qty%>" size="5" maxlength="5" /></td>
                          <td align="right" nowrap="nowrap" bgcolor="#666666"><font color="#FFFFFF">
                            <input name="jobp_discountamt" type="text" id="jobp_discountamt" value="<%=jobp_discountamt%>" size="5" onkeydown="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);" onkeyup="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);" style="text-align:right" />
                             <select name="jobp_discounttype" id="jobp_discounttype" onchange="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);" onkeyup="calctotal(document.formorderparts.jobp_unitcost.value, document.formorderparts.jobp_qty.value, document.formorderparts.jobp_discounttype.value, document.formorderparts.jobp_discountamt.value, document.formorderparts.jobp_subtotal);">
                              <option value="%" <%if jobp_discounttype = "%" then response.write " selected"%>>%</option>
                              <option value="RM" <%if jobp_discounttype = "RM" then response.write " selected"%>>RM</option>
                            </select>
                          </font></td>
                          <td align="right" bgcolor="#666666"><input name="jobp_subtotal" type="text" id="jobp_subtotal" style="text-align:right; background-color: #cccccc;" onfocus="this.blur();" value="<%=jobp_subtotal%>" size="10" maxlength="10" /></td>
                          <td align="center" bgcolor="#666666"><input type="button" name="additems33" id="additems33" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("jobp_id")%>','action.asp?type=<%=stype%>','<%=jobp_subtotal%>');" /></td>
                        </tr>
                        </form>
                    <%end if%>    
                        
 <%				i = 1
				sql1 = "SELECT jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, jobp_netcost, jobp_qty, jobp_subtotal, jobp_faultycode " & _
	                   "FROM tbljob_parts where job_code = '" & job_code & "' order by jobp_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%> 
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="left"><%=rs1("jobp_partcode")%></td>
                          <td align="left"><%=rs1("jobp_desc")%></td>
                          <td align="left"><%=rs1("jobp_faultycode")%></td>
                          <td align="right"><%=chknumber2(rs1("jobp_unitcost"))%></td>
                          <td align="right"><%=rs1("jobp_qty")%></td>
                          <td align="right">- <%=chknumber2(rs1("jobp_unitcost")-rs1("jobp_netcost"))%></td>
                          <td align="right"><%=chknumber2(rs1("jobp_subtotal"))%></td>
                          <td align="center" nowrap="nowrap">
                       <%if job_status<>"Posted" and (job_inv_no="" or isnull(job_inv_no)) then %> 
                          <input type="button" name="aditem211" id="aditem211" value="Edit" onclick="document.location.href='rm_jobsheet.asp?jobp_id=<%=rs1("jobp_id")%>&job_code=<%=job_code%>#spareparts'" />
                          <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmDel('<%=rs1("jobp_partcode")%>','action.asp?type=delJobDetail&jobp_id=<%=rs1("jobp_id")%>&job_code=<%=job_code%>')" />
                       <%end if%>
                          </td>
                        </tr>
 <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                        
                        <form name="formorderparts_total" id="formorderparts_total" method="post" action="action.asp?type=editJobDetailTotal&job_code=<%=job_code%>&#spareparts" >
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="7" align="right"><strong>Total Spare Part Charges</strong></td>
                          <td align="right"><input name="job_totalPartsAmt" type="text" id="job_totalPartsAmt" style="text-align:right"  value="<%=chknumber2(job_totalPartsAmt)%>" size="10" maxlength="10" /></td>
                          <td>&nbsp;</td>
                        </tr>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="7" align="right"><strong>Labour Charges</strong></td>
                          <td align="right"><input name="job_totallabourAmt" type="text" id="job_totallabourAmt" style="text-align:right"  value="<%=chknumber2(job_totallabourAmt)%>" size="10" maxlength="10" /></td>
                          <td>
                              Tech Allowance
                           <%if job_overwty_allowance = "Yes" then%>                          
                            <input type="checkbox" name="job_overwty_allowance" id="job_overwty_allowance" checked  />
                          <%else%> 
                            <input type="checkbox" name="job_overwty_allowance" id="job_overwty_allowance" value="Yes" />
                          <%end if%>
                          </td>
                        </tr>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="7" align="right"><strong>Other Charges</strong></td>
                          <td align="right"><input name="job_totaltransportAmt" type="text" id="job_totaltransportAmt" style="text-align:right"  value="<%=chknumber2(job_totaltransportAmt)%>" size="10" maxlength="10" /></td>
                          <td align="center">
                          <%if (job_inv_no="" or isnull(job_inv_no)) then %>
                          <input type="submit" name="button11" id="button9" value="Save" />
                          <%end if%></td>
                        </tr>
                         </form>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="7" align="right"><strong>Total</strong><div class="total1"> </div></td>
                          <td><div align="right" class="total1"> <%=chknumber2(job_totalAmt)%> </div></td>
                          <td>&nbsp;</td>
                        </tr>
                      </table></td>
                    </tr>
                    
                    <tr>
                      <td align="left" bgcolor="#FFFFFF" scope="col">
          <%if (job_status="Done" or job_status="Posted") and (job_inv_no="" or isnull(job_inv_no)) then %>
          <input type="button" name="SubmitJob5" id="SubmitJob5" value="Generate Invoice" onclick="javascript:confirmAction('<%=job_code%>','action.asp?type=generateinvoicejob&job_code=<%=job_code%>')" />
          <%end if%>
          Invoice No: <a href="rm_invoice_new.asp?inv_no=<%=job_inv_no%>" target="_blank"><strong><%=(job_inv_no)%></strong></a> | Invoice Date: <strong><%=chkdate(job_inv_date)%></strong></td>
                      <td align="right" bgcolor="#FFFFFF" 
          scope="col">&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col">&nbsp;</td>
                    </tr>
                    
                    
                    
                    <tr>
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#CD6155">
                          <td colspan="8" bgcolor="#E8E8E8"><strong><font size="2">Related Jobs (Same Serial Number)</font></strong></td>
                          </tr>
                        <tr bgcolor="#475387">
                          <td><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Job No.</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Model</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>SN</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Service Date</strong></font></td>
                          <td width="10%" align="center"><font color="#FFFFFF"><strong> Status</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                          <td align="center"><strong><font color="#FFFFFF">City</font></strong></td>
                          </tr>
 <%		     
                i = 1
				
				if len(job_tech_SN) > 5 then 
				sql1 = "SELECT top 100 tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
				"tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
				"tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
				"tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
				"tbljob.job_Model, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
				"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
				"tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN,  " & _
				"tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
				"tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo,tbljob.job_wrty_photo2,tbljob.job_wrty_photo3, tbljob.job_hq_remark,  " & _
				"tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
				"tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
				"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbltechnician.tech_name, tbltechnician.tech_tel1 " & _
				"FROM tbljob inner join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
				"and tbljob.job_tech_SN = '" & job_tech_SN & "' and tbljob.job_code <> '" & job_code & "' "
				else
				sql1 = "SELECT tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
				"tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
				"tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
				"tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
				"tbljob.job_Model, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
				"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
				"tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN,  " & _
				"tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
				"tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo,tbljob.job_wrty_photo2,tbljob.job_wrty_photo3, tbljob.job_hq_remark,  " & _
				"tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
				"tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
				"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbltechnician.tech_name, tbltechnician.tech_tel1 " & _
				"FROM tbljob inner join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null " & _
				"and tbljob.job_id is null"
				end if
				
				
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>  
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="center"><strong><a href="rm_jobsheet.asp?job_code=<%=rs1("job_code")%>" target="_blank"><%=rs1("job_code")%></a></strong></td>
                          <td align="center"><%=rs1("job_tech_model")%></td>
                          <td align="center"><%=rs1("job_SN_no")%></td>
                          <td align="center"><%=chkdate(rs1("job_date"))%></td>
                          <td align="center"><%=rs1("job_status")%></td>
                          <td align="center"><%=rs1("job_tech_code") & "-" & rs1("tech_name") %></td>
                          <td align="center"><%=rs1("job_cust_city")%></td>
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
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col" class="auto-style1"></td>
                    </tr>
                    
                     <%if Request.Cookies("GAPS")("slevel") = "sc" then %>     
                    <form name="formorderparts_officeuse" id="formorderparts_officeuse" method="post" action="action.asp?type=editJob_officeuse&job_code=<%=job_code%>&#spareparts" >
                    <tr>
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV5">
                        <tbody>
                          <tr bgcolor="#E8E8E8">
                            <td bgcolor="#E8E8E8" scope="col"><strong><font size="2">Job Posting Status</font></strong></td>
                          </tr>
                          <tr >
                            <td nowrap="nowrap"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV6">
                              <tbody>
                                <tr >
                                  <td width="11%" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Issue Remark</strong></font></td>
                                  <td align="left"><select name="job_office_issueRemark" id="job_office_issueRemark">
                                    <option value="Posted" <%if job_office_issueRemark="Posted" then response.write " selected"%>>Posted</option>
                                    <option value="Write-Off" <%if job_office_issueRemark="Write-Off" then response.write " selected"%>>Write-Off</option>
                                    
                                    <%if job_rcn_no = "" or isnull(job_rcn_no) then  %>
                                    <option value="RCN" <%if job_office_issueRemark="RCN" then response.write " selected"%>>RCN</option>
                                    <%end if%>
                                    
                                    <option value="Cancel" <%if job_office_issueRemark="Cancel" then response.write " selected"%>>Cancel Job</option>
                                  </select></td>
                                  <td width="20%" align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Tax Invoice</strong></font></td>
                                  <td width="30%" align="left"><label for="select4"></label>
                                    <select name="job_office_taxinvoice" id="select4">
                                      <option value="No" <%if job_office_taxinvoice="No" then response.write " selected"%>>No</option>
                                      <option value="Yes" <%if job_office_taxinvoice="Yes" then response.write " selected"%>>Yes</option>
                                    </select></td>
                                  </tr>
                                <tr>
                                  <td bgcolor="#475387"><font color="#FFFFFF"><strong>Supervisor/ Manager</strong></font></td>
                                  <td align="left"><%=job_office_supervisor%></td>
                                  <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>RCN No</strong></font></td>
                                  <td align="left"><strong><a href="rm_rcn_new.asp?rcn_no=RCN10003" target="_blank"><%=job_rcn_no%></a></strong></td>
                                </tr>
                                <tr>
                                  <td bgcolor="#475387"><font color="#FFFFFF"><strong>Job Status</strong></font></td>
                                  <td><strong><%=job_status%></strong></td>
                                  <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>RCN Date</strong></font></td>
                                  <td align="left"><font color="#000000"><strong><%=chkdate(job_rcn_Date)%></strong></font></td>
                                </tr>
                                <tr>
                                  <td bgcolor="#475387">&nbsp;</td>
                                  <td></td>
                                  <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Invoice</strong></font></td>
                                  <td align="left"><a href="rm_invoice_new.asp?inv_no=<%=job_inv_no%>" target="_blank"><strong><%=(job_inv_no)%></strong></a></td>
                                </tr>
                                <tr>
                                  <td bgcolor="#475387">&nbsp;</td>
                                  <td>
                                  </td>
                                  <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                                  <td align="left"><strong><%=chkdate(job_inv_date)%></strong></td>
                                  </tr>
                                </tbody>
                            </table></td>
                            </tr>
                        </tbody>
                      </table></td>
                    </tr>
                    <tr>
                      <td bgcolor="#FFFFFF" 
          scope="col"><font color="#CC0000" size="2"><strong>Claim submitted by technician</strong></font>&nbsp;&nbsp; 
                         <strong> <%if job_submitforclaims="Yes" then response.write "YES" else response.write "NO"%></strong></td>
                      <td align="right" bgcolor="#FFFFFF" scope="col">
                          <%if job_status <> "Posted" then %>
                            <% if job_submitforclaims="Yes" then %>
                                <input type="submit" name="button6" id="button6" value="Update Job Status" />
                            <%else%>
                                <input type="submit" name="button6" id="button6" disabled value="Update Job Status" />
                          <%end if%>
                          <%end if%>
                       </td>
                    </tr>
                    <tr align="right">
                      <td colspan="2" bgcolor="#FFFFFF"></td>
                    </tr>
                    </form>
                  <%end if %>
                    
                 
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
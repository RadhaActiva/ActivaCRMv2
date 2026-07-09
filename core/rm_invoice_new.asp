<!-- #include file="header.asp" -->
<%
inv_cust_cnty_id = request.querystring("inv_cust_cnty_id")
inv_cust_name=request.querystring("inv_cust_name")
inv_cust_tel1=request.querystring("inv_cust_tel1")
inv_cust_tel2=request.querystring("inv_cust_tel2")
inv_cust_address=request.querystring("inv_cust_address")
inv_cust_postcode = request.form("inv_cust_postcode")

if inv_cust_postcode = "" then
    inv_cust_postcode = request.QueryString("inv_cust_postcode")
end if

set rs = server.CreateObject("adodb.recordset")
if request("inv_no") <> "" then	  
sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
		"inv_cust_city, inv_cust_city_id, inv_cust_cnty_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
		"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
		"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_remark, inv_job_code, inv_posteddate, inv_postedby,  " & _
		"inv_payment, inv_cnamount, inv_balance, inv_payment_type, inv_chequeno, inv_payment_remark, inv_dono, inv_dodate  " & _
		"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
		'response.write sql
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			inv_id = rs("inv_id") 
			inv_no = rs("inv_no") 
			inv_date = rs("inv_date") 
			inv_cust_code = rs("inv_cust_code")
			inv_cust_name = rs("inv_cust_name")
			inv_cust_address = rs("inv_cust_address") 
			inv_cust_postcode = rs("inv_cust_postcode") 
			inv_cust_state = rs("inv_cust_state") 
			inv_cust_state_id = rs("inv_cust_state_id") 
			inv_cust_city = rs("inv_cust_city") 
			inv_cust_city_id = rs("inv_cust_city_id") 
            inv_cust_cnty_id = rs("inv_cust_cnty_id")
			inv_cust_email = rs("inv_cust_email") 
			inv_cust_tel1 = rs("inv_cust_tel1") 
			inv_cust_tel2 = rs("inv_cust_tel2")  
			inv_createddate = rs("inv_createddate") 
			inv_createdby = rs("inv_createdby") 
			inv_approveddate = rs("inv_approveddate") 
			inv_approvedby = rs("inv_approvedby") 
			inv_status = rs("inv_status") 
			inv_job_code = rs("inv_job_code")
			inv_gstRate = rs("inv_gstRate")
			inv_gstAmt = rs("inv_gstAmt")
			inv_totalPartsAmt = rs("inv_totalPartsAmt")
			inv_totalAmt = rs("inv_totalAmt")
			inv_tech_code = rs("inv_tech_code")
			inv_posteddate = rs("inv_posteddate") 
			inv_postedby = rs("inv_postedby") 
			inv_remark = rs("inv_remark")
			
			inv_payment = rs("inv_payment") 
			inv_cnamount = rs("inv_cnamount") 
			inv_balance = rs("inv_balance") 
			inv_payment_type = rs("inv_payment_type") 
			inv_chequeno = rs("inv_chequeno") 
			inv_payment_remark = rs("inv_payment_remark")
			inv_dono = rs("inv_dono")
			inv_dodate = rs("inv_dodate")
			
		End If
		rs.Close
	  stype = "editinvoice"	
	  actionname = "Save" 
 else    
	  stype = "addinvoice"
	  actionname = "Save" 		
	  inv_date = date()	 
	  inv_status = "Open" 
	  inv_tech_code="walk-in"  	
end if


if inv_job_code <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_cnty_id, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
		"job_hq_category_code, job_hq_received_date, job_totalPartsAmt, job_totallabourAmt, job_totaltransportAmt, job_totalAmt, job_repair_date, job_return_tech_date,  " & _
		"job_office_issueRemark, job_office_supervisor, job_office_taxinvoice, job_rcn_no, job_rcn_Date, job_inv_no, job_inv_date, job_do_no, job_do_date " & _
	    "FROM tbljob WHERE job_code = '" & inv_job_code & "' "
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
 
 if inv_tech_code = "" or isnull(inv_tech_code) then 
    inv_tech_code = "walk-in"
 end if

if inv_cust_postcode <> "" and inv_cust_cnty_id = "129" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & inv_cust_postcode & "' "
		rs1.Open sql1,strconnect,0,1,&H0001   
		If Not rs1.EOF Then
             inv_cust_state_id = rs1("state_id") 'will auto populate state
             inv_cust_state=  rs1("state_name")
             inv_cust_city_id = rs1("city_id") 'will auto populate city
             inv_cust_city = rs1("post_office")    
        end if
    rs1.close
end if

custlabel=""
set rs3 = server.CreateObject("adodb.recordset")
sql3 = "SELECT cust_type FROM tblcustomer WHERE cust_code = '" & inv_cust_code & "' "
rs3.Open sql3,strconnect,0,1,&H0001
If Not rs3.EOF Then
    custlabel=rs3("cust_type")
    rs3.close
end if

%>

<script language="javascript">

    function getPostcode(p)
    {
    document.getElementById('inv_cust_postcode').value = p;
    document.formorder.submit();
    }

    function getCountrycode(c)
    {
    document.getElementById('inv_cust_cnty_id').value = c;
    document.formorder.submit();
    }

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.forminvoicedetail.action = orderlinks;
	document.forminvoicedetail.submit();
   }
}


function confirmInvoice(id, otype, orderlinks) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.forminvoice.action = orderlinks;
	document.forminvoice.submit();
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
          <td align="center">
              <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                     <table width="100%"><!--for UI Purpose-->
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Invoice</div></td>
                        <td width="23%" align="right" class="titleblue1">
                        <%if inv_status="Posted" then %>
                        <a href="rm_invoice_new_print.asp?inv_no=<%=inv_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a>&nbsp;&nbsp; <a href="/CRMone/pdfside/default.aspx?inv_no=<%=inv_no%>" target="_blank"><img src="images/pdf.png" height="35" width="30" alt="Pre-Printed Dot Matrix" border="0" style="border:0"/></a> 
                        <% end if%>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
                </tr>
             
                 <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td width="49%" rowspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer Information <font color="#006400">(<%=custlabel%>)</font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Code *</strong></font></td>
                        <td align="left">
                          <input name="inv_cust_code" type="text" id="inv_cust_code" style="background-color: #cccccc;" value="<%=inv_cust_code%>" maxlength="50" onclick="this.blur();" />
                          [<a href="javascript:popup('rm_invoice_new_customer.asp?searchitem=tblcustomer.cust_code&amp;searchvalue=<%=cust_code%>&formname=&fields=inv_cust_code','cb17','scrollbars=yes,resizable=yes,width=900,height=650')">Select Customer</a>]
                          </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
                        <td align="left"><input name="inv_cust_name" type="text" id="inv_cust_name" value="<%=inv_cust_name%>" size="50" maxlength="100" /></td>
                      </tr>
                        <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left">
                                    <select name="inv_cust_cnty_id" id="inv_cust_cnty_id" style="width:150px"  onblur="getCountrycode(this.value)">
                                    <option value="<%=inv_cust_cnty_id%>"></option>                                       
                                    <%
                                          sql = "SELECT cnty_name,cnty_id from tblcountry"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((inv_cust_cnty_id)) = cstr((rs1("cnty_id"))) then
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
                          <textarea name="inv_cust_address" cols="50" rows="5" id="inv_cust_address"><%=inv_cust_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left">
                             <%if inv_cust_cnty_id ="129" then' %> 
                                <input name="inv_cust_postcode" type="text" id="inv_cust_postcode" value="<%=inv_cust_postcode%>"  onblur="getPostcode(this.value)" size="10" maxlength="10" /></td>
                            <%else%> 
                                <input name="inv_cust_postcode" type="text" id="inv_cust_postcode" value="<%=inv_cust_postcode%>" size="10" maxlength="10" /></td>
                            <%end if%>                        
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left">
                         <%if inv_cust_cnty_id ="129" then' %> 
                            <input name="inv_cust_state" type="text" id="inv_cust_state" value="<%=inv_cust_state%>" size="30" maxlength="50" />
                            <input name="inv_cust_state_id" type="hidden" id="inv_cust_state_id" value="<%=inv_cust_state_id%>" size="30" "readonly" maxlength="50" />
                        <%else%> 
                            <input name="inv_cust_state" type="text" id="inv_cust_state" value="<%=inv_cust_state%>" size="30" style="background-color: #cccccc;" readonly maxlength="50" />
                            <input name="inv_cust_state_id" type="hidden" id="inv_cust_state_id" value="<%=inv_cust_state_id%>" size="30" "readonly" maxlength="50" />
                        <%end if%> 
                            </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                             <%if job_cust_cnty_id ="129" then' %> 
                             <input name="inv_cust_city" type="hidden" id="inv_cust_city" value="<%=inv_cust_city%>" size="30" maxlength="50" />
                             <select name="inv_cust_city_id" id="inv_cust_city_id" style="width:150px">
                             <option value="<%=inv_cust_city_id%>"></option>
                                    <%
                                          sql = "SELECT distinct city_id, post_office FROM tblpostcode where postcode = '" & job_cust_postcode & "'"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((inv_cust_city_id)) = cstr((rs1("city_id"))) then
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
                                    <input name="inv_cust_city_id" type="text" id="inv_cust_city_id" value="<%=inv_cust_city_id%>" readonly size="6" maxlength="50" /> 
                                    <input name="inv_cust_city" type="text" id="inv_cust_city" value="<%=inv_cust_city%>" size="30" readonly maxlength="50" />
                            <%end if%>
                            </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="inv_cust_email" type="text" id="inv_cust_email" value="<%=inv_cust_email%>" size="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><label for="inv_cust_tel1"></label>
                          <input name="inv_cust_tel1" type="text" id="inv_cust_tel1" value="<%=inv_cust_tel1%>" size="30" maxlength="50" />
                          e.g 0121234657</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="inv_cust_tel2" type="text" id="inv_cust_tel2" value="<%=inv_cust_tel2%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td valign="top"><strong>
                          <textarea name="inv_remark" cols="50" rows="3" id="inv_remark"><%=inv_remark%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                        <td valign="top"><%=inv_createdby%> @ 
                          <%=chkdatetime(inv_createddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Submitted By</strong></font></td>
                        <td valign="top"><%=inv_approvedby%> @ <%=chkdatetime(inv_approveddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Posted By</strong></font></td>
                        <td valign="top"><%=inv_postedby%> @ 
                          <%=chkdatetime(inv_posteddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Invoice Information</font></strong></td>
                      </tr>
                      
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Invoice  
                          No.</strong></font></td>
                        <td align="left"><strong><%=inv_no%>
                          <input type="hidden" name="inv_no" id="inv_no" value="<%=inv_no%>" />
                        </strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Invoice  
                          Date*</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(inv_date)%></strong></font></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Invoice  
                          Status.</strong></font></td>
                        <td colspan="3" align="left"><strong><%=inv_status%></strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician/Salesman</strong></font></td>
                        <td colspan="3" align="left"><select name="inv_tech_code" id="inv_tech_code">
                          <option value=""></option>
                          <%			
				sql = "SELECT tech_id, tech_code, tech_name FROM tbltechnician where tech_status = 'Y' "	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (inv_tech_code) = (rs("tech_code")) then
					  response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                        </select></td>
                        </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>DO No</strong></font></td>
                        <td align="left"><strong><%=inv_dono%></strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>DO  
                          Date*</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(inv_dodate)%></strong></font></td>
                      </tr>
                    </tbody>
                  </table>
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <br /></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV4">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Job Sheet Information</font></strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job  
                          No.<br />
                          <font size="1">(System Generate) </font></strong></font></td>
                        <td align="left"><strong><%=job_code%>
                          <input type="hidden" name="job_code" id="job_code" value="<%=job_code%>" />
                        </strong></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job  
                          Date*</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(job_date)%></strong></font></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status*</strong></font></td>
                        <td align="left"><strong><%=job_status%></strong></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Purchase  Date</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(job_purchase_date)%></strong></font></td>
                      </tr>
                      <tr align="left" >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
                        <td><%=job_onlineWrtyNo%></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                        <td><strong><%=job_onlineWrtyStatus%></strong></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type*</strong></font></td>
                        <td align="left"><%=job_type%></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>S/N</strong></font></td>
                        <td align="left"><%=job_SN_no%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Model*</strong></font></td>
                        <td colspan="3" align="left"><%=job_Model%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
                        <td colspan="3" align="left"><%=job_faulty_reason_cs%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Faulty Desc*</strong></font></td>
                        <td colspan="3" align="left"><strong><%=job_faulty_desc%></strong></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Reported By *</strong></font></td>
                        <td colspan="3"><%=job_reportedby%></td>
                        </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Appointment Date*</strong></font></td>
                        <td><font color="#000000"><strong><%=chkdate(job_appointment_date)%></strong></font></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>App Time </strong></font></td>
                        <td><%=job_appointment_time%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician*</strong></font></td>
                        <td colspan="3"><%=job_tech_code%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Appointment Remark<br />
                        </strong></font></td>
                        <td colspan="3"><strong><%=job_appointment_remark%></strong></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                
              
                <tr>
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><%if inv_status="Open" then %><input type="submit" name="button" id="button" value="<%=actionname%>" />
                  <%end if%></td>
              </tr>
                </form>
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                
                <%if inv_no <> "" then %>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td align="center"><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Spare Part 
                            Code</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Description</strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Unit Price (RM)</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Discount </strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Total 
                            Amt (RCP)</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Action</strong></font></td>
                        </tr>
                        <%

if request("invd_id") <> "" then
		sql = "SELECT invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_desc, invd_unitcost, invd_qty, invd_discountamt, " & _
		      "invd_discounttype, invd_netcost, invd_subtotal " & _
	          "FROM tblinvoice_detail where invd_id = '" & request("invd_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   invd_id = rs("invd_id")
		   invd_inv_no = rs("invd_inv_no")
		   invd_job_code = rs("invd_job_code")
		   invd_partcode = rs("invd_partcode")
		   invd_desc = rs("invd_desc")
		   invd_unitcost = rs("invd_unitcost")
		   invd_qty = rs("invd_qty")
		   invd_discountamt = rs("invd_discountamt")
		   invd_discounttype = rs("invd_discounttype")  
		   invd_netcost = rs("invd_netcost")   
		   invd_subtotal = rs("invd_subtotal")
        end if
		rs.close
		sbutton = "Update"
		stype="editInvoiceDetail"	
else
		sbutton = "Add"
		stype="addInvoiceDetail"
		invd_qty = "1"	
		invd_unitcost = "0.00"	
		invd_discountamt = "0.00"
		invd_netcost = "0.00"	
		invd_subtotal = "0.00"	
end if

%>
                         <%if inv_status="Open" then %>
                        <form name="forminvoicedetail" id="forminvoicedetail" method="post" action="rm_jobsheet.asp#spareparts" >
                          <tr>
                            <td bgcolor="#666666"><font color="#FFFFFF"><strong><a name="spareparts" id="spareparts"></a></strong></font></td>
                            <td align="left" bgcolor="#666666"><input name="invd_partcode" type="text" id="invd_partcode" value="<%=invd_partcode%>" maxlength="50" />
                              [<a href="javascript:popup('rm_invoice_new_spareparts.asp?searchitem=md_code&amp;job_code=<%=job_code%>&searchvalue=<%=cust_code%>&formname=forminvoicedetail&fieldname=invd_partcode&inv_tech_code=<%=inv_tech_code%>','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                            <td align="left" bgcolor="#666666"><font color="#FFFFFF"> </font>
                              <label for="invd_desc"></label>
                              <input name="invd_desc" type="text" id="invd_desc" value="<%=invd_desc%>" size="30" maxlength="100" /></td>
                            <td align="right" bgcolor="#666666"><font color="#FFFFFF">
                              <input type="hidden" name="inv_no" id="inv_no" value="<%=inv_no%>" />
                              <input type="hidden" name="invd_id" id="invd_id" value="<%=invd_id%>" />
                              <input name="invd_unitcost" type="text" id="invd_unitcost" style="text-align:right; background-color: #cccccc;" onkeydown="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);" value="<%=invd_unitcost%>" size="5" maxlength="10" />
                            </font></td>
                            <td align="right" bgcolor="#666666"><input name="invd_qty" type="text" id="invd_qty" style="text-align:right" onkeydown="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);" value="<%=invd_qty%>" size="5" maxlength="5" /></td>
                            <td align="right" nowrap="nowrap" bgcolor="#666666"><font color="#FFFFFF">
                              <input name="invd_discountamt" type="text" id="invd_discountamt" value="<%=invd_discountamt%>" size="5" onkeydown="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);" style="text-align:right" />
                              <select name="invd_discounttype" id="invd_discounttype" onchange="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.invd_unitcost.value, document.forminvoicedetail.invd_qty.value, document.forminvoicedetail.invd_discountamt.value, document.forminvoicedetail.invd_discounttype.value, document.forminvoicedetail.invd_subtotal);">
                                <option value="%" <%if invd_discounttype = "%" then response.write " selected"%>>%</option>
                                <option value="RM" <%if invd_discounttype = "RM" then response.write " selected"%>>RM</option>
                              </select>
                            </font></td>
                            <td align="right" bgcolor="#666666"><input name="invd_subtotal" type="text" id="invd_subtotal" style="text-align:right; background-color: #cccccc;" onfocus="this.blur();" value="<%=invd_subtotal%>" size="10" maxlength="10" /></td>
                            <td align="center" bgcolor="#666666"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=request("invd_id")%>','action.asp?type=<%=stype%>','<%=invd_subtotal%>');" /></td>
                          </tr>
                        </form>
                        <%end if%>
                        <%				i = 1
				sql1 = "SELECT invd_id, invd_inv_no, invd_job_code, invd_partcode, invd_desc, invd_unitcost, invd_qty, invd_discountamt, " & _
				       "invd_discounttype, invd_netcost, invd_subtotal	FROM tblinvoice_detail where invd_inv_no = '" & inv_no & "' order by invd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="left"><%=rs1("invd_partcode")%></td>
                          <td align="left"><%=rs1("invd_desc")%></td>
                          <td align="right"><%=chknumber2(rs1("invd_unitcost"))%></td>
                          <td align="right"><%=rs1("invd_qty")%></td>
                          <td align="right" nowrap="nowrap">- <%=chknumber2(rs1("invd_discountamt"))%> <%=rs1("invd_discounttype")%></td>
                          <td align="right"><%=chknumber2(rs1("invd_subtotal"))%></td>
                          <td align="center" nowrap="nowrap">
						     <%if inv_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_invoice_new.asp?invd_id=<%=rs1("invd_id")%>&amp;inv_no=<%=inv_no%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("invd_partcode")%>','action.asp?type=delInvoiceDetail&amp;invd_id=<%=rs1("invd_id")%>&amp;inv_no=<%=inv_no%>')" />
                            <%end if%></td>
                        </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>                        
                          <tr bgcolor="#EAEAEA">
                            <td height="25" colspan="6" align="right"><strong>Total, RM</strong></td>
                            <td align="right"><%=chknumber2(inv_totalAmt)%></td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr bgcolor="#EAEAEA">
                            <td height="25" colspan="6" align="right"><strong>CN, RM </strong></td>
                            <td align="right">-<%=chknumber2(inv_cnamount)%></td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr bgcolor="#EAEAEA">
                            <td height="25" colspan="6" align="right"><strong>Payment, RM </strong></td>
                            <td align="right"><strong>-<%=chknumber2(inv_payment)%></strong></td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr bgcolor="#EAEAEA">
                            <td height="25" colspan="6" align="right"><strong>Balance, RM</strong></td>
                            <td align="right"><strong><%=chknumber2(inv_balance)%></strong></td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr bgcolor="#EAEAEA">
                            <td height="25" colspan="8" align="left">**GST <%=inv_gstRate%>% Inclusive, 
                              GST Amount: RM <%=chknumber2(inv_gstAmt)%></td>
                          </tr>
                      </table></td>

                    </tr>
                    <form name="forminvoice" id="forminvoice" method="post" action="action.asp?type=submitInvoice&inv_no=<%=inv_no%>&#spareparts" >
                    <tr>
                     <td width="55%" rowspan="2" align="left" valign="top" bgcolor="#FFFFFF" scope="col"> <!--<strong>Email</strong>
                        <input name="emailto_DO" type="text" id="emailto_DO" value="<%=inv_cust_email%>" size="50" maxlength="150" />
                        <input type="button" name="Submit523" value="Email Invoice" style="{width:200px}" onclick="javascript:popup('rm_invoice_new_email.asp?inv_no=<%=inv_no%>&amp;emailto=' + forminvoice.emailto_DO.value + '&amp;email_remark=' + forminvoice.email_remark.value,'cb17','scrollbars=yes,resizable=yes,width=600px,height=600px')" />
                        <br />
                        <strong><font color="#000000">Email Message: </font></strong>:
                        <input name="email_remark" type="text" id="email_remark" size="55" maxlength="200" />
                        <br />
                        <br />
                        <br />
                        <br />-->
                        <%if inv_status<>"Posted" and inv_status<>"Open" and (request.Cookies("GAPS")("sloginid") = "davidhui" or request.Cookies("GAPS")("sloginid")="ERICLOH") then %>
                        <input type="button" name="CancelJob2" id="CancelJob2" value="Revert Invoice " onclick="javascript:confirmAction('<%=inv_no%>','action.asp?type=RevertInvoice&amp;inv_no=<%=inv_no%>')" />
                        <%end if%></td>
                      <td width="45%" align="right" bgcolor="#FFFFFF" 
          scope="col"><%if inv_status="Open" then %>
                        <input type="button" name="SubmitJob2" id="SubmitJob2" value="Submit Invoice " onclick="javascript:confirmAction('<%=inv_no%>','action.asp?type=submitInvoice&amp;inv_no=<%=inv_no%>')" />
                        <%end if%>
                        <br />
                        <%if inv_status="Submitted" then %>
                        <input type="button" name="PostedInvoice" id="PostedInvoice" value="Posted Invoice & Generate DO " onclick="javascript:confirmAction('<%=inv_no%>','action.asp?type=PostedInvoice&amp;inv_no=<%=inv_no%>')" />
                        <%end if%></td>
                    </tr>
                    <tr>
                      <td align="right" bgcolor="#FFFFFF" scope="col">
                      
                      <%if inv_status="Posted" then %>
                      <table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV5">
                        <tbody>
                          <tr bgcolor="#E8E8E8">
                            <td colspan="2" scope="col"><strong>Payment Information</strong></td>
                            </tr>
                          <tr >
                            <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Payment Amount</strong></font></td>
                            <td align="left"><strong>
                              <input name="inv_payment" type="text" id="inv_payment" value="<%=inv_balance%>" size="20" maxlength="20" />
                              </strong></td>
                            </tr>
                          <tr >
                            <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF">Payment Type</font></strong></td>
                            <td align="left">
                              <select name="inv_payment_type" id="inv_payment_type">
                                <option value="Cash" <%if inv_payment_type="Cash" then response.write " selected"%>>Cash</option>
                                <option value="Cheque" <%if inv_payment_type="Cheque" then response.write " selected"%>>Cheque</option>
                                <option value="Contra" <%if inv_payment_type="Contra" then response.write " selected"%>>Contra</option>
                                <option value="BankIn" <%if inv_payment_type="BankIn" then response.write " selected"%>>BankIn</option>
                                <option value="CreditCard" <%if inv_payment_type="CreditCard" then response.write " selected"%>>CreditCard</option>
                                </select></td>
                            </tr>
                          <tr >
                            <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cheque Number / Reference</strong></font></td>
                            <td align="left"><strong>
                              <input name="inv_chequeno" type="text" id="inv_chequeno" size="20" maxlength="20" />
                              </strong></td>
                            </tr>
                          <tr >
                            <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Payment Remark</strong></font></td>
                            <td align="left"><strong>
                              <textarea name="inv_payment_remark" cols="40" rows="3" id="inv_payment_remark"></textarea>
                              </strong></td>
                            </tr>
                          <tr >
                            <td colspan="2" align="right" nowrap="nowrap" bgcolor="#CD6155"><input type="button" name="SubmitJob" id="SubmitJob" value="Update Payment &amp; Generate Receipt" onclick="javascript:confirmInvoice('<%=inv_no%>','Confirm','action.asp?type=paymentInvoice&inv_no=<%=inv_no%>')" /></td>
                            </tr>
                          </tbody>
                      </table>
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
                
                <!-- CN Information Section - MOVED INSIDE the main table -->
      
            <table width="100%">
     
                <tr>
                  <td colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" bgcolor="#FFFFFF"><strong><font size="2">CN Information</font></strong></td>
                </tr>
                <tr>
                  <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>CN No.</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>CN Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> Part Code</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Customer</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Mobile</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer City</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Status</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Done Date</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted Date</span></strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Qty</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Amt</strong></font></td>
                    </tr>
                    <% 
j = 1
sql = "SELECT tblcn.cn_id, tblcn.cn_no, tblcn.cn_status, tblcn.cn_date, tblcn.cn_inv_no, tblcn.cn_inv_date, " & _
		"tblcn.cn_cust_code, tblcn.cn_cust_name, tblcn.cn_cust_address, " & _
		"tblcn.cn_cust_postcode, tblcn.cn_cust_state, tblcn.cn_cust_state_id, tblcn.cn_cust_city, tblcn.cn_cust_city_id, " & _
		"tblcn.cn_cust_email, tblcn.cn_cust_tel1, " & _
		"tblcn.cn_cust_tel2, tblcn.cn_createddate, tblcn.cn_createdby, tblcn.cn_job_code, tblcn.cn_do_no, tblcn.cn_invoice_no, " & _
		"tblcn.cn_totalqty, tblcn.cn_totalPartsAmt, " & _
		"tblcn.cn_remark, tblcn.cn_labourAmt, tblcn.cn_transportAmt, tblcn.cn_gstAmt, tblcn.cn_totalAmt, tblcn.cn_emailsent, " & _
		"tblcn.cn_emailsentdate, tblcn.cn_returnedby, " & _
		"tblcn.cn_returneddate, tblcn.cn_submittedby, tblcn.cn_submitteddate, tblcn.cn_doneby, tblcn.cn_donedate, " & _
		"tblcn.cn_postedby, tblcn.cn_posteddate, tblcn.cn_cancelledby, " & _
		"tblcn.cn_cancelleddate, tblcn_detail.cnd_partcode, tblcn_detail.cnd_qty, tblcn_detail.cnd_subtotal " & _
		"FROM tblcn inner join tblcn_detail on tblcn.cn_no = tblcn_detail.cnd_cn_no " & _
		"where tblcn.cn_id is not null " & _
		"and tblcn.cn_inv_no='" & inv_no & "' and cn_status='Posted'  order by tblcn.cn_id desc"
	  
'response.write sql	  
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if rs.eof then 
   response.write "No CN issued for this invoice."
end if

while not rs.eof 
%>
                    <tr>
                      <td height="40"><%=j%></td>
                      <td nowrap="nowrap"><strong><font color="#0000FF"><a href="rm_cn_new.asp?cn_no=<%=rs("cn_no")%>"><%=rs("cn_no")%></a></font></strong></td>
                      <td nowrap="nowrap"><%=chkdate(rs("cn_date"))%></td>
                      <td nowrap="nowrap"><%=rs("cn_inv_no")%></td>
                      <td><%=(rs("cnd_partcode"))%></td>
                      <td><%=(rs("cn_cust_name"))%></td>
                      <td><%=rs("cn_cust_tel1")%></td>
                      <td><%=rs("cn_cust_city")%></td>
                      <td><%=(rs("cn_status"))%></td>
                      <td><%=chkdate(rs("cn_donedate"))%></td>
                      <td><%=rs("cn_postedby")%> @<%=chkdate(rs("cn_posteddate"))%></td>
                      <td align="center"><%=rs("cnd_qty")%></td>
                      <td align="center"><%=chknumber2(rs("cnd_subtotal"))%></td>
                    </tr>
                    <%
cn_totalqty = cn_totalqty + rs("cnd_qty")
cn_totalAmt = cn_totalAmt + rs("cnd_subtotal")
j = j + 1
rs.MoveNext
wend
rs.Close
Set rs = Nothing
%>
                    <tr>
                      <td colspan="11" align="right"><strong> Total </strong></td>
                      <td align="center"><strong> <%=cn_totalqty%></strong></td>
                      <td align="center"><strong> <%=chknumber2(cn_totalAmt)%></strong></td>
                    </tr>
                  </table></td>
                </tr>
                
                <!-- Receipt Information Section - MOVED INSIDE the main table -->
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font size="2">Receipt Information</font></strong></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt No.</span></strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt Date</span></strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Invoice No</strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer</span></strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Customer Mobile</span></strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Payment Type</span></strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Payment Refer</strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Posted By</span></strong></font></td>
                      <td height="25" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt Status</span></strong></font></td>
                      <td height="25" align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>Receipt Amt</span></strong></font></td>
                    </tr>
                    <% 
j = 1
sql = "SELECT receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
	  "receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
	  "receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
	  "receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
	  "FROM tblreceipt where receipt_inv_no='" & inv_no & "' and receipt_status='Posted'  order by receipt_id desc"
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if rs.eof then 
   response.write "No Receipt issued for this invoice."
end if

while not rs.eof 
%>
                    <tr>
                      <td height="25"><%=j%></td>
                      <td height="25" nowrap="nowrap"><strong><font color="#0000FF"><a href="rm_receipt_new.asp?receipt_no=<%=rs("receipt_no")%>&receipt_inv_no=<%=rs("receipt_inv_no")%>"><%=rs("receipt_no")%></a></font></strong></td>
                      <td height="25" nowrap="nowrap"><%=chkdate(rs("receipt_date"))%></td>
                      <td height="25" nowrap="nowrap"><%=rs("receipt_inv_no")%></td>
                      <td height="25"><%=(rs("receipt_cust_name"))%></td>
                      <td height="25"><%=rs("receipt_cust_tel1")%></td>
                      <td height="25"><%=rs("receipt_paymenttype")%></td>
                      <td height="25"><%=rs("receipt_remark")%></td>
                      <td height="25"><%=(rs("receipt_createdby"))%></td>
                      <td height="25"><%=(rs("receipt_status"))%></td>
                      <td height="25" align="center"><%=chknumber2(rs("receipt_totalpayment"))%></td>
                    </tr>
                    <%
receipt_totalpayment = receipt_totalpayment + rs("receipt_totalpayment")
j = j + 1
rs.MoveNext
wend
rs.Close
Set rs = Nothing
%>
                    <tr>
                      <td colspan="10" align="right"><strong> Total </strong></td>
                      <td align="center"><strong> <%=chknumber2(receipt_totalpayment)%></strong></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr> 
              </table>
          </td>
        </tr>
<!-- #include file="footer.asp" -->
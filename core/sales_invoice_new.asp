<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("inv_no") <> "" then	  
sql = "SELECT inv_id, inv_no, inv_date, inv_cust_code, inv_cust_name, inv_cust_address, inv_cust_postcode, inv_cust_state, inv_cust_state_id, " & _
		"inv_cust_city, inv_cust_city_id, inv_cust_email, inv_cust_tel1, inv_cust_tel2, inv_createddate, inv_createdby, inv_tech_code,  " & _
		"inv_totalqty, inv_totalPartsAmt, inv_labourAmt, inv_transportAmt, inv_gstAmt, inv_gstRate, inv_gstCode, inv_totalAmt, inv_emailsent,  " & _
		"inv_emailsentdate, inv_status, inv_approvedby, inv_approveddate, inv_job_code  " & _
		"FROM tblinvoice WHERE inv_no = '" & request("inv_no") & "' "
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
			inv_cust_email = rs("inv_cust_email") 
			inv_cust_tel1 = rs("inv_cust_tel1") 
			inv_cust_tel2 = rs("inv_cust_tel2")  
			inv_createddate = rs("inv_createddate") 
			inv_createdby = rs("inv_createdby") 
			inv_approveddate = rs("inv_approveddate") 
			inv_approvedby = rs("inv_approvedby") 
			inv_status = rs("inv_status") 
			inv_job_code = rs("inv_job_code")
			
			inv_gstAmt = rs("inv_gstAmt")
			inv_totalPartsAmt = rs("inv_totalPartsAmt")
			inv_totalAmt = rs("inv_totalAmt")
		End If
		rs.Close
	  stype = "editinvoice"	
	  actionname = "Save" 
 else    
	  stype = "addinvoice"
	  actionname = "Save" 		
	  inv_date = date()	 
	  inv_status = "Open"   	
end if


if inv_job_code <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_email, job_cust_tel1, " & _
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
                        <td width="77%" class="titleblue1"><div align="left">Invoice</div></td>
                        <td width="23%" align="right" class="titleblue1"><a href="rm_invoice_new_print.asp?inv_no=<%=inv_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print | Email this page" border="0" style="border:0"/></a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
                </tr>
<tr>
                <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                  <tbody>
                    <tr>
                      <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer  
                        Information </font></strong></td>
                    </tr>
                    <tr>
                      <td width="22%" align="left" bgcolor="#333333"><font color="#FFFFFF"><strong>Cust Code *</strong></font></td>
                      <td align="left"><%=inv_cust_code%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
                      <td align="left"><%=inv_cust_name%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                      <td align="left"><%=inv_cust_address%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                      <td align="left"><%=inv_cust_postcode%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>State*</strong></font></td>
                      <td align="left"><%=inv_cust_state%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>City*</strong></font></td>
                      <td align="left"><%=inv_cust_city%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Email </strong></font></td>
                      <td valign="top">inv_cust_email</td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                      <td valign="top"><%=inv_cust_tel1%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                      <td valign="top"><%=inv_cust_tel2%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                      <td valign="top"><strong><%=inv_remark%></strong></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                      <td valign="top"><%=inv_createdby%> @ <%=chkdatetime(inv_createddate)%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#333333"><font color="#FFFFFF"><strong>Submitted By</strong></font></td>
                      <td valign="top"><%=inv_approvedby%> @ <%=chkdatetime(inv_approveddate)%></td>
                    </tr>
                  </tbody>
                </table></td>
                <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                  <tbody>
                    <tr bgcolor="#E8E8E8">
                      <td colspan="4" scope="col"><strong><font size="2"> Invoice Information</font></strong></td>
                    </tr>
                    <tr >
                      <td nowrap="nowrap" bgcolor="#333333"><font color="#FFFFFF"><strong>Invoice  
                        No.</strong></font></td>
                      <td align="left"><strong><%=inv_no%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#333333"><font color="#FFFFFF"><strong>Invoice  
                        Date*</strong></font></td>
                      <td><font color="#000000"><strong><%=chkdate(inv_date)%></strong></font></td>
                    </tr>
                    <tr >
                      <td nowrap="nowrap" bgcolor="#333333"><font color="#FFFFFF"><strong>Invoice  
                        Status</strong></font></td>
                      <td align="left"><strong><%=inv_status%></strong></td>
                      <td align="left" nowrap="nowrap" bgcolor="#333333">&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                  </tbody>
                </table>
                  <br />
                  <br />
                  <br />
                  <br />
                  <br />
                  <br />
                  <table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV4">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Job Sheet Information</font></strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#333333"><font color="#FFFFFF"><strong>Job  
                          No.<br />
                          <font size="1">(System Generate) </font></strong></font></td>
                        <td align="left"><%=job_code%></td>
                        <td align="left" nowrap="nowrap" bgcolor="#333333"><font color="#FFFFFF"><strong>Job  
                          Date*</strong></font></td>
                        <td><font color="#000000"><%=chkdate(job_date)%></font></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#333333"><font color="#FFFFFF"><strong>Status*</strong></font></td>
                        <td align="left"><%=job_status%></td>
                        <td align="left" bgcolor="#333333"><font color="#FFFFFF"><strong>Purchase  Date</strong></font></td>
                        <td><font color="#000000"><%=chkdate(job_purchase_date)%></font></td>
                      </tr>
                      <tr align="left" >
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
                        <td><%=job_onlineWrtyNo%></td>
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                        <td><%=job_onlineWrtyStatus%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Type*</strong></font></td>
                        <td align="left"><%=job_type%></td>
                        <td align="left" bgcolor="#333333"><font color="#FFFFFF"><strong>S/N</strong></font></td>
                        <td align="left"><%=job_SN_no%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Model*</strong></font></td>
                        <td colspan="3" align="left"><%=job_Model%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
                        <td colspan="3" align="left"><%=job_faulty_reason_cs%></td>
                      </tr>
                      <tr>
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Faulty Desc*</strong></font></td>
                        <td colspan="3" align="left"><%=job_faulty_desc%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Reported By *</strong></font></td>
                        <td colspan="2"><%=job_reportedby%></td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr >
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Appointment Date*</strong></font></td>
                        <td><font color="#000000"><%=chkdate(job_appointment_date)%></font></td>
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>App Time </strong></font></td>
                        <td><%=job_appointment_time%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Technician*</strong></font></td>
                        <td colspan="3"><%=job_tech_code%></td>
                      </tr>
                      <tr >
                        <td bgcolor="#333333"><font color="#FFFFFF"><strong>Appointment Remark<br />
                        </strong></font></td>
                        <td colspan="3"><strong><%=job_appointment_remark%></strong></td>
                      </tr>
                    </tbody>
                  </table></td>
              </tr>
                
              
              <tr>
                <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
              
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <tr valign="top">
                      <td width="45%" colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#333333">
                          <td><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Spare Part 
                            Code</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Description</strong></font></td>
                          <!--Added By sanjay on 23/Feb/2012-->
                          <td align="right"><font color="#FFFFFF"><strong>Unit Price (RM)</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Discount </strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Total 
                            Amt (RCP)</strong></font></td>
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
                          <td align="right">- <%=chknumber2(rs1("invd_unitcost")-rs1("invd_netcost"))%></td>
                          <td align="right"><%=chknumber2(rs1("invd_subtotal"))%></td>
                          </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                        
                          <tr bgcolor="#EAEAEA">
                            <td height="25" colspan="6" align="right"><strong>Total</strong></td>
                            <td align="right"><%=chknumber2(inv_totalPartsAmt)%></td>
                          </tr>
                          <tr bgcolor="#EAEAEA">
                            <td height="25" colspan="7" align="left">**GST 6% Inclusive, 
                              GST Amount: RM <%=chknumber2(inv_gstAmt)%></td>
                          </tr>
                      </table></td>
                    </tr>
                    <form name="forminvoice" id="forminvoice" method="post" action="action.asp?type=submitInvoice&inv_no=<%=inv_no%>&#spareparts" >
                  
                     </form>
                    <tr align="right">
                      <td colspan="2" bgcolor="#FFFFFF"></td>
                    </tr>
                     
                    <tr>
                      <td width="55%"></tbody></td>
                    </tr>
                  </table></td>
                </tr>
                 
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
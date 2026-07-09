<!-- #include file="database/datastore.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")
if request("job_code") <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_email, job_cust_tel1, " & _
		"job_cust_tel2, job_remark, job_createddate, job_createdby, job_JS_receiveddate, job_JS_receivedby, job_status, job_purchase_date, job_onlineWrtyNo, job_onlineWrtyStatus,  " & _
		"job_type, job_SN_no, job_Model, job_model_desc, job_faulty_reason_cs, job_faulty_desc, job_reportedby, job_appointment_date, job_appointment_time, job_tech_code, job_appointment_remark,  " & _
		"job_emailsentdate, job_emailsent, job_smssentdate, job_smssent, job_tech_type, job_tech_model, job_tech_model_desc, job_tech_tax_invoice, job_tech_SN, job_tech_faulty_code, job_tech_faulty_reason,  " & _
		"job_tech_faulty_action, job_tech_status, job_tech_product_collectdate, job_tech_service_date, job_tech_returntoCustDate, job_actual_wrty_status, job_wrty_photo, job_tech_logby, job_tech_logdate, job_hq_remark,  " & _
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
			job_submittedby = rs("job_submittedby")
			job_submitteddate = rs("job_submitteddate")
			job_doneby = rs("job_doneby")
			job_donedate = rs("job_donedate")
			job_postedby = rs("job_postedby")
			job_posteddate = rs("job_posteddate")
			job_cancelledby = rs("job_cancelledby")
			job_cancelleddate = rs("job_cancelleddate")
		End If	
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link href="inc/gaps_print.css" rel="stylesheet" type="text/css" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Riegen CRM</title>
</head>

<body>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="2" class="titleblue1"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <th width="19%" scope="row"><img src="images/Riegen.png" width="100" /></th>
             <td width="81%"><strong>Riegen Marketing Sdn Bhd</strong> <small>202401008163 (1554013-U)</small><br />
                B-3-A-18 & B-3A-19, Block Bougainvilla, 10 Boulevard, Lebuhraya Sprint, <br />
                PJU6A, 47400 Petaling Jaya, 
                <br />
                Selangor Darul Ehsan<br />
				  <a href="http://www.riegen.com.my/">www.riegen.com.my</a> | Tel:  03-77319139<br/></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td colspan="2" class="titleblue1"><hr /></td>
      </tr>
      <tr>
        <td width="77%" class="titleblue1"><div align="left">Job Sheet</div></td>
        <td width="23%" align="right" class="titleblue1">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
  </tr>
  <tr>
    <td width="45%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
      <tbody>
        <tr>
          <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer  
            Information </font></strong></td>
        </tr>
        <tr>
          <td width="22%" align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Cust Code *</strong></font></td>
          <td align="left"><%=job_cust_code%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
          <td align="left"><%=job_cust_name%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Address *</strong></font></td>
          <td align="left"><%=job_cust_address%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
          <td align="left"><%=job_cust_postcode%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>State*</strong></font></td>
          <td align="left"><%=job_cust_state%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>City*</strong></font></td>
          <td align="left"><%=job_cust_city%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Email </strong></font></td>
          <td valign="top"><%=job_cust_email%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
          <td valign="top"><%=job_cust_tel1%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
          <td valign="top"><%=job_cust_tel2%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Remark</strong></font></td>
          <td valign="top"><%=job_remark%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
          <td valign="top"><%=job_createdby%> @ <%=chkdatetime(job_createddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>JS Received Acknowlege</strong></font></td>
          <td valign="top"><%=job_JS_receivedby%> @ <%=chkdatetime(job_JS_receiveddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Submitted by</strong></font></td>
          <td valign="top"><%=job_submittedby%> @ <%=chkdatetime(job_submitteddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Done by</strong></font></td>
          <td valign="top"><%=job_doneby%> @ <%=chkdatetime(job_donedate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Posted by</strong></font></td>
          <td valign="top"><%=job_postedby%> @ <%=chkdatetime(job_posteddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Cancelled  by</strong></font></td>
          <td valign="top"><%=job_cancelledby%> @ <%=chkdatetime(job_cancelleddate)%></td>
        </tr>
      </tbody>
    </table></td>
    <td width="55%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
      <tbody>
        <tr bgcolor="#E8E8E8">
          <td colspan="4" scope="col"><strong><font size="2"> Job Sheet Information</font></strong></td>
        </tr>
        <tr >
          <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Job  
            No.</strong></font></td>
          <td align="left"><strong><%=job_code%></strong></td>
          <td align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Job  
            Date*</strong></font></td>
          <td><font color="#000000"><strong><%=chkdate(job_date)%> </strong></font></td>
        </tr>
        <tr >
          <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Status*</strong></font></td>
          <td align="left"><strong><%=job_status%></strong></td>
          <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Purchase  Date</strong></font></td>
          <td><%=chkdate(job_purchase_date)%></td>
        </tr>
        <tr align="left" >
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Online Wrty No.</strong></font></td>
          <td><%=job_onlineWrtyNo%></td>
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
          <td><strong>
            <%if job_onlineWrtyStatus="Over" then %>
            <font color="#FF0000">Over</font>
            <%else%>
            <font color="#000000">Under</font>
            <%end if%>
          </strong></td>
        </tr>
        <tr>
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Type*</strong></font></td>
          <td align="left"><%=job_type%></td>
          <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>S/N</strong></font></td>
          <td align="left"><%=job_SN_no%></td>
        </tr>
        <tr>
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>item codel*</strong></font></td>
          <td colspan="3" align="left"><%=job_Model%></td>
        </tr>
        <tr>
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Model Description</strong></font></td>
          <td colspan="3" align="left"><%=job_Model_desc%></td>
        </tr>
        <tr>
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
          <td colspan="3" align="left"><%=job_faulty_reason_cs%></td>
        </tr>
        <tr>
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Faulty Desc*</strong></font></td>
          <td colspan="3" align="left"><strong><%=job_faulty_desc%> </strong></td>
        </tr>
        <tr >
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Reported By *</strong></font></td>
          <td colspan="2"><%=job_reportedby%></td>
          <td></td>
        </tr>
        <tr >
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Appointment Date*</strong></font></td>
          <td><font color="#000000"><strong><%=chkdate(job_appointment_date)%></strong></font></td>
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>App Time </strong></font></td>
          <td><%=job_appointment_time%></td>
        </tr>
        <tr >
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Technician*</strong></font></td>
          <td colspan="3"><%=job_tech_code%></td>
        </tr>
        <tr >
          <td bgcolor="#475387"><font color="#FFFFFF"><strong>Appointment Remark<br />
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
      <tr align="right">
        <td colspan="2" valign="top" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
          <tbody>
            <tr bgcolor="#E8E8E8">
              <td colspan="4" align="left" bgcolor="#E8E8E8" scope="col"><strong>Technical Findings</strong></td>
            </tr>
            <tr >
              <td width="12%" align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Type*</strong></font></td>
              <td width="38%" align="left"><%=job_tech_type%></td>
              <td width="29%" align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Actual Model Number*</strong></font></td>
              <td width="21%" align="left"><%=job_tech_model%></td>
            </tr>
            <tr >
              <td align="left" nowrap="nowrap" bgcolor="#475387">&nbsp;</td>
              <td align="left">&nbsp;</td>
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Model Desc</strong></font></td>
              <td align="left"><%=job_tech_model_desc%></td>
            </tr>
            <tr >
              <td align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Tax Invoice*</strong></font></td>
              <td align="left"><%=job_tech_tax_invoice%></td>
              <td align="left" bgcolor="#475387"><strong><font color="#FFFFFF">Service Status<strong>*</strong></font></strong></td>
              <td align="left"><%=job_tech_status%></td>
            </tr>
            <tr>
              <td align="left" bgcolor="#475387">&nbsp;</td>
              <td align="left">&nbsp;</td>
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Service  Date</strong></font></td>
              <td align="left"><font color="#000000"><strong><%=chkdate(job_tech_service_date)%></strong></font></td>
            </tr>
            <tr >
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>S/N*</strong></font></td>
              <td align="left"><%=job_tech_SN%></td>
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Product Collection Date</strong></font></td>
              <td align="left"><font color="#000000"><strong><%=chkdate(job_tech_product_collectdate)%></strong></font></td>
            </tr>
            <tr >
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Faulty Reason*</strong></font></td>
              <td align="left"><%=job_tech_faulty_reason%></td>
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Actual Warranty Status*</strong></font></td>
              <td align="left"><%=job_actual_wrty_status%></td>
            </tr>
            <tr >
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Repair Action / <br />
                Remark*<br />
              </strong>(150 Chars) <strong></strong></font></td>
              <td align="left"><strong><%=job_tech_faulty_action%> </strong></td>
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Warranty Photo<br />
                (Optional) </strong></font></td>
              <td align="left"><a href="shared/<%=job_wrty_photo%>" target="_blank"><img src="shared/<%=job_wrty_photo%>" alt="Click on to Pop-up" width="100" border="0" /></a></td>
            </tr>
          </tbody>
        </table></td>
      </tr>
      <tr>
        <td colspan="2" scope="col">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" 
          scope="col"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV8">
          <tbody>
            <tr bgcolor="#E8E8E8">
              <td colspan="4" bgcolor="#E8E8E8" scope="col"><strong>Office Remarks</strong></td>
            </tr>
            <tr >
              <td width="12%" rowspan="2" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Remark<br />
              </strong>(150 Chars) </font></td>
              <td width="38%" rowspan="2" align="left"><strong><%=job_hq_remark%></strong></td>
              <td width="29%" align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>HQ Received Date</strong></font></td>
              <td width="21%" align="left"><%=chkdate(job_hq_received_date)%></td>
            </tr>
            <tr >
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Repair Date</strong></font></td>
              <td align="left"><font color="#000000"><strong><%=chkdate(job_repair_date)%></td>
            </tr>
            <tr >
              <td nowrap="nowrap" bgcolor="#475387">&nbsp;</td>
              <td align="left">&nbsp;</td>
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Return to Technician Date</strong></font></td>
              <td align="left"><font color="#000000"><strong><%=chkdate(job_return_tech_date)%></strong></font></td>
            </tr>
            <tr >
              <td width="12%" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Category Code</strong></font></td>
              <td width="38%" align="left"><%=job_hq_category_code%></td>
              <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Return to Customer Date</strong></font></td>
              <td align="left"><font color="#000000"><strong><%=chkdate(job_tech_returntoCustDate)%></strong></font></td>
            </tr>
          </tbody>
        </table></td>
      </tr>
      <tr>
        <td colspan="2" 
          scope="col">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" bgcolor="#E8E8E8" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr>
            <td><strong><font size="2">Spare-Part &amp; Services Charges</font></strong></td>
          </tr>
        </table></td>
      </tr>
      <tr valign="top">
        <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="1" cellspacing="0" cellpadding="8">
          <tr bgcolor="#475387">
            <td><font color="#FFFFFF"><strong>No</strong></font></td>
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
          <%				i = 1
				sql1 = "SELECT jobp_id, job_code, jobp_partcode, jobp_desc, jobp_unitcost, jobp_discountamt, jobp_discounttype, jobp_netcost, jobp_qty, jobp_subtotal " & _
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
            <td align="right"><%=chknumber2(rs1("jobp_unitcost"))%></td>
            <td align="right"><%=rs1("jobp_qty")%></td>
            <td align="right">- <%=chknumber2(rs1("jobp_unitcost")-rs1("jobp_netcost"))%></td>
            <td align="right"><%=chknumber2(rs1("jobp_subtotal"))%></td>
            <td align="center"></td>
          </tr>
          <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
          <form name="formorderparts_total" id="formorderparts_total" method="post" action="action.asp?type=editJobDetailTotal&amp;job_code=<%=job_code%>&amp;#spareparts" >
            <tr bgcolor="#EAEAEA">
              <td height="25" colspan="6" align="right"><strong>Total Spare Part Charges</strong></td>
              <td align="right"><%=chknumber2(job_totalPartsAmt)%></td>
              <td>&nbsp;</td>
            </tr>
            <tr bgcolor="#EAEAEA">
              <td height="25" colspan="6" align="right"><strong>Labour Charges</strong></td>
              <td align="right"><%=chknumber2(job_totallabourAmt)%></td>
              <td>&nbsp;</td>
            </tr>
            <tr bgcolor="#EAEAEA">
              <td height="25" colspan="6" align="right"><strong>Transport Charges</strong></td>
              <td align="right"><%=chknumber2(job_totaltransportAmt)%></td>
              <td align="center"></td>
            </tr>
          </form>
          <tr bgcolor="#EAEAEA">
            <td height="25" colspan="6" align="right"><strong>Total</strong>
              <div class="total1"> </div></td>
            <td><div align="right" class="total1"> <%=chknumber2(job_totalAmt)%></div></td>
            <td><div align="center"></div></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td colspan="2" bgcolor="#FFFFFF" 
          scope="col">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
          <tr bgcolor="#475387">
            <td colspan="8" bgcolor="#E8E8E8"><strong><font size="2">Related Jobs</font></strong></td>
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
          <%				i = 1
				sql1 = "SELECT top 10 tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
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
				"tbljob.job_tech_SN = '" & job_tech_SN & "' and tbljob.job_code <> '" & job_code & "' "
				rs1.Open sql1,strconnect,3,3,&H0001
              
              
          while Not rs1.EOF
%>
          <tr>
            <td align="center"><%=i%>.</td>
            <td align="center"><strong><a href="rm_jobsheet.asp?job_code=<%=rs1("job_code")%>" target="_blank"><%=rs1("job_code")%></a></strong></td>
            <td align="center"><%=rs1("job_tech_model")%></td>
            <td align="center"><%=rs1("job_Model")%></td>
            <td align="center"><%=rs1("job_date")%></td>
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
          scope="col">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV5">
          <tbody>
            <tr bgcolor="#E8E8E8">
              <td bgcolor="#E8E8E8" scope="col"><strong><font size="2">Job Posting Status</font></strong></td>
            </tr>
            <tr >
              <td nowrap="nowrap"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV6">
                <tbody>
                  <tr >
                    <td width="11%" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Issue Remark</strong></font></td>
                    <td align="left"><%=job_office_issueRemark%></td>
                    <td width="20%" align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Tax Invoice</strong></font></td>
                    <td width="30%" align="left"><%=job_office_taxinvoice%></td>
                  </tr>
                  <tr>
                    <td bgcolor="#475387"><font color="#FFFFFF"><strong>Supervisor/ Manager</strong></font></td>
                    <td align="left"><%=job_office_supervisor%></td>
                    <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>RCN No</strong></font></td>
                    <td align="left"><strong><%=job_rcn_no%></strong></td>
                  </tr>
                  <tr>
                    <td bgcolor="#475387">&nbsp;</td>
                    <td></td>
                    <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>RCN Date</strong></font></td>
                    <td align="left"><font color="#000000"><strong><%=chkdate(job_rcn_Date)%></strong></font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#475387">&nbsp;</td>
                    <td></td>
                    <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Invoice</strong></font></td>
                    <td align="left"><%=job_inv_no%></td>
                  </tr>
                  <tr>
                    <td bgcolor="#475387">&nbsp;</td>
                    <td></td>
                    <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Invoice Date</strong></font></td>
                    <td align="left"><strong><%=chkdate(job_inv_date)%></strong></td>
                  </tr>
                </tbody>
              </table></td>
            </tr>
          </tbody>
        </table></td>
      </tr>
      <tr align="right">
        <td colspan="2" bgcolor="#FFFFFF"></td>
      </tr>
      <tr>
        <td></tbody></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF">**** Note: This Document is computer generated and no signature is required.  **** </td>
  </tr>
</table>
</body>
</html>
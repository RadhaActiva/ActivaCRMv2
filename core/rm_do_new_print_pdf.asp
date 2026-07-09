<!-- #include file="database/dbconnect.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")

if request("do_no") <> "" then	  
sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, do_cust_state, " & _
		"do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_remark, do_createddate, do_createdby, do_job_code,  " & _
		"do_tech_code, do_totalqty, do_labourAmt, do_transportAmt, do_totalAmt, do_emailsent, do_emailsentdate, do_deliveredby, do_delivereddate,  " & _
		"do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate, " & _
		"do_purchase_date, do_onlineWrtyNo, do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, " & _
		"do_appointment_date, do_appointment_time, do_appointment_remark " & _
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
			do_cust_email = rs("do_cust_email") 
			do_cust_tel1 = rs("do_cust_tel1") 
			do_cust_tel2 = rs("do_cust_tel2")  
			do_remark = rs("do_remark")
			do_createddate = rs("do_createddate") 
			do_createdby = rs("do_createdby") 
			do_tech_code = rs("do_tech_code") 
			do_totalqty = rs("do_totalqty") 
			do_labourAmt = rs("do_labourAmt")
			do_transportAmt = rs("do_transportAmt")
			do_totalAmt = rs("do_totalAmt")
			do_emailsent = rs("do_emailsent")
			do_emailsentdate = rs("do_emailsentdate")
			do_deliveredby = rs("do_deliveredby")
			do_delivereddate = rs("do_delivereddate")
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
end if


if do_job_code <> "" then	  
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_email, job_cust_tel1, " & _
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Riegen Marketing CRM</title>
<link href="inc/gaps_print.css" rel="stylesheet" type="text/css" />

</head>

<body>
<table width="680" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2" align="center" valign="top" bgcolor="#000000"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <th width="19%" scope="row"><img src="images/riegen.png" width="100" /></th>
             <td width="81%"><strong>Riegen Marketing Sdn Bhd</strong> <small>202401008163 (1554013-U)</small><br />
                B-3-A-18 & B-3A-19, Block Bougainvilla, 10 Boulevard, Lebuhraya Sprint, <br />
                PJU6A, 47400 Petaling Jaya, 
                <br />
                Selangor Darul Ehsan<br />
				  <a href="http://www.riegen.com.my/">www.riegen.com.my</a> | Tel:  03-77319139<br/></td>
          </tr>
        </table></td>      
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><hr /></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><div align="left">DELIVERY ORDER (DO)</div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td width="300" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
      <tbody>
        <tr>
          <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer  
            Information </font></strong></td>
        </tr>
        <tr>
          <td width="39%" align="left" bgcolor="#FFFFFF"><font color="#000000"><strong>Cust Code *</strong></font></td>
          <td width="61%" align="left" bgcolor="#FFFFFF"><%=do_cust_code%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Cust Name *</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_cust_name%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Address *</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_cust_address%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Postcode*</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_cust_postcode%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>State*</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_cust_state%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>City*</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_cust_city%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Email </strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_cust_email%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Tel. No. 1*</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_cust_tel1%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Tel. No. 2</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_cust_tel2%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Remark</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_remark%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Created by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_createdby%> @ <%=chkdatetime(do_createddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Delivered by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_deliveredby%> @ <%=chkdatetime(do_delivereddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Done by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_doneby%> @ <%=chkdatetime(do_donedate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Posted by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_postedby%> @ <%=chkdatetime(do_posteddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Cancelled  by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=do_cancelledby%> @ <%=chkdatetime(do_cancelleddate)%></td>
        </tr>
      </tbody>
    </table></td>
    <td valign="top" bgcolor="#FFFFFF"><table width="95%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
      <tbody>
        <tr bgcolor="#E8E8E8">
          <td colspan="2" scope="col"><strong><font size="2"> DO Information</font></strong></td>
        </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>DO  
            No.<br />
            <font size="1">(System Generate) </font></strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><strong><%=do_no%></strong></td>
          </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>DO  
            Date</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=chkdate(do_date)%></td>
          </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Status</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_status%></td>
          </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Purchase  Date</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><font color="#000000"><%=chkdate(do_purchase_date)%></font></td>
        </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Job Code</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><strong><%=do_job_code%></strong></td>
          </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Job Date</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=chkdate(job_date)%></td>
        </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Online Wrty No.</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_onlineWrtyNo%></td>
          </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Wrty Status</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><strong>
            <%if do_onlineWrtyStatus="Over" then %>
            <font color="#FF0000">Over</font>
            <%else%>
            <font color="#000000">Under</font>
            <%end if%>
          </strong></td>
          </tr>
        <tr>
          <td align="left" bgcolor="#FFFFFF"><font color="#000000"><strong>S/N</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_SN_no%></td>
        </tr>
        <tr>
          <td align="left" bgcolor="#FFFFFF"><font color="#000000"><strong>Type</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_type%></td>
        </tr>
        <tr>
          <td align="left" bgcolor="#FFFFFF"><font color="#000000"><strong>Model</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_Model%></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><font color="#000000"><strong>Model <br />
            Description</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_Model_desc%></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><font color="#000000"><strong>Appointment <br />
            Date</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=chkdate(do_appointment_date)%></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><font color="#000000"><strong>Technician</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_tech_code%></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><strong><font color="#000000">Invoice No.</font></strong></td>
          <td align="left" bgcolor="#FFFFFF"><%=do_inv_no%></td>
          </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><strong><font color="#000000">Invoice Date</font></strong></td>
          <td align="left" bgcolor="#FFFFFF"><%=chkdate(do_inv_date)%></td>
          </tr>
      </tbody>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" cellpadding="8" cellspacing="0">
      <tr bgcolor="#333333">
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>No</strong></font></td>
        <td align="left" bgcolor="#EAEAEA"><font color="#000000"><strong>Item Code</strong></font></td>
        <td align="left" bgcolor="#EAEAEA"><font color="#000000"><strong> Model Description</strong></font></td>
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Qty</strong></font></td>
      </tr>
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
        <td align="center" bgcolor="#FFFFFF"><%=i%>.</td>
        <td align="left" bgcolor="#FFFFFF"><%=rs1("dod_partcode")%></td>
        <td align="left" bgcolor="#FFFFFF"><%=rs1("dod_desc")%></td>
        <td align="center" bgcolor="#FFFFFF"><%=rs1("dod_qty")%></td>
      </tr>
      <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
      <tr bgcolor="#FFFFFF">
        <td height="25" colspan="3" align="right"><strong>Total</strong></td>
        <td align="center"><%=chknumber0(do_totalqty)%></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0">
      <tr>
        <td colspan="2" valign="top">&nbsp;</td>
        </tr>
      <tr>
        <td width="74%" valign="top"><font color="#000000"><strong>Terms and Condition</strong><br />
          <br />
          ALL CHEQUES SHOULD BE CROSSED AND MADE PAYABLE TO &quot;<strong>RIEGEN MARKETING SDN BHD</strong>&quot;, PAYMENT TERMS AS PER OUR ARRANGEMENT.<br />
OVERDUE INTEREST OF 1.5% PER MONTH WILL BE CHARGED FROM OVERDUE DATE.<br />
          </font></td>
        <td width="26%"><p> I/ We  hereby agreed and accepted the above<br />
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </p>
          <p>&nbsp;</p>
          <p> ................................................<br />
            Customer Name:<br />
            Company  Stamp (if any)<br />
            Position:<br />
            Date:&nbsp; </p></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF">**** Note: This DO is computer generated and no signature is required.  **** </td>
  </tr>
</table>
</body>
</html>
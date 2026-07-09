<!-- #include file="database/datastore.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Riegen CRM</title>
<link href="inc/gaps_print.css" rel="stylesheet" type="text/css" />

</head>

<body> 

<%

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
sql = "SELECT job_id, job_code, job_date, job_cust_code, job_cust_name, job_cust_address, job_cust_postcode, job_cust_state, job_cust_city, job_cust_email, job_cust_tel1, " & _
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


if job_onlineWrtyNo <> "" then	  
sql = "SELECT FROM dealername tblonlinewarranty WHERE warrantyno = '" & job_onlineWrtyNo & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			rcn_Dealer = rs("dealername") 
		End If
		rs.Close
end if

sql = "select cnty_name from tblcountry where cnty_id =" & rcn_cust_cnty_id	
rcn_country_name = selectid(sql)

%>

        <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                          <tr>
                            <th width="19%" scope="row"><img src="images/Riegen.png" width="100" /></th>
             <td width="81%" align="left"><strong>Riegen Marketing Sdn Bhd</strong> <small>202401008163 (1554013-U)</small><br />
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
                        <td width="77%" class="titleblue1"><div align="left">REPLACEMENT / CREDIT NOTE REQUEST FORM (RCN)</div></td>
                        <td width="23%" align="right" class="titleblue1">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              <tr>
                <td rowspan="2" valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV4">
                  <tbody>
                    <tr>
                      <td width="22%" align="left" valign="top" bgcolor="#475387"><strong><font color="#FFFFFF">To</font></strong></td>
                      <td align="left">Finance and Admin Department</td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Attention</strong></font></td>
                      <td align="left">Accountant / Assitant Accountant</td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>From</strong></font></td>
                      <td align="left">Service Department/Branch</td>
                    </tr>
                  </tbody>
                </table></td>
                <td valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV5">
                  <tbody>
                    <tr>
                      <td width="22%" align="left" valign="top" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>RCN  
                          No.</strong></font></td>
                      <td align="left"><%=rcn_no%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Approved By</strong></font></td>
                      <td align="left"><%=rcn_postedby%> @ <%=chkdatetime(rcn_posteddate)%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>RCN  
                        Date</strong></font></td>
                      <td align="left"><%=chkdate(rcn_date)%></td>
                    </tr>
                    <tr>
                      <td align="left" valign="top" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                      <td align="left"><%=rcn_createdby%> @ <%=chkdatetime(rcn_createddate)%></td>
                    </tr>
                  </tbody>
                </table></td>
              </tr>
              <tr>
                <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
              <tr>
                <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong>Kindly be advice the following items are not serviable. Please arrange replacement for the Dealer/Customer as per DELIVERY INSTRUCTIONS stated below:</strong></td>
              </tr>
              <tr>
                <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
              <tr>
                <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Dealer / Customer  
                        Information </font></strong></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Dealer</strong></font></td>
                        <td align="left"><%=rcn_Dealer%></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Cust Code </strong></font></td>
                        <td align="left"><%=rcn_cust_code%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Cust Name </strong></font></td>
                        <td align="left"><%=rcn_cust_name%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Address </strong></font></td>
                        <td align="left"><%=rcn_cust_address%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Postcode</strong></font></td>
                        <td align="left"><%=rcn_cust_postcode%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>State</strong></font></td>
                        <td align="left"><%=rcn_cust_state%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>City</strong></font></td>
                        <td align="left"><%=rcn_cust_city%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left" bgcolor="#FFFFFF"><%=rcn_country_name%></td>
                       </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><%=rcn_cust_email%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Tel. No. 1</strong></font></td>
                        <td valign="top"><%=rcn_cust_tel1%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><%=rcn_cust_tel2%></td>
                      </tr>
                    </tbody>
                </table></td>
                <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Additional Information</font></strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Job No.</strong></font></td>
                        <td align="left"><%=rcn_job_code%></td>
                        <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Posted Date</strong></font></td>
                        <td><%=chkdate(job_posteddate)%></td>
                      </tr>
                      <tr align="left" >
                        <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Status</strong></font></td>
                        <td><strong><%=rcn_status%></strong></td>
                        <td bgcolor="#475387"><font color="#FFFFFF"><strong>SN No.</strong></font></td>
                        <td><%=rcn_SN_no%></td>
                      </tr>
                      <tr align="left" >
                        <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Wrty Card No.</strong></font></td>
                        <td>
                        <strong><%=rcn_onlineWrtyNo%></strong></td>
                        <td bgcolor="#475387"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                        <td><strong><%=rcn_onlinewrtyStatus%></strong></td>
                      </tr>
                      <tr>
                        <td bgcolor="#475387"><font color="#FFFFFF"><strong>Model</strong></font></td>
                        <td align="left"><strong><%=rcn_modelcode%></strong></td>
                        <!--<td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Type</strong></font></td>-->
						<td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>RCN Type</strong></font></td>
                        <td align="left">
                        <strong><%=rcn_modeltype%></strong></td>
                      </tr>
                      <tr >
                        <td bgcolor="#475387"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                        <td colspan="3"><strong><%=rcn_tech_code%> - <%=tech_name%></strong></td>
                      </tr>
                      <tr >
                        <td bgcolor="#475387"><font color="#FFFFFF"><strong>Faulty Action</strong></font></td>
                        <td colspan="3"><strong><%=job_tech_faulty_reason%></strong></td>
                      </tr>
                      <tr >
                        <td bgcolor="#475387"><font color="#FFFFFF"><strong>Faulty Remark</strong></font></td>
                        <td colspan="3"><strong><%=job_tech_faulty_action%></strong></td>
                      </tr>
                    </tbody>
                </table></td>
              </tr>
              
              
              
             
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                  
                    <tr valign="top">
                      <td width="45%" colspan="3" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="1" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Model Code</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Model Description</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Job Sheet No</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>RCN Type</strong></font></td>
                          </tr>
                 
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
                          <td align="center"><%=rs1("rcnd_qty")%></td>
                          <td align="center"><%=rcn_job_code%></td>
                          <!--<td align="center">Replacement</td>-->
						  <td align="center"><%=rcn_modeltype%></td>
                          </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="6" align="left"><strong>Delivery Instruction / Remark: </strong><%=rcn_remark%></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr align="right">
                      <td colspan="3" bgcolor="#FFFFFF"></td>
                    </tr>
                    <tr>
                      <td width="55%"></tbody></td>
                    </tr>
                  </table></td>
                </tr>
             
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0">
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
              </table>

</body>
</html>
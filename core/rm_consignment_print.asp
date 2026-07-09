<!-- #include file="database/datastore.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("sp_no") <> "" then	  
sql = "SELECT sp_id, sp_no, sp_tech_code, sp_tech_name, sp_tech_address, sp_tech_postcode, sp_tech_state, sp_tech_city, sp_tech_email, sp_tech_tel1, " & _
		"sp_tech_tel2, sp_tech_carplateno, sp_createddate, sp_createdby, sp_date, sp_status, sp_submitteddate, sp_submittedby, sp_approveddate, sp_approvedby, " & _
		"sp_deliverydate, sp_deliveryby, sp_trackingno, sp_couriercompany, sp_confirmedreceiveddate,  sp_confirmedreceivedby, " & _
		"sp_rejecteddate, sp_rejectedremark, sp_remark, sp_totalqty, sp_labourAmt, sp_transportAmt, sp_gstAmt, sp_gstRate, sp_gstCode, sp_totalAmt,  " & _
		"sp_emailsent, sp_emailsentdate, sp_logby, sp_logdate, sp_posteddate, sp_postedby, sp_canceldate, sp_cancelby " & _
		"FROM tblsparepartrequest WHERE sp_no = '" & request("sp_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sp_id = rs("sp_id") 
			sp_no = rs("sp_no") 
			sp_tech_code = rs("sp_tech_code")
			sp_tech_name = rs("sp_tech_name")
			sp_tech_address = rs("sp_tech_address")
			sp_tech_postcode = rs("sp_tech_postcode") 
			sp_tech_state = rs("sp_tech_state") 
			sp_tech_city = rs("sp_tech_city") 
			sp_tech_email = rs("sp_tech_email") 
			sp_tech_tel1 = rs("sp_tech_tel1") 
			sp_tech_tel2 = rs("sp_tech_tel2") 
			sp_tech_carplateno = rs("sp_tech_carplateno") 
			sp_createddate = rs("sp_createddate")  
			sp_createdby = rs("sp_createdby")
			sp_date = rs("sp_date") 
			sp_status = rs("sp_status") 
			sp_submitteddate = rs("sp_submitteddate") 
			sp_submittedby = rs("sp_submittedby") 
			sp_approveddate = rs("sp_approveddate") 
			sp_approvedby = rs("sp_approvedby") 
			sp_deliverydate = rs("sp_deliverydate") 
			sp_deliveryby = rs("sp_deliveryby") 
			sp_trackingno = rs("sp_trackingno") 
			sp_couriercompany = rs("sp_couriercompany")
			sp_confirmedreceiveddate = rs("sp_confirmedreceiveddate") 
			sp_confirmedreceivedby = rs("sp_confirmedreceivedby") 
			sp_rejecteddate = rs("sp_rejecteddate") 
			sp_rejectedremark = rs("sp_rejectedremark") 
			sp_remark = rs("sp_remark") 
			sp_totalqty = rs("sp_totalqty") 
			sp_labourAmt = rs("sp_labourAmt") 
			sp_transportAmt = rs("sp_transportAmt") 
			sp_gstAmt = rs("sp_gstAmt") 
			sp_gstRate = rs("sp_gstRate") 
			sp_gstCode = rs("sp_gstCode") 
			sp_totalAmt = rs("sp_totalAmt") 
			sp_emailsent = rs("sp_emailsent") 
			sp_emailsentdate = rs("sp_emailsentdate") 
			sp_logby = rs("sp_logby") 
			sp_logdate = rs("sp_logdate") 
			sp_postedby = rs("sp_postedby") 
			sp_posteddate = rs("sp_posteddate") 
			sp_cancelby = rs("sp_cancelby") 
			sp_canceldate = rs("sp_canceldate") 
		End If
		rs.Close
	  stype = "editSparepartsRequest"	
	  actionname = "Save" 
 else    
	  stype = "AddSparepartsRequest"
	  actionname = "Save" 		
	  sp_date = date()	 
	  sp_status = "Open"   	
	  

      sql = "SELECT tech_id, tech_code, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
            "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id " & _
	        "FROM tbltechnician WHERE tech_code = '" & request.Cookies("GAPS")("job_tech_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sp_tech_code = rs("tech_code")
			sp_tech_name = rs("tech_name")
			sp_tech_icno = rs("tech_icno")
			sp_tech_address = rs("tech_address")
			sp_tech_postcode = rs("tech_postcode") 
			sp_tech_state = rs("tech_state") 
			sp_tech_state_id = rs("tech_state_id") 
			sp_tech_city = rs("tech_city") 
			sp_tech_city_id = rs("tech_city_id") 
			sp_tech_email = rs("tech_email") 
			sp_tech_tel1 = rs("tech_tel1") 
			sp_tech_tel2 = rs("tech_tel2") 
			sp_tech_carmodel = rs("tech_carmodel") 
			sp_tech_carplateno = rs("tech_carplateno") 
			sp_tech_carcolour = rs("tech_carcolour") 
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
<table width="650" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <th width="19%" scope="row"><img src="images/Riegen.png" width="100" /></th>
            <td width="81%"><p><strong>Riegen Marketing Sdn Bhd</strong> <small>202401008163 (1554013-U)</small><br />
              B-3-A-18 & B-3A-19, Block Bougainvilla, 0 Boulevard, Lebuhraya Sprint, PJU6A, 47400 Petaling Jaya<br />
              <a href="http://www.riegen.com.my/">www.riegen.com.my</a> | Tel:  03-77319139 (Service Hotline)&nbsp;
              <br />
            </p></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><hr />
         </td>
      </tr>
      <tr>
        <td class="titleblue1"><div align="left"><strong>Spare Part Request</strong></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
  </tr>
  <tr>
    <td width="300" rowspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
      <tbody>
        <tr>
          <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Technician  
            Information </font></strong></td>
        </tr>
        <tr>
          <td width="22%" align="left" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Techician Code*</strong></font></td>
          <td align="left"><label for="textfield4"></label>
            <strong><%=sp_tech_code%>
              <input name="sp_tech_code" type="hidden" id="sp_tech_code" value="<%=sp_tech_code%>" />
            </strong></td>
        </tr>
        <tr>
          <td align="left" valign="top" nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Techician Name*</strong></font></td>
          <td align="left"><%=sp_tech_name%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Address *</strong></font></td>
          <td align="left"><%=sp_tech_address%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
          <td align="left"><%=sp_tech_postcode%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>State*</strong></font></td>
          <td align="left"><%=sp_tech_state%></td>
          </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><strong><font color="#FFFFFF">City*</font></strong></td>
          <td align="left"><%=sp_tech_city%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Email </strong></font></td>
          <td align="left"><label for="select8"><%=sp_tech_email%></label></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
          <td valign="top"><%=sp_tech_tel1%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
          <td valign="top"><label for="textfield3"><%=sp_tech_tel2%></label></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#475387"><strong><font color="#FFFFFF">Car Plate No</font></strong></td>
          <td valign="top"><%=sp_tech_carplateno%></td>
        </tr>
      </tbody>
    </table></td>
    <td width="350" valign="top" bgcolor="#FFFFFF">
      <table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
        <tbody>
          <tr bgcolor="#E8E8E8">
            <td colspan="2" scope="col"><strong><font size="2"> Spare Part Information</font></strong></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Request  
              No.</strong></font></td>
            <td align="left" nowrap="nowrap"><strong><%=sp_no%></strong></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Requested Date</strong></font></td>
            <td nowrap="nowrap"><label for="sp_couriercompany"><%=chkdate(sp_date)%></label></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Status</strong></font></td>
            <td align="left" nowrap="nowrap"><strong><%=sp_status%></strong></td>
          </tr>
          <tr align="left" >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Courrier Track No.</strong></font></td>
            <td nowrap="nowrap"><label for="textfield5"><%=sp_trackingno%></label></td>
          </tr>
          <tr>
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Courrier Company</strong></font></td>
            <td align="left" nowrap="nowrap"><%=sp_couriercompany%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Remark</strong></font></td>
            <td nowrap="nowrap"><%=sp_remark%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Created by</strong></font></td>
            <td nowrap="nowrap"><%=sp_createdby%> @ <%=chkdatetime(sp_createddate)%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Submitted Date </strong></font></td>
            <td nowrap="nowrap"><%=sp_submittedby%> @ <%=chkdate(sp_submitteddate)%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Approved  Date</strong></font></td>
            <td nowrap="nowrap"><%=sp_approvedby%> @ <%=chkdate(sp_approveddate)%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Delivery  Date</strong></font></td>
            <td nowrap="nowrap"><%=sp_deliveryby%> @ <%=chkdate(sp_deliverydate)%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Confirmed Received Date</strong></font></td>
            <td nowrap="nowrap"><%=sp_confirmedreceivedby%> @ <%=chkdate(sp_confirmedreceiveddate)%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Posted Date </strong></font></td>
            <td nowrap="nowrap"><%=sp_postedby%> @ <%=chkdate(sp_posteddate)%></td>
          </tr>
          <tr >
            <td nowrap="nowrap" bgcolor="#475387"><font color="#FFFFFF"><strong>Rejected  Date</strong></font></td>
            <td nowrap="nowrap"><%=sp_rejectedby%> @ <%=chkdate(sp_rejecteddate)%></td>
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
    <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
      <tbody>
      </tbody>
      <tr valign="top">
        <td width="45%" colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="1" cellpadding="8" cellspacing="0">
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
          <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
          <tr>
            <td height="25" align="left"><table width="100%" border="0" cellspacing="0" cellpadding="8">
              <tr bgcolor="#475387">
                <td><font color="#FFFFFF"><strong>No.</strong></font></td>
                <td><font color="#FFFFFF"><strong>Spare Part 
                  Code</strong></font></td>
                <td align="center"><font color="#FFFFFF"><strong> Description</strong></font></td>
                <td width="10%" align="center"><font color="#FFFFFF"><strong>My Stock Qty</strong></font></td>
                <!--Added By sanjay on 23/Feb/2012-->
                <td align="right"><font color="#FFFFFF"><strong>Unit Price (RM)</strong></font></td>
                <td width="10%" align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                <td align="right"><font color="#FFFFFF"><strong>Total 
                  Amt (RCP)</strong></font></td>
                </tr>
             
             
              <%				i = 1
				sql1 = "SELECT spd_id, spd_sp_no, spd_tech_code, spd_partcode, spd_currentstock, spd_desc, spd_unitcost, spd_qty, spd_subtotal " & _
	                   "FROM tblsparepartrequest_detail where spd_sp_no = '" & sp_no & "' order by spd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
              <tr>
                <td align="center"><%=i%>.</td>
                <td align="left"><%=rs1("spd_partcode")%></td>
                <td align="center"><%=rs1("spd_desc")%></td>
                <td align="center"><%=rs1("spd_currentstock")%></td>
                <td align="right"><%=rs1("spd_unitcost")%></td>
                <td align="right"><%=rs1("spd_qty")%></td>
                <td align="right"><%=chknumber2(rs1("spd_subtotal"))%></td>
                </tr>
              <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
              <tr bgcolor="#EAEAEA">
                <td height="25" colspan="5" align="right"><strong>Total</strong>
                  <div class="total1"> </div></td>
                <td height="25" align="right"><%=sp_totalqty%></td>
                <td align="right"><strong><%=chknumber2(sp_totalAmt)%></strong></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td align="left"><table width="100%" border="0" cellpadding="3" cellspacing="0">
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
          </table></td>
      </tr>
      <form name="forminvoice" id="forminvoice" method="post" action="action.asp?type=submitInvoice&amp;inv_no=<%=inv_no%>&amp;#spareparts" >
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
</table>
</body>
</html>
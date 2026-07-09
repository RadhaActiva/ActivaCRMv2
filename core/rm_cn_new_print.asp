<!-- #include file="database/datastore.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")

if request("cn_no") <> "" then	  
sql = "SELECT cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
	  "cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
	  "cn_job_code, cn_do_no, cn_invoice_no, cn_totalqty, cn_totalPartsAmt, cn_remark, cn_labourAmt, cn_transportAmt, cn_gstAmt, cn_totalAmt,  " & _
	  "cn_emailsent, cn_emailsentdate, cn_returnedby, cn_returneddate, cn_submittedby, cn_submitteddate, cn_doneby, cn_donedate, cn_postedby,  " & _
	  "cn_posteddate, cn_cancelledby, cn_cancelleddate " & _
	  "FROM tblcn WHERE cn_no = '" & request("cn_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			cn_id = rs("cn_id") 
			cn_no = rs("cn_no") 
			cn_status = rs("cn_status")
			cn_date = rs("cn_date") 
			cn_inv_no = rs("cn_inv_no")  
			cn_inv_date = rs("cn_inv_date")  
			cn_cust_code = rs("cn_cust_code") 
			cn_cust_name = rs("cn_cust_name")  
			cn_cust_address = rs("cn_cust_address")  
			cn_cust_postcode = rs("cn_cust_postcode")  
			cn_cust_state = rs("cn_cust_state")  
			cn_cust_state_id = rs("cn_cust_state_id")  
			cn_cust_city = rs("cn_cust_city")  
			cn_cust_city_id = rs("cn_cust_city_id") 
			cn_cust_email = rs("cn_cust_email")  
			cn_cust_tel1 = rs("cn_cust_tel1")  
			cn_cust_tel2 = rs("cn_cust_tel2")  
			cn_createddate = rs("cn_createddate")  
			cn_createdby = rs("cn_createdby")  
			cn_job_code = rs("cn_job_code")  
			cn_do_no = rs("cn_do_no")  
			cn_invoice_no = rs("cn_invoice_no")  
			cn_totalqty = rs("cn_totalqty")  
			cn_totalPartsAmt = rs("cn_totalPartsAmt")  
			cn_remark = rs("cn_remark")  
			cn_labourAmt = rs("cn_labourAmt")  
			cn_transportAmt = rs("cn_transportAmt")  
			cn_gstAmt = rs("cn_gstAmt")  
			cn_totalAmt = rs("cn_totalAmt")  
			cn_emailsent = rs("cn_emailsent")  
			cn_emailsentdate = rs("cn_emailsentdate")  
			cn_returnedby = rs("cn_returnedby")  
			cn_returneddate = rs("cn_returneddate")  
			cn_submittedby = rs("cn_submittedby")  
			cn_submitteddate = rs("cn_submitteddate")  
			cn_doneby = rs("cn_doneby")  
			cn_donedate = rs("cn_donedate")  
			cn_postedby = rs("cn_postedby")  
			cn_posteddate = rs("cn_posteddate")  
			cn_cancelledby = rs("cn_cancelledby")  
			cn_cancelleddate = rs("cn_cancelleddate")  
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
            <th width="19%" scope="row"><img src="images/Riegen.png" width="100" /></th>
               <td width="81%"><strong>Riegen Marketing Sdn Bhd</strong> 202401008163 (1554013-U)<br />
                B-3-A-18 & B-3A-19, Block Bougainvilla, 10 Boulevard, Lebuhraya Sprint, <br />
                PJU6A, 47400 Petaling Jaya, 
                <br />
                Selangor Darul Ehsan<br />
				  <a href="http://www.riegen.com.my/">www.riegen.com.my</a> | Tel:  03-77319139<br/></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><hr /></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><div align="left">CREDIT NOTE (CN)</div></td>
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
          <td width="61%" align="left" bgcolor="#FFFFFF"><%=cn_cust_code%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Cust Name *</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=cn_cust_name%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Address *</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=cn_cust_address%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Postcode*</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=cn_cust_postcode%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>State*</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=cn_cust_state%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>City*</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=cn_cust_city%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Email </strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_cust_email%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Tel. No. 1*</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_cust_tel1%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Tel. No. 2</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_cust_tel2%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Remark</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_remark%></td>
        </tr>
      </tbody>
    </table></td>
    <td align="right" valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
      <tbody>
        <tr bgcolor="#E8E8E8">
          <td colspan="2" scope="col"><strong><font size="2"> CN Information</font></strong></td>
        </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>CN  
            No.</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><strong><%=cn_no%></strong></td>
          </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>CN  
            Date</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=chkdate(cn_date)%></td>
          </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Status</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=cn_status%></td>
          </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Job Code</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><strong><%=cn_job_code%></strong></td>
        </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>DO No</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=(cn_do_no)%></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><strong><font color="#000000">Invoice No.</font></strong></td>
          <td align="left" bgcolor="#FFFFFF"><%=cn_inv_no%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Created by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_createdby%> @ <%=chkdatetime(cn_createddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Returned by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_returnedby%> @ <%=chkdatetime(cn_returneddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Done by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_doneby%> @ <%=chkdatetime(cn_donedate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Posted by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_postedby%> @ <%=chkdatetime(cn_posteddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Cancelled  by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=cn_cancelledby%> @ <%=chkdatetime(cn_cancelleddate)%></td>
        </tr>
      </tbody>
    </table></td>
  </tr>
 
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="8" cellspacing="0">
      <tr bgcolor="#475387">
        <td><font color="#FFFFFF"><strong>No</strong></font></td>
        <td align="left"><font color="#FFFFFF"><strong>Spare Part 
          Code</strong></font></td>
        <td align="left"><font color="#FFFFFF"><strong> Description</strong></font></td>
        <!--Added By sanjay on 23/Feb/2012-->
        <td align="right"><font color="#FFFFFF"><strong>Unit Price<br />
          (RM)</strong></font></td>
        <td align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
        <td align="right"><font color="#FFFFFF"><strong>Discount </strong></font></td>
        <td align="right"><font color="#FFFFFF"><strong>Total 
          Amt<br />
          (RCP)</strong></font></td>
      </tr>
      <%				i = 1
				sql1 = "SELECT cnd_id, cnd_cn_no, cnd_job_code, cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt, " & _
				       "cnd_discounttype, cnd_netcost, cnd_subtotal	FROM tblcn_detail where cnd_cn_no = '" & cn_no & "' order by cnd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
      <tr>
        <td align="center"><%=i%>.</td>
        <td align="left"><%=rs1("cnd_partcode")%></td>
        <td align="left"><%=rs1("cnd_desc")%></td>
        <td align="right"><%=chknumber2(rs1("cnd_unitcost"))%></td>
        <td align="right"><%=rs1("cnd_qty")%></td>
        <td align="right">- <%=chknumber2(rs1("cnd_unitcost")-rs1("cnd_netcost"))%></td>
        <td align="right"><%=chknumber2(rs1("cnd_subtotal"))%></td>
      </tr>
      <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
      <tr bgcolor="#EAEAEA">
        <td height="25" colspan="3" align="left">**GST 6% Inclusive, 
          GST Amount: (SR) RM <%=chknumber2(cn_gstAmt)%></td>
        <td height="25" colspan="3" align="right"><strong>Total, RM</strong></td>
        <td align="right"><%=chknumber2(cn_totalPartsAmt)%></td>
      </tr>
      <tr>
        <td height="25" colspan="7" align="left"><table width="100%" border="0" cellpadding="3" cellspacing="0">
          <tr>
            <td colspan="2" valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td width="74%" valign="top"><font color="#000000"><strong>Terms and Condition</strong><br />
              <br />
              ALL CHEQUES SHOULD BE CROSSED AND MADE PAYABLE TO &quot;<strong>RIEGEN MARKETING SDN BHD</strong>&quot;, PAYMENT TERMS AS PER OUR ARRANGEMENT.<br />
              OVERDUE INTEREST OF 1.5% PER MONTH WILL BE CHARGED FROM OVERDUE DATE.<br />
            </font></td>
            <td width="26%"><p> <strong>Approved by</strong><br />
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
              <p> <strong>................................................<br />
                Name:<br />
                Position:<br />
                Date:&nbsp; </strong></p></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
   <tr>
     <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
   </tr>
</table>
</body>
</html>
<!-- #include file="database/datastore.asp" -->

<%
set rs = server.CreateObject("adodb.recordset")
if request("receipt_no") <> "" then	  
sql = "SELECT receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
	"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
	"receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
	"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
	"FROM tblreceipt WHERE receipt_no = '" & request("receipt_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
     while Not rs.EOF
            receipt_job_code = receipt_job_code + " / " + rs("receipt_job_code") + " "  
            receipt_inv_no = receipt_inv_no + " / " + rs("receipt_inv_no") 
		    receipt_totalpayment = receipt_totalpayment + rs("receipt_totalpayment") 
    		receipt_id = rs("receipt_id") 
			receipt_no = rs("receipt_no") 
			receipt_status = rs("receipt_status")
			receipt_date = rs("receipt_date") 
			receipt_inv_date = rs("receipt_inv_date")  
			receipt_createddate = rs("receipt_createddate")  
			receipt_createdby = rs("receipt_createdby")  
			receipt_inv_date = rs("receipt_inv_date") 
			receipt_remark = rs("receipt_remark")  
			receipt_paymenttype = rs("receipt_paymenttype")  			
			receipt_emailsent = rs("receipt_emailsent")  
			receipt_emailsentdate = rs("receipt_emailsentdate")  
			receipt_cancelledby = rs("receipt_cancelledby")  
			receipt_cancelleddate = rs("receipt_cancelleddate") 
        rs.movenext
		wend
		rs.Close
           	
		receipt_job_code = replace(receipt_job_code,"/","",1,1)
        receipt_inv_no = replace(receipt_inv_no,"/","",1,1)
end if
 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Riegen CRM</title>
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
        <td bgcolor="#FFFFFF" class="titleblue1"><hr /></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><div align="left"><strong>RECEIPT</strong></div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
     <td align="right" valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
      <tbody>
        <tr bgcolor="#E8E8E8">
          <td colspan="2" scope="col"><strong><font size="2"> Receipt Information</font></strong></td>
        </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Receipt  
            No.</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><strong><%=receipt_no%></strong></td>
          </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Receipt  
            Date</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=chkdate(receipt_date)%></td>
          </tr>
        <tr >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Status</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><%=receipt_status%></td>
          </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><strong><font color="#000000">Invoice No.</font></strong></td>
          <td align="left" bgcolor="#FFFFFF"><strong><%=receipt_inv_no%></strong></td>
        </tr>
        <tr align="left" >
          <td align="left" nowrap="nowrap" bgcolor="#FFFFFF"><font color="#000000"><strong>Job Code</strong></font></td>
          <td align="left" bgcolor="#FFFFFF"><strong><%=receipt_job_code%></strong></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><strong><font color="#000000">Payment Amount</font></strong></td>
          <td align="left" bgcolor="#FFFFFF"><strong>RM <%=receipt_totalpayment%></strong></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><strong><font color="#000000">Payment Type</font></strong></td>
          <td align="left" bgcolor="#FFFFFF"><%=receipt_paymenttype%></td>
        </tr>
        <tr >
          <td align="left" bgcolor="#FFFFFF"><strong><font color="#000000">Payment Remark</font></strong></td>
          <td align="left" bgcolor="#FFFFFF"><%=receipt_remark%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Created by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=receipt_createdby%> @ <%=chkdatetime(receipt_createddate)%></td>
        </tr>
        <tr>
          <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Cancelled  by</strong></font></td>
          <td valign="top" bgcolor="#FFFFFF"><%=receipt_cancelledby%> @ <%=chkdatetime(receipt_cancelleddate)%></td>
        </tr>
      </tbody>
    </table></td>
  </tr>
 
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" valign="top" bgcolor="#FFFFFF"><font color="#000000"><strong>Terms and Condition</strong><br />
        <br />
ALL CHEQUES SHOULD BE CROSSED AND MADE PAYABLE TO &quot;<strong>RIEGEN MARKETING SDN BHD</strong>&quot;, PAYMENT TERMS AS PER OUR ARRANGEMENT. <br />
OVERDUE INTEREST OF 1.5% PER MONTH WILL BE CHARGED FROM OVERDUE DATE.<br />
<br />
**** Note: This Receipt is computer generated and no signature is required.  ****    </font></td>
  </tr>
</table>
</body>
</html>
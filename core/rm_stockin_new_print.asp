<!-- #include file="database/datastore.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")

if request("st_no") <> "" then	  
sql = "SELECT st_id, st_no, st_date, st_referenceno, st_status, st_fromwarehouse, st_towarehouse, st_remark, st_createddate, st_createdby, st_submitteddate, " & _
		"st_submittedby, st_approveddate, st_approveddate, st_approvedby, st_cancelleddate, st_cancelledby, st_totalqty, st_totalaAmt, st_emailsent, st_emailsentdate " & _
		"FROM tblstockin WHERE st_no = '" & request("st_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			st_id = rs("st_id") 
			st_no = rs("st_no") 
			st_date = rs("st_date")
			st_referenceno = rs("st_referenceno") 
			st_status = rs("st_status")
			st_fromwarehouse = rs("st_fromwarehouse") 
			st_towarehouse = rs("st_towarehouse")
			st_remark = rs("st_remark")
			st_createddate = rs("st_createddate") 
			st_createdby = rs("st_createdby") 
			st_submitteddate = rs("st_submitteddate") 
			st_submittedby = rs("st_submittedby") 
			st_approveddate = rs("st_approveddate") 
			st_approvedby = rs("st_approvedby") 
			st_cancelleddate = rs("st_cancelleddate") 
			st_cancelledby = rs("st_cancelledby") 
			st_totalqty = rs("st_totalqty") 
			st_totalaAmt = rs("st_totalaAmt") 
			st_emailsent = rs("st_emailsent")  
			st_emailsentdate = rs("st_emailsentdate")
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
    <td align="center" valign="top" bgcolor="#000000"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><table width="100%" border="0" cellspacing="0" cellpadding="3">
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
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><hr /></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><div align="left">STOCK-IN</div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
      <tr>
        <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Stock-In No</strong></font></td>
        <td nowrap="nowrap"><strong><%=st_no%></strong></td>
        <td height="22" nowrap="nowrap" bgcolor="#475387 "><strong><font color="#FFFFFF"><strong>Stock-In date</strong></font></strong></td>
        <td><%=chkdate(st_date)%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
        <td nowrap="nowrap"><%=st_referenceno%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Status</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=st_status%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Store</strong></font></strong></td>
        <td nowrap="nowrap"><%=st_towarehouse%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=st_remark%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
        <td nowrap="nowrap"><%=st_createdby%> @ <%=chkdatetime(st_createddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
        <td><%=st_approvedby%> @ <%=chkdatetime(st_approveddate)%></td>
      </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
        <td nowrap="nowrap"><%=st_submittedby%> @ <%=chkdatetime(st_submitteddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
        <td><%=st_cancelledby%> @ <%=chkdatetime(st_cancelleddate)%></td>
        </tr>
    </table></td>
  </tr>
 
  <tr>
    <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" cellpadding="8" cellspacing="0">
      <tr bgcolor="#333333">
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>No</strong></font></td>
        <td align="left" bgcolor="#EAEAEA"><font color="#000000"><strong>Item 
          Code</strong></font></td>
        <td align="left" bgcolor="#EAEAEA"><font color="#000000"><strong> Item Description</strong></font></td>
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Stock-In Qty</strong></font></td>
      </tr>
      <%				i = 1
				sql1 = "SELECT std_id, std_st_no, std_itm_code, std_itm_desc, std_unitcost, std_qty, std_subtotal, std_referid " & _
	                  "FROM tblstockin_detail where std_st_no = '" & st_no & "' order by std_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
      <tr>
        <td align="center" bgcolor="#FFFFFF"><%=i%>.</td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("std_itm_code")%></span></td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("std_itm_desc")%></span></td>
        <td align="center" bgcolor="#FFFFFF"><span class="tktTotals"><%=rs1("std_qty")%></span></td>
      </tr>
    <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
      <tr bgcolor="#FFFFFF">
        <td height="25" colspan="3" align="right"><strong>Total</strong></td>
        <td align="center"><strong><%=st_totalqty%></strong></td>
      </tr>
    </table></td>
  </tr>
   <tr>
     <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
   </tr>
   <tr>
    <td valign="top" bgcolor="#FFFFFF">**** Note: This Document is computer generated and no signature is required.  **** </td>
  </tr>
</table>
</body>
</html>
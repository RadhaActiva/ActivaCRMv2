<!-- #include file="database/datastore.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")

if request("so_no") <> "" then	  
sql = "SELECT so_id, so_no, so_date, so_referenceno, so_status, so_fromwarehouse, so_towarehouse, so_remark, so_createddate, so_createdby, so_submitteddate, " & _
		"so_submittedby, so_approveddate, so_approveddate, so_approvedby, so_cancelleddate, so_cancelledby, so_totalqty, so_totalaAmt, so_emailsent, so_emailsentdate " & _
		"FROM tblstockOut WHERE so_no = '" & request("so_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			so_id = rs("so_id") 
			so_no = rs("so_no") 
			so_date = rs("so_date")
			so_referenceno = rs("so_referenceno") 
			so_status = rs("so_status")
			so_fromwarehouse = rs("so_fromwarehouse") 
			so_towarehouse = rs("so_towarehouse")
			so_remark = rs("so_remark")
			so_createddate = rs("so_createddate") 
			so_createdby = rs("so_createdby") 
			so_submitteddate = rs("so_submitteddate") 
			so_submittedby = rs("so_submittedby") 
			so_approveddate = rs("so_approveddate") 
			so_approvedby = rs("so_approvedby") 
			so_cancelleddate = rs("so_cancelleddate") 
			so_cancelledby = rs("so_cancelledby") 
			so_totalqty = rs("so_totalqty") 
			so_totalaAmt = rs("so_totalaAmt") 
			so_emailsent = rs("so_emailsent")  
			so_emailsentdate = rs("so_emailsentdate")
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
            <th width="19%" scope="row"><img src="images/Riegen.png" width="100" /></th>
              <td width="81%"><strong>Riegen Marketing Sdn Bhd</strong><small>202401008163 (1554013-U)</small><br />
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
        <td bgcolor="#FFFFFF" class="titleblue1"><div align="left">STOCK-OUT</div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
      <tr>
        <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Stock-Out No</strong></font></td>
        <td nowrap="nowrap"><strong><%=so_no%></strong></td>
        <td height="22" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Stock-Out date</strong></font></strong></td>
        <td><%=chkdate(so_date)%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
        <td nowrap="nowrap"><%=so_referenceno%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Status</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=so_status%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Store</strong></font></strong></td>
        <td nowrap="nowrap"><%=so_fromwarehouse%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=so_remark%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
        <td nowrap="nowrap"><%=so_createdby%> @ <%=chkdatetime(so_createddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
        <td><%=so_approvedby%> @ <%=chkdatetime(so_approveddate)%></td>
      </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
        <td nowrap="nowrap"><%=so_submittedby%> @ <%=chkdatetime(so_submitteddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
        <td><%=so_cancelledby%> @ <%=chkdatetime(so_cancelleddate)%></td>
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
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Unit Price</strong></font></td>
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Stock-Out Qty</strong></font></td>
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Sub-Total</strong></font></td>
      </tr>
     <%				i = 1
				sql1 = "SELECT sod_id, sod_so_no, sod_itm_code, sod_itm_desc, sod_unitcost, sod_qty, sod_subtotal, sod_referid " & _
	                  "FROM tblstockOut_detail where sod_so_no = '" & so_no & "' order by sod_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
      <tr>
        <td align="center" bgcolor="#FFFFFF"><%=i%>.</td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("sod_itm_code")%></span></td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("sod_itm_desc")%></span></td>
        <td align="center" bgcolor="#FFFFFF"><span class="tktTotals"><%=chknumber2(rs1("sod_unitcost"))%></span></td>
        <td align="center" bgcolor="#FFFFFF"><span class="tktTotals"><%=rs1("sod_qty")%></span></td>
        <td align="center" bgcolor="#FFFFFF"><span class="tktTotals"><%=chknumber2(rs1("sod_subtotal"))%></span></td>
      </tr>
    <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
      <tr bgcolor="#FFFFFF">
        <td height="25" colspan="4" align="right"><strong>Total</strong></td>
        <td align="center"><strong><%=so_totalqty%></strong></td>
        <td align="center"><strong><%=chknumber2(so_totalaAmt)%></strong></td>
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
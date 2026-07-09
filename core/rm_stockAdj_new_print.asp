<!-- #include file="database/datastore.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")

if request("sj_no") <> "" then	  
sql = "SELECT sj_id, sj_no, sj_date, sj_referenceno, sj_status, sj_fromwarehouse, sj_towarehouse, sj_remark, sj_createddate, sj_createdby, sj_submitteddate, " & _
		"sj_submittedby, sj_approveddate, sj_approveddate, sj_approvedby, sj_cancelleddate, sj_cancelledby, sj_totalqty, sj_totalaAmt, sj_emailsent, sj_emailsentdate " & _
		"FROM tblstockadj WHERE sj_no = '" & request("sj_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sj_id = rs("sj_id") 
			sj_no = rs("sj_no") 
			sj_date = rs("sj_date")
			sj_referenceno = rs("sj_referenceno") 
			sj_status = rs("sj_status")
			sj_fromwarehouse = rs("sj_fromwarehouse") 
			sj_towarehouse = rs("sj_towarehouse")
			sj_remark = rs("sj_remark")
			sj_createddate = rs("sj_createddate") 
			sj_createdby = rs("sj_createdby") 
			sj_submitteddate = rs("sj_submitteddate") 
			sj_submittedby = rs("sj_submittedby") 
			sj_approveddate = rs("sj_approveddate") 
			sj_approvedby = rs("sj_approvedby") 
			sj_cancelleddate = rs("sj_cancelleddate") 
			sj_cancelledby = rs("sj_cancelledby") 
			sj_totalqty = rs("sj_totalqty") 
			sj_totalaAmt = rs("sj_totalaAmt") 
			sj_emailsent = rs("sj_emailsent")  
			sj_emailsentdate = rs("sj_emailsentdate")
		End If
		rs.Close
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
    <td align="center" valign="top" bgcolor="#000000"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td bgcolor="#FFFFFF" class="titleblue1"><table width="100%" border="0" cellspacing="0" cellpadding="3">
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
        <td bgcolor="#FFFFFF" class="titleblue1"><div align="left">STOCK-ADJ</div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
      <tr>
        <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong>Stock-ADJ No</strong></font></td>
        <td nowrap="nowrap"><strong><%=sj_no%></strong></td>
        <td height="22" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Stock-ADJ date</strong></font></strong></td>
        <td><%=chkdate(sj_date)%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
        <td nowrap="nowrap"><%=sj_referenceno%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Status</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=sj_status%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong> Store</strong></font></strong></td>
        <td nowrap="nowrap"><%=sj_fromwarehouse%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=sj_remark%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
        <td nowrap="nowrap"><%=sj_createdby%> @ <%=chkdatetime(sj_createddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
        <td><%=sj_approvedby%> @ <%=chkdatetime(sj_approveddate)%></td>
      </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
        <td nowrap="nowrap"><%=sj_submittedby%> @ <%=chkdatetime(sj_submitteddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
        <td><%=sj_cancelledby%> @ <%=chkdatetime(sj_cancelleddate)%></td>
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
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Current Qty</strong></font></td>
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>ADJ Qty</strong></font></td>
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Different Qty</strong></font></td>
      </tr>
    <%				i = 1
				sql1 = "SELECT sjd_id, sjd_sj_no, sjd_itm_code, sjd_itm_desc, sjd_unitcost, sjd_current_qty, sjd_adjust_qty, sjd_diff_qty, sjd_subtotal, sjd_referid " & _
	                   "FROM tblstockadj_detail where sjd_sj_no = '" & sj_no & "' order by sjd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
      <tr>
        <td align="center" bgcolor="#FFFFFF"><%=i%>.</td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("sjd_itm_code")%></span></td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("sjd_itm_desc")%></span></td>
        <td align="center" bgcolor="#FFFFFF"><%=rs1("sjd_current_qty")%></td>
        <td align="center" bgcolor="#FFFFFF"><span class="tktTotals"><%=rs1("sjd_adjust_qty")%></span></td>
        <td align="center" bgcolor="#FFFFFF"><span class="tktTotals"><%=rs1("sjd_diff_qty")%></span></td>
      </tr>
 <%	
				sjd_current_qty = sjd_current_qty + rs1("sjd_current_qty")
				sjd_adjust_qty = sjd_adjust_qty + rs1("sjd_adjust_qty")
				sjd_diff_qty = sjd_diff_qty + rs1("sjd_diff_qty")
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
   
      <tr bgcolor="#FFFFFF">
        <td height="25" colspan="3" align="right"><strong>Total</strong></td>
        <td align="center"><strong><%=sjd_current_qty%></strong></td>
        <td align="center"><strong><%=sjd_adjust_qty%></strong></td>
        <td align="center">&nbsp;</td>
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
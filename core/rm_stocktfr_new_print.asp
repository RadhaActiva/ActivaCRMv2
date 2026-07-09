<!-- #include file="database/datastore.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")

if request("sf_no") <> "" then	  
sql = "SELECT sf_id, sf_no, sf_date, sf_referenceno, sf_status, sf_fromwarehouse, sf_towarehouse, sf_remark, sf_createddate, sf_createdby, sf_submitteddate, " & _
		"sf_submittedby, sf_approveddate, sf_approveddate, sf_approvedby, sf_cancelleddate, sf_cancelledby, sf_totalqty, sf_totalaAmt, sf_emailsent, sf_emailsentdate " & _
		"FROM tblstocktransfer WHERE sf_no = '" & request("sf_no") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			sf_id = rs("sf_id") 
			sf_no = rs("sf_no") 
			sf_date = rs("sf_date")
			sf_referenceno = rs("sf_referenceno") 
			sf_status = rs("sf_status")
			sf_fromwarehouse = rs("sf_fromwarehouse") 
			sf_towarehouse = rs("sf_towarehouse")
			sf_remark = rs("sf_remark")
			sf_createddate = rs("sf_createddate") 
			sf_createdby = rs("sf_createdby") 
			sf_submitteddate = rs("sf_submitteddate") 
			sf_submittedby = rs("sf_submittedby") 
			sf_approveddate = rs("sf_approveddate") 
			sf_approvedby = rs("sf_approvedby") 
			sf_cancelleddate = rs("sf_cancelleddate") 
			sf_cancelledby = rs("sf_cancelledby") 
			sf_totalqty = rs("sf_totalqty") 
			sf_totalaAmt = rs("sf_totalaAmt") 
			sf_emailsent = rs("sf_emailsent")  
			sf_emailsentdate = rs("sf_emailsentdate")
		End If
		rs.Close
	  stype = "editStockTransfer"	
	  actionname = "Save"  	
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
        <td bgcolor="#FFFFFF" class="titleblue1"><div align="left">STOCK-TRANSFER</div></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#EBEBEB">
      <tr>
        <td align="left" bgcolor="#475387"><font color="#FFFFFF"><strong><font color="#FFFFFF"><strong>Stock-Transfer </strong></font> No</strong></font></td>
        <td nowrap="nowrap"><strong><%=sf_no%></strong></td>
        <td height="22" nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Stock-Transfer date</strong></font></strong></td>
        <td><%=chkdate(sf_date)%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Reference No</strong></font></strong></td>
        <td nowrap="nowrap"><%=sf_referenceno%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Status</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=sf_status%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>From Store</strong></font></strong></td>
        <td nowrap="nowrap"><%=sf_fromwarehouse%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Remark</strong></font></strong></td>
        <td align="left" nowrap="nowrap"><%=sf_remark%></td>
        </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>To Store</strong></font></strong></td>
        <td nowrap="nowrap"><%=sf_towarehouse%></td>
        <td nowrap="nowrap" bgcolor="#475387">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Created by</strong></font></strong></td>
        <td nowrap="nowrap"><%=sf_createdby%> @ <%=chkdatetime(sf_createddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Approved By</strong></font></strong></td>
        <td><%=sf_approvedby%> @ <%=chkdatetime(sf_approveddate)%></td>
      </tr>
      <tr>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Submitted By</strong></font></strong></td>
        <td nowrap="nowrap"><%=sf_submittedby%> @ <%=chkdatetime(sf_submitteddate)%></td>
        <td nowrap="nowrap" bgcolor="#475387"><strong><font color="#FFFFFF"><strong>Cancelled by</strong></font></strong></td>
        <td><%=sf_cancelledby%> @ <%=chkdatetime(sf_cancelleddate)%></td>
        </tr>
    </table></td>
  </tr>
 
  <tr>
    <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><table width="98%" border="1" cellpadding="8" cellspacing="0">
      <tr bgcolor="#475387">
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>No</strong></font></td>
        <td align="left" bgcolor="#EAEAEA"><font color="#000000"><strong>Item 
          Code</strong></font></td>
        <td align="left" bgcolor="#EAEAEA"><font color="#000000"><strong> Item Description</strong></font></td>
        <td align="center" bgcolor="#EAEAEA"><font color="#000000"><strong>Stock-In Qty</strong></font></td>
      </tr>
      <%				i = 1
				sql1 = "SELECT sfd_id, sfd_st_no, sfd_itm_code, sfd_itm_desc, sfd_unitcost, sfd_qty, sfd_subtotal, sfd_referid " & _
	                  "FROM tblstocktran_detail where sfd_st_no = '" & sf_no & "' order by sfd_id"	   
					   'response.write sql1
				set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
      <tr>
        <td align="center" bgcolor="#FFFFFF"><%=i%>.</td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("sfd_itm_code")%></span></td>
        <td align="left" bgcolor="#FFFFFF"><span style="text-align: left"><%=rs1("sfd_itm_desc")%></span></td>
        <td align="center" bgcolor="#FFFFFF"><span class="tktTotals"><%=rs1("sfd_qty")%></span></td>
      </tr>
    <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%> 
      <tr bgcolor="#FFFFFF">
        <td height="25" colspan="3" align="right"><strong>Total</strong></td>
        <td align="center"><strong><%=sf_totalqty%></strong></td>
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
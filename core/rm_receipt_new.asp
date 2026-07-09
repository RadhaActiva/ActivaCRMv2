<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

receipt_no = request.form("receipt_no")
if receipt_no = "" then
    receipt_no = request.QueryString("receipt_no")
end if

receipt_inv_no = request.form("receipt_inv_no")
if receipt_inv_no = "" then
    receipt_inv_no = request.QueryString("receipt_inv_no")
end if

if receipt_no <> "" then	  
sql = "SELECT receipt_id, receipt_no, receipt_status, receipt_date, receipt_inv_no, receipt_inv_date, receipt_cust_code, receipt_cust_name, " & _
	"receipt_cust_address, receipt_cust_postcode, receipt_cust_state, receipt_cust_state_id, receipt_cust_city, receipt_cust_city_id,  " & _
	"receipt_cust_email, receipt_cust_tel1, receipt_cust_tel2, receipt_createddate, receipt_createdby, receipt_job_code, receipt_remark, " & _ 
	"receipt_paymenttype, receipt_totalpayment, receipt_emailsent, receipt_emailsentdate, receipt_cancelleddate, receipt_cancelledby " & _
	"FROM tblreceipt WHERE receipt_no = '" & receipt_no & "' and receipt_inv_no = '" & receipt_inv_no & "'"
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			receipt_id = rs("receipt_id") 
			receipt_no = rs("receipt_no") 
			receipt_status = rs("receipt_status")
			receipt_date = rs("receipt_date") 
			receipt_inv_no = rs("receipt_inv_no")  
			receipt_inv_date = rs("receipt_inv_date")  
			receipt_cust_code = rs("receipt_cust_code") 
			receipt_cust_name = rs("receipt_cust_name")  
			receipt_cust_address = rs("receipt_cust_address")  
			receipt_cust_postcode = rs("receipt_cust_postcode")  
			receipt_cust_state = rs("receipt_cust_state")  
			receipt_cust_state_id = rs("receipt_cust_state_id")  
			receipt_cust_city = rs("receipt_cust_city")  
			receipt_cust_city_id = rs("receipt_cust_city_id") 
			receipt_cust_email = rs("receipt_cust_email")  
			receipt_cust_tel1 = rs("receipt_cust_tel1")  
			receipt_cust_tel2 = rs("receipt_cust_tel2")  
			receipt_createddate = rs("receipt_createddate")  
			receipt_createdby = rs("receipt_createdby")  
			
			receipt_job_code = rs("receipt_job_code")  
			receipt_inv_no = rs("receipt_inv_no")  
			receipt_inv_date = rs("receipt_inv_date") 
			receipt_remark = rs("receipt_remark")  
			
			receipt_paymenttype = rs("receipt_paymenttype")  
			receipt_totalpayment = rs("receipt_totalpayment") 
			receipt_emailsent = rs("receipt_emailsent")  
			receipt_emailsentdate = rs("receipt_emailsentdate")  
			receipt_cancelledby = rs("receipt_cancelledby")  
			receipt_cancelleddate = rs("receipt_cancelleddate")  
		End If
		rs.Close
	    stype = "editReceipt"	
	    actionname = "Save" 
end if

if receipt_no <> "" then	  'check if to gen receipt for single inv or multiple invoices as format is diff
    sql = "select count(*) as receiptnum  from tblreceipt WHERE receipt_no = '" & receipt_no & "'"
	receiptnum = selectid(sql)
end if 
%>

<script language="javascript">

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.forminvoicedetail.action = orderlinks;
	document.forminvoicedetail.submit();
   }
}

</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><font color="#CC0000">View </font>Receipt</td>
                        <td width="23%" align="right" class="titleblue1">
                        <%if receipt_status="Posted" then %>
                            <% if receiptnum >  1 then%>	
                                 <a href="rm_receipt_multiple.asp?receipt_no=<%=receipt_no%>" target="_blank"><img src="images/m_im_icon_print.gif" alt="Print A4 format" border="0" style="border:0"/></a>
                            <%else%>
                                 <a href="rm_receipt_new_print.asp?receipt_no=<%=receipt_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35"" alt="Print A4 format" border="0" style="border:0"/></a>
                            <%end if%>
                            <%end if%>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                
                
                  <%if receipt_no <> "" then %>    
                 <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Code *</strong></font></td>
                        <td align="left"><label for="receipt_cust_code"></label>
                          <input name="receipt_cust_code" type="text" id="receipt_cust_code" style="background-color: #cccccc;" value="<%=receipt_cust_code%>" maxlength="50" />
                          <input type="receipt_no" name="receipt_no" id="receipt_no"  value="<%=receipt_no%>"/></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
                        <td align="left"><input name="receipt_cust_name" type="text" id="receipt_cust_name" value="<%=receipt_cust_name%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><strong>
                          <textarea name="receipt_cust_address" cols="50" rows="3" id="receipt_cust_address"><%=receipt_cust_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left"><strong>
                          <input name="receipt_cust_postcode" type="text" id="textfield16" value="<%=receipt_cust_postcode%>" size="20" maxlength="20" />
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                           <td align="left">
                        <input name="receipt_cust_state" type="text" id="receipt_cust_state" value="<%=receipt_cust_state%>" size="30" readonly maxlength="50" />
                        <input name="receipt_cust_state_id" type="hidden" id="receipt_cust_state_id" value="<%=receipt_cust_state_id%>" size="30" maxlength="50" />
                        </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                        <input name="receipt_cust_city" type="text" id="receipt_cust_city" value="<%=receipt_cust_city%>" size="30" readonly maxlength="50" />
                        <input name="receipt_cust_city_id" type="hidden" id="receipt_cust_city_id" value="<%=receipt_cust_city_id%>" size="30" maxlength="50" />
                        </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="receipt_cust_email" type="text" id="receipt_cust_email" value="<%=receipt_cust_email%>" size="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><label for="receipt_cust_tel1"></label>
                          <input name="receipt_cust_tel1" type="text" id="receipt_cust_tel1" value="<%=receipt_cust_tel1%>" size="30" maxlength="50" />
                          e.g 0121234657</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="receipt_cust_tel2" type="text" id="receipt_cust_tel2" value="<%=receipt_cust_tel2%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td valign="top"><strong>
                          <textarea name="receipt_remark" cols="50" rows="3" id="receipt_remark"><%=receipt_remark%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td colspan="2" align="right" valign="top" ><input type="submit" name="button" id="button" value="<%=actionname%>" /></td>
                        </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> Receipt Information</font></strong></td>
                      </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Receipt  
                          No.<br />
                          <font size="1">(System Generate) </font></strong></font></td>
                        <td colspan="3" align="left"><strong><%=receipt_no%></strong></td>
                        </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Receipt  
                          Date</strong></font></td>
                        <td colspan="3" align="left"><%=chkdate(receipt_date)%></td>
                        </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status</strong></font></td>
                        <td colspan="3" align="left"><strong><%=receipt_status%></strong></td>
                        </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Invoice No.</strong></font></td>
                        <td colspan="3" align="left"><strong><%=receipt_inv_no%></strong></td>
                      </tr>
                      <tr align="left" >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job Code</strong></font></td>
                        <td colspan="3" align="left"><strong><%=receipt_job_code%></strong></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Payment Amount</strong></font></td>
                        <td colspan="3" align="left"><strong>RM <%=receipt_totalpayment%></strong></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Payment Type</strong></font></td>
                        <td colspan="3" align="left"><%=receipt_paymenttype%></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Payment Remark</strong></font></td>
                        <td colspan="3" align="left"><%=receipt_remark%></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Created by</strong></font></td>
                        <td colspan="3" align="left"><%=receipt_createdby%> @ <%=chkdatetime(receipt_createddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cancelled  by</strong></font></td>
                        <td colspan="3" align="left"><%=receipt_cancelledby%> @ <%=chkdatetime(receipt_cancelleddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF"><label for="receipt_no"></label></td>
              </tr>
              </form>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col">&nbsp;</td>
                    </tr>
                    <form name="formDOemail" id="formDOemail" method="post" action="action.asp?type=submitCN&amp;receipt_no=<%=receipt_no%>&amp;#spareparts" >
                      <tr>
                        <td width="55%" align="left" bgcolor="#FFFFFF" scope="col"><!-- <strong>Email</strong>
                          <input name="emailto_DO" type="text" id="emailto_DO" value="<%=receipt_cust_email%>" size="50" maxlength="150" />
                          <input type="button" name="Submit523" value="Email Receipt" style="{width:200px}" onclick="javascript:popup('rm_receipt_new_email.asp?receipt_no=<%=receipt_no%>&emailto=' + formDOemail.emailto_DO.value + '&email_remark=' + formDOemail.email_remark.value,'cb17','scrollbars=yes,resizable=yes,width=600px,height=600px')" />
                          <br />
<strong><font color="#000000">Remark Message: </font></strong>:
                                <input name="email_remark" type="text" id="email_remark" size="50" maxlength="200" />
                          <br />
                          <br />-->
                          <%if receipt_status<>"Cancel" and (request.Cookies("GAPS")("sloginid") = "davidhui" or request.Cookies("GAPS")("sloginid")="ERICLOH") then %>
                          <input type="button" name="CancelJob" id="CancelJob" value="Cancel Receipt " onclick="javascript:confirmAction('<%=receipt_no%>','action.asp?type=CancelReceipt&receipt_no=<%=receipt_no%>')" />
                        <%end if%></td>
                        <td width="45%" align="right" bgcolor="#FFFFFF" 
          scope="col"><br /></td>
                      </tr>
                    </form>
                    <tr align="right">
                      <td colspan="2" bgcolor="#FFFFFF"></td>
                    </tr>
                    <tr>
                      <td></tbody></td>
                    </tr>
                  </table></td>
                </tr>
                <%end if%>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            height: 68px;
        }
    </style>
</head>
<%

set rs = server.CreateObject("adodb.recordset")

if request("cn_no") <> "" then	  
sql = "SELECT cn_id, cn_no, cn_status, cn_date, cn_inv_no, cn_inv_date, cn_cust_code, cn_cust_name, cn_cust_address, cn_cust_postcode, " & _
	  "cn_cust_state, cn_cust_state_id, cn_cust_city, cn_cust_city_id, cn_cust_cnty_id,cn_cust_email, cn_cust_tel1, cn_cust_tel2, cn_createddate, cn_createdby,  " & _
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
            cn_cust_cnty_id = rs("cn_cust_cnty_id")
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
	  stype = "editCN"	
	  actionname = "Save" 
 else    
	  stype = "addCN"
	  actionname = "Save" 		
	  cn_date = date()	 
	  cn_status = "Open"   	
end if

if cn_cust_postcode <> "" and cn_cust_cnty_id = "129" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & cn_cust_postcode & "' "
		rs1.Open sql1,strconnect,0,1,&H0001   
		If Not rs1.EOF Then
             cn_cust_state_id = rs1("state_id") 'will auto populate state
             cn_cust_state =  rs1("state_name")
             cn_cust_city_id = rs1("city_id") 'will auto populate city
             cn_cust_city = rs1("post_office")    
        end if
    rs1.close
end if
 
%>

<script language="javascript">

function getPostcode(p)
{
     document.getElementById('cn_cust_postcode').value = p;
     document.formorder.submit();
    }

    function getCountrycode(c) {
        document.getElementById('cn_cust_cnty_id').value = c;
        document.formorder.submit();
    }

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.forminvoicedetail.action = orderlinks;
	document.forminvoicedetail.submit();
   }
}


function calctotal(unitprice,qty,discountamt,discounttype,subtotal) {

var discount = 0;
var temp = 0;

if (discounttype == "%")
{
	discount = unitprice * (discountamt/100);
	temp = (unitprice-discount) * qty;
	subtotal.value = temp.toFixed(2);
	}
else
{
	discount = discountamt;
	temp = (unitprice-discount) * qty;
	subtotal.value = temp.toFixed(2);
	}
}
// -->
</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><font color="#CC0000">Create </font>CN</td>
                        <td width="23%" align="right" class="titleblue1">
                        <%if cn_status="Posted" then %>
                        <a href="rm_cn_new_print.asp?cn_no=<%=cn_no%>" target="_blank"><img src="images/A4_icon.png"  height="35" width="35" alt="Print A4 format" border="0" style="border:0"/></a>
                        <%end if%>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                
                 <%if cn_no = "" then %>    
                <form name="forminvoice" method="post" action="action.asp?type=generateCN">
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV4">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Invoice Information</font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Invoice No. [<a href="javascript:popup('rm_cn_invoiceno.asp','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </strong></font></td>
                        <td align="left">
                          <input name="cn_invoice_no" type="text" id="cn_invoice_no" value="<%=cn_invoice_no%>" maxlength="50" />
                          <input type="submit" name="button3" id="button3" value="Generate CN" /></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                </form>
                <%end if%>
                
                  <%if cn_no <> "" then %>    
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
                        <td align="left"><label for="cn_cust_code"></label>
                          <input name="cn_cust_code" type="text" id="cn_cust_code" style="background-color: #cccccc;" value="<%=cn_cust_code%>" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cust Name *</strong></font></td>
                        <td align="left"><input name="cn_cust_name" type="text" id="cn_cust_name" value="<%=cn_cust_name%>" size="50" maxlength="100" /></td>
                      </tr>
                         <tr>
                        <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left">
                                    <select name="cn_cust_cnty_id" id="cn_cust_cnty_id" style="width:150px"  onblur="getCountrycode(this.value)">
                                    <option value="<%=cn_cust_cnty_id%>"></option>                                       
                                    <%
                                          sql = "SELECT cnty_name,cnty_id from tblcountry"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((cn_cust_cnty_id)) = cstr((rs1("cnty_id"))) then
					                               response.write "<option value='" & rs1("cnty_id") & "' selected>" & rs1("cnty_name") & "</option>"
					                            else
					                               response.write "<option value='" & rs1("cnty_id") & "'>" & rs1("cnty_name") & "</option>"
					                            end if 	                                
                                          rs1.movenext 
                                          wend                                                                       
                                          rs1.close 
                                    %>
                                    </select>       
                        </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><strong>
                          <textarea name="cn_cust_address" cols="50" rows="3" id="cn_cust_address"><%=cn_cust_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left"><strong>
                          <input name="cn_cust_postcode" type="text" id="cn_cust_postcode" value="<%=cn_cust_postcode%>" onblur="getPostcode(this.value)" size="20" maxlength="20" />
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left">                        
                         <input name="job_cust_state" type="text" id="cn_cust_state" value="<%=cn_cust_state%>" size="30" readonly maxlength="50" />
                         <input name="job_cust_state_id" type="hidden" id="cn_cust_state_id" value="<%=cn_cust_state_id%>" size="30" readonly maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                         <input name="cn_cust_city" type="text" id="cn_cust_city" value="<%=cn_cust_city%>" size="30" readonly maxlength="50" />
                         <input name="jcn_cust_city_id" type="hidden" id="cn_cust_city_id" value="<%=cn_cust_city_id%>" size="30" readonly maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="cn_cust_email" type="text" id="cn_cust_email" value="<%=cn_cust_email%>" size="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><label for="cn_cust_tel1"></label>
                          <input name="cn_cust_tel1" type="text" id="cn_cust_tel1" value="<%=cn_cust_tel1%>" size="30" maxlength="50" />
                          e.g 0121234657</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="cn_cust_tel2" type="text" id="cn_cust_tel2" value="<%=cn_cust_tel2%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td valign="top"><strong>
                          <textarea name="cn_remark" cols="50" rows="3" id="cn_remark"><%=cn_remark%></textarea>
                        </strong></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong><font size="2"> CN Information</font></strong></td>
                      </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>CN  
                          No.<br />
                          <font size="1">(System Generate) </font></strong></font></td>
                        <td colspan="3" align="left"><strong><%=cn_no%></strong></td>
                        </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>CN  
                          Date</strong></font></td>
                        <td colspan="3" align="left"><%=chkdate(cn_date)%></td>
                        </tr>
                      <tr >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status</strong></font></td>
                        <td colspan="3" align="left"><strong><%=cn_status%></strong></td>
                        </tr>
                      <tr align="left" >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Job Code</strong></font></td>
                        <td colspan="3" align="left"><input name="cn_job_code" type="text" id="cn_job_code" value="<%=cn_job_code%>" size="20" maxlength="50" /></td>
                        </tr>
                      <tr align="left" >
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>DO No.</strong></font></td>
                        <td colspan="3" align="left"><label for="textfield5">
                          <input name="cn_do_no" type="text" id="cn_do_no" value="<%=cn_do_no%>" size="20" maxlength="50" />
                        </label></td>
                        </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Invoice No.</strong></font></td>
                        <td colspan="3" align="left"><input name="cn_inv_no" type="text" id="cn_inv_no" value="<%=cn_inv_no%>" size="20" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Created by</strong></font></td>
                        <td colspan="3" align="left"><%=cn_createdby%>@ <%=chkdatetime(cn_createddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Returned by</strong></font></td>
                        <td colspan="3" align="left"><%=cn_returnedby%> @ <%=chkdatetime(cn_returneddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Done by</strong></font></td>
                        <td colspan="3" align="left"><%=cn_doneby%> @ <%=chkdatetime(cn_donedate)%></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Posted by</strong></font></td>
                        <td colspan="3" align="left"><%=cn_postedby%> @ <%=chkdatetime(cn_posteddate)%></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Cancelled  by</strong></font></td>
                        <td colspan="3" align="left"><%=cn_cancelledby%> @ <%=chkdatetime(cn_cancelleddate)%></td>
                      </tr>
                      
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF"><label for="cn_no"></label>
                    <input type="hidden" name="cn_no" id="cn_no" value="<%=cn_no%>" />
                    <%if cn_status="Open" then %>
                    <input type="submit" name="button" id="button" value="<%=actionname%>" />
                  <%end if%></td>
              </tr>
              </form>
              
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><a name="spareparts" id="spareparts"></a></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong>Spare Part 
                            Code</strong></font></td>
                          <td align="left"><font color="#FFFFFF"><strong> Description</strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Unit Price (RM)</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Qty</strong></font></td>
                          <td width="5%" align="right"><font color="#FFFFFF"><strong>Discount </strong></font></td>
                          <td align="right"><font color="#FFFFFF"><strong>Total 
                            Amt (RCP)</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Action</strong></font></td>
                        </tr>
                        <%

if request("cnd_id") <> "" then
		sql = "SELECT cnd_id, cnd_cn_no, cnd_job_code, cnd_partcode, cnd_desc, cnd_unitcost, cnd_qty, cnd_discountamt, " & _
		      "cnd_discounttype, cnd_netcost, cnd_subtotal " & _
	          "FROM tblcn_detail where cnd_id = '" & request("cnd_id") & "'"	
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
		   cnd_id = rs("cnd_id")
		   cnd_cn_no = rs("cnd_cn_no")
		   cnd_job_code = rs("cnd_job_code")
		   cnd_partcode = rs("cnd_partcode")
		   cnd_desc = rs("cnd_desc")
		   cnd_unitcost = rs("cnd_unitcost")
		   cnd_qty = rs("cnd_qty")
		   cnd_discountamt = rs("cnd_discountamt")
		   cnd_discounttype = rs("cnd_discounttype")  
		   cnd_netcost = rs("cnd_netcost")   
		   cnd_subtotal = rs("cnd_subtotal")
        end if
		rs.close
		sbutton = "Update"
		stype="editCNDetail"	
else
		sbutton = "Add"
		stype="addCNDetail"
		cnd_qty = "1"	
		cnd_unitcost = "0.00"	
		cnd_discountamt = "0.00"
		cnd_netcost = "0.00"	
		cnd_subtotal = "0.00"	
end if

%>
                        <%if cn_status="Open" then %>
                        <form name="forminvoicedetail" id="forminvoicedetail" method="post" action="rm_jobsheet.asp#spareparts" >
                          <tr>
                            <td bgcolor="#666666" class="auto-style1"></td>
                            <td align="left" bgcolor="#666666" class="auto-style1"><input name="cnd_partcode" type="text" id="cnd_partcode" value="<%=cnd_partcode%>" maxlength="50" />
                              [<a href="javascript:popup('rm_cn_new_model.asp','cb17','scrollbars=yes,resizable=yes,width=500,height=500')">Select</a>] </td>
                            <td align="left" bgcolor="#666666" class="auto-style1"><font color="#FFFFFF"> </font>
                              <label for="cnd_desc"></label>
                              <textarea name="cnd_desc" cols="30" rows="3" id="cnd_desc"><%=cnd_desc%></textarea></td>
                            <td align="right" bgcolor="#666666" class="auto-style1"><font color="#FFFFFF">
                              <input type="hidden" name="cn_no" id="cn_no" value="<%=cn_no%>" />
                              <input type="hidden" name="cnd_id" id="cnd_id" value="<%=cnd_id%>" />
                              <input name="cnd_unitcost" type="text" id="cnd_unitcost" style="text-align:right; background-color: #cccccc;" onkeydown="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);" value="<%=cnd_unitcost%>" size="5" maxlength="10" />
                            </font></td>
                            <td align="right" bgcolor="#666666" class="auto-style1"><input name="cnd_qty" type="text" id="cnd_qty" style="text-align:right" onkeydown="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);" value="<%=cnd_qty%>" size="5" maxlength="5" /></td>
                            <td align="right" nowrap="nowrap" bgcolor="#666666" class="auto-style1"><font color="#FFFFFF">
                              <input name="cnd_discountamt" type="text" id="cnd_discountamt" value="<%=cnd_discountamt%>" size="5" onkeydown="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);" style="text-align:right" />
                              <select name="cnd_discounttype" id="cnd_discounttype" onchange="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);" onkeyup="calctotal(document.forminvoicedetail.cnd_unitcost.value, document.forminvoicedetail.cnd_qty.value, document.forminvoicedetail.cnd_discountamt.value, document.forminvoicedetail.cnd_discounttype.value, document.forminvoicedetail.cnd_subtotal);">
                                <option value="%" <%if cnd_discounttype = "%" then response.write " selected"%>>%</option>
                                <option value="RM" <%if cnd_discounttype = "RM" then response.write " selected"%>>RM</option>
                              </select>
                            </font></td>
                            <td align="right" bgcolor="#666666" class="auto-style1"><input name="cnd_subtotal" type="text" id="cnd_subtotal" style="text-align:right; background-color: #cccccc;" onfocus="this.blur();" value="<%=cnd_subtotal%>" size="10" maxlength="10" /></td>
                            <td align="center" bgcolor="#666666" class="auto-style1"><input type="button" name="button2" id="button2" value="<%=sbutton%>" onclick="javascript:confirmForm('<%=cn_no%>','action.asp?type=<%=stype%>','<%=cnd_subtotal%>');" /></td>
                          </tr>
                        </form>
                        <%end if%>
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
                          <td align="center" nowrap="nowrap"><%if cn_status="Open" then %>
                            <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_cn_new.asp?cnd_id=<%=rs1("cnd_id")%>&amp;cn_no=<%=cn_no%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("cnd_partcode")%>','action.asp?type=delCNDetail&amp;cnd_id=<%=rs1("cnd_id")%>&amp;cn_no=<%=cn_no%>')" />
                            <%end if%></td>
                        </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="6" align="right"><strong>Total</strong></td>
                          <td align="right"><strong><%=chknumber2(cn_totalAmt)%></strong></td>
                          <td>&nbsp;</td>
                        </tr>
                        <tr bgcolor="#EAEAEA">
                          <td height="25" colspan="8" align="left">**GST 6% Inclusive, 
                            GST Amount: RM <%=chknumber2(cn_gstAmt)%></td>
                        </tr>
                      </table></td>
                    </tr>
                    <form name="formDOemail" id="formDOemail" method="post" action="action.asp?type=submitCN&amp;cn_no=<%=cn_no%>&amp;#spareparts" >
                      <tr>
                      <td width="55%" align="left" bgcolor="#FFFFFF" scope="col"> <!--  <strong>Email</strong>
                          <input name="emailto_DO" type="text" id="emailto_DO" value="<%=cn_cust_email%>" size="50" maxlength="150" />
                          <input type="button" name="Submit523" value="Email CN" style="{width:200px}" onclick="javascript:popup('rm_cn_new_email.asp?cn_no=<%=cn_no%>&emailto=' + formDOemail.emailto_DO.value + '&email_remark=' + formDOemail.email_remark.value,'cb17','scrollbars=yes,resizable=yes,width=600px,height=600px')" />
                          <br />
<strong><font color="#000000">Remark Message: </font></strong>:
                                <input name="email_remark" type="text" id="email_remark" size="50" maxlength="200" />
                          <br />
                          <br />-->
                          <%if cn_status<>"Posted" or cn_status<>"Cancel" and (request.Cookies("GAPS")("sloginid") = "davidhui" or request.Cookies("GAPS")("sloginid")="ERICLOH") then %>
                          <input type="button" name="CancelJob" id="CancelJob" value="Cancel CN " onclick="javascript:confirmAction('<%=cn_no%>','action.asp?type=CancelCN&cn_no=<%=cn_no%>')" />
                        <%end if%></td>
                        <td width="45%" align="right" bgcolor="#FFFFFF" 
          scope="col"><%if cn_status="Open" then %>
                          <input type="button" name="SubmitJob2" id="SubmitJob2" value="Submit CN " onclick="javascript:confirmAction('<%=cn_no%>','action.asp?type=SubmitCN&cn_no=<%=cn_no%>')" />
                          <%end if%>
                          <br />
                          <%if cn_status="Submitted" then %>
                          <input type="button" name="DoneInvoice" id="DoneInvoice" value="Done CN " onclick="javascript:confirmAction('<%=cn_no%>','action.asp?type=DoneCN&cn_no=<%=cn_no%>')" />
                          <%end if%>
                          <br />
                          <%if cn_status="Done" then %>
                          <input type="button" name="PostedInvoice" id="PostedInvoice" value="Posted CN " onclick="javascript:confirmAction('<%=cn_no%>','action.asp?type=PostedCN&cn_no=<%=cn_no%>')" />
                          <%end if%>
                        </td>
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
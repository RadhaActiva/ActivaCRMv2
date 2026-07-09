<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

if request("warrantyno") <> "" then	  
sql = "SELECT refer_id, warrantyno, productmodel, othermodel, serialno, dealername, purchase_date, invoiceno, deliveryno, " & _
      "customername, customericno, customeremail, customeraddress, customerpostcode, customerstate, customercity, customertel1, customertel2, customerfax " & _
	  "FROM tblonlinewarranty WHERE warrantyno = '" & request("warrantyno") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			refer_id = rs("refer_id") 
			warrantyno = rs("warrantyno") 
			productmodel = rs("productmodel") 
			othermodel = rs("othermodel")
			serialno = rs("serialno")
			dealername = rs("dealername") 
			purchase_date = rs("purchase_date") 
			invoiceno = rs("invoiceno") 
			deliveryno = rs("deliveryno") 
			customername = rs("customername") 
			customericno = rs("customericno") 
			customeremail = rs("customeremail") 
			customeraddress = rs("customeraddress") 
			customerpostcode = rs("customerpostcode")  
			customerstate = rs("customerstate") 
			customercity = rs("customercity") 
			customertel1 = rs("customertel1") 
			customertel2 = rs("customertel2") 
			customerfax = rs("customerfax") 
		End If
		rs.Close
	  stype = "editWarrantyno"	
	  actionname = "Save" 
 else    
	  stype = "addWarrantyno"
	  actionname = "Save" 	
end if

%>

<script language="javascript">

function confirmForm(id,orderlinks,otype) 
{

  if (confirm("Are you sure you want to " + otype + " \n ID: " + id))
   {
	document.forminvoice.action = orderlinks;
	document.forminvoice.submit();
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
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Warranty </div></td>
                        <td width="23%" align="right" class="titleblue1">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
                </tr>
<tr>
                <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Warranty  
                        Information </font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Warranty Online No </strong></font></td>
                        <td align="left">
                        <%=warrantyno%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Product Model </strong></font></td>
                        <td align="left"><%=productmodel%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Other Model </strong></font></td>
                        <td align="left"><%=othermodel%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Serial No</strong></font></td>
                        <td align="left"><%=serialno%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Dealer Name</strong></font></td>
                        <td align="left"><%=dealername%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Purchased Date</strong></font></td>
                        <td align="left">
                        <%=chkdate(purchase_date)%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Dealer Invoice No </strong></font></td>
                        <td valign="top"><%=invoiceno%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Delivery No</strong></font></td>
                        <td valign="top">
                        <%=deliveryno%></td>
                      </tr>
                    </tbody>
                </table></td>
                <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer   
                        Information </font></strong></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Customer Name</strong></font></td>
                        <td align="left"><label for="inv_cust_code2"></label>
                          <%=customername%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong> ICNO</strong></font></td>
                        <td align="left"><%=customericno%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><p><font color="#FFFFFF"><strong> Email </strong></font></p></td>
                        <td align="left"><%=customeremail%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address</strong></font></td>
                        <td align="left"><%=customeraddress%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode</strong></font></td>
                        <td align="left"><%=customerpostcode%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State</strong></font></td>
                        <td align="left">
                          <%=customerstate%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City</strong></font></td>
                        <td valign="top"><%=customercity%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel 1</strong></font></td>
                        <td valign="top"><label for="inv_cust_tel2"></label>
                          <%=customertel1%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel 2</strong></font></td>
                        <td valign="top"><%=customertel2%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Fax</strong></font></td>
                        <td valign="top"><%=customerfax%></td>
                      </tr>
                    </tbody>
                  </table></td>
              </tr>
                
              
              <tr>
                <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
              
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
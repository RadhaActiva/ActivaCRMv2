<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset")

cust_name=request.querystring("cust_name")
cust_icno=request.querystring("cust_icno")
cust_cnty_id=request.querystring("cust_cnty_id")


if request("cust_code") <> "" then	  
      sql = "SELECT cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, cust_status, " & _
      "cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email, cust_tel1,  " & _
      "cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic, cust_icno, cust_debtor_code, cust_branch_code " & _
	  "FROM tblcustomer WHERE cust_code = '" & request("cust_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			cust_id = rs("cust_id") 
			cust_createdby = rs("cust_createdby") 
			cust_createddate = rs("cust_createddate") 
			cust_code = rs("cust_code") 
			cust_name = rs("cust_name") 
			cust_type = rs("cust_type") 
			cust_status = rs("cust_status")
			cust_reg_no = rs("cust_reg_no")
			cust_company = rs("cust_company")
			
			cust_address = rs("cust_address")
			cust_postcode = rs("cust_postcode") 
			cust_state = rs("cust_state") 
			cust_state_id = rs("cust_state_id") 
			cust_city = rs("cust_city") 
			cust_city_id = rs("cust_city_id") 
			cust_cnty_id = rs("cust_cnty_id") 
			cust_email = rs("cust_email") 
			cust_tel1 = rs("cust_tel1") 
			cust_tel2 = rs("cust_tel2") 
			cust_fax = rs("cust_fax") 
			cust_website = rs("cust_website") 
			cust_password = rs("cust_password")
			cust_gstregno = rs("cust_gstregno")  
			cust_lastjob_code = rs("cust_lastjob_code")  
			cust_source = rs("cust_source") 
			cust_attention = rs("cust_attention") 
			cust_pic = rs("cust_pic")  
			cust_icno = rs("cust_icno")  
            cust_debtor_code = rs("cust_debtor_code")
            cust_branch_code = rs("cust_branch_code")
            'cstateID= rs("cust_state_id")
		End If
		rs.Close
	  stype = "editCustomer"	
	  actionname = "Save" 
 else    
	  stype = "addCustomer"
	  actionname = "Save"  
end if

cust_state_id=""
cust_city_id=""
if cust_postcode <> "" and cust_cnty_id="129" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & cust_postcode & "' "
		rs1.Open sql1,strconnect,0,1,&H0001
		If Not rs1.EOF Then
             cust_state_id = rs1("state_id") 'will auto populate state
             cust_city_id = rs1("city_id") 'will auto populate city
             city_name = rs1("post_office")
        end if
    rs1.close
end if

custlabel=""
set rs3 = server.CreateObject("adodb.recordset")
sql3 = "SELECT cust_type FROM tblcustomer WHERE cust_code = '" & request("cust_code") & "' "
rs3.Open sql3,strconnect,0,1,&H0001
If Not rs3.EOF Then
    custlabel=rs3("cust_type")
    rs3.close
end if

%>

<script type="text/javascript">
    function validateCustomerForm() {
        
        var tel = document.forms["formorder"]["cust_tel1"].value;
        var cname = document.forms["formorder"]["cust_name"].value;
        var cpostcode = document.forms["formorder"]["cust_postcode"].value;
        var ccntyid = document.forms["formorder"]["cust_cnty_id"].value;
        var caddr = document.forms["formorder"]["cust_address"].value;
        caddr = caddr.trim();

        if (!cname || cname.trim() === "") {
            alert("Customer name is required.");
            document.forms["formorder"]["cust_name"].focus();
            return false;   // stop submit
        }

        if (caddr.length < 8) {
            alert("Customer address must be at least 8 characters.");
            document.forms["cust_address"].focus();
            return false;
        }

        if (!tel || tel.trim() === "") {
            alert("Customer phone number is required.");
            document.forms["formorder"]["cust_tel1"].focus();
            return false;   // stop submit
        }

        var telPattern = /^[0-9+\-]+$/;

        if (tel !== "" && !telPattern.test(tel)) {
            alert("Customer phone number can only contain numbers and '+ -'");
            document.forms["formorder"]["cust_tel1"].focus();
            return false;
        }

        //  no consecutive dash
        if (tel.indexOf("--") !== -1) {
            alert("Customer phone number cannot contain consecutive '-'");
            document.forms["formorder"]["cust_tel1"].focus();
            return false;
        }

        if (!ccntyid || ccntyid.trim() === "" || ccntyid.toLowerCase() === "select") {
            alert("Customer country is required.");
            document.forms["formorder"]["cust_cnty_id"].focus();
            return false;   // stop submit
        }

        if (!cpostcode || cpostcode.trim() === "") {
            alert("Customer postcode is required.");
            document.forms["formorder"]["cust_postcode"].focus();
            return false;   // stop submit
        }
       
        return true;
    }
</script>

<script language="javascript">

    function getPostcode(p) {
        //alert(s);
        //document.getElementById('cust_name').value = s;
        document.getElementById('cust_postcode').value = p;
        document.formorder.submit();
    }

    function getCountrycode(c) {
        document.getElementById('cust_cnty_id').value = c;
        document.formorder.submit();
    }
// -->
</script>

        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
               <table width="100%"><!--for UI Purpose-->
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Customer </font>Master</div></td>
                        <td width="23%" align="right" class="titleblue1">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                <form name="formorder" method="post" action="action.asp?type=<%=stype%>" onsubmit="return validateCustomerForm();">
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV4">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Customer Information <font color="#006400">(<%=custlabel%>)</font></strong></td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Customer Code <br />
                        </strong>(System Generate) </font></td>
                        <td align="left"><%=cust_code%> 
                          <input type="hidden" name="cust_id" id="cust_id" value="<%=cust_id%>" />
                          <input type="hidden" name="cust_code" id="cust_code" value="<%=cust_code%>"/>                          
                        </td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Customer Name *</strong></font></td>
                        <td align="left">
                          <input name="cust_name" type="text" id="cust_name" value="<%=cust_name%>" size="50" maxlength="100" />
                        </td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong> ICNO *</strong></font></td>
                        <td align="left">
                          <input name="cust_icno" type="text" id="cust_icno" value="<%=cust_icno%>" size="50" maxlength="100" /></td>
                      </tr>
                         <td align="left" valign="top" bgcolor="#CD6155" class="auto-style4"><font color="#FFFFFF"><strong>Country*</strong></font></td>
                        <td align="left">
                                    <!--<select name="cust_cnty_id" id="cust_cnty_id" style="width:150px"  onblur="getCountrycode(this.value)">-->
                            <select name="cust_cnty_id" id="cust_cnty_id" style="width:150px">
                                    <option value="<%=cust_cnty_id%>"></option>                                       
                                    <%
                                          sql = "SELECT cnty_name,cnty_id from tblcountry"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((cust_cnty_id)) = cstr((rs1("cnty_id"))) then
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
                        </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><strong>
                          <textarea name="cust_address" cols="50" rows="3" id="cust_address"><%=cust_address%></textarea>
                        </strong>
                        </td>                      
                      </tr>

                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left">
                           <%if cust_cnty_id ="129" then' %> 
                                <input name="cust_postcode" type="text" id="cust_postcode" value="<%=cust_postcode%>"  onblur="getPostcode(this.value)" size="10" maxlength="10" /></td>
                            <%else%> 
                                <input name="cust_postcode" type="text" id="cust_postcode" value="<%=cust_postcode%>" size="10" maxlength="10" /></td>
                            <%end if%>                                                    
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font>                        
                        </td>
                        <td align="left">
                        <%if cust_cnty_id ="129" then' %> 
                            <input name="cust_state" type="text" id="cust_state" value="<%=cust_state%>" size="30" readonly maxlength="50" />
                            <input name="cust_state_id" type="hidden" id="cust_state_id" value="<%=cust_state_id%>" size="30" "readonly" maxlength="50" />
                        <%else%> 
                            <input name="cust_state" type="text" id="cust_state" value="<%=cust_state%>" size="30" style="background-color: #cccccc;" readonly maxlength="50" />
                            <input name="cust_state_id" type="hidden" id="cust_state_id" value="<%=cust_state_id%>" size="30" "readonly" maxlength="50" />
                        <%end if%>
                        </td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left">
                           <%if cust_cnty_id ="129" then' %> 
                             <!--<input name="cust_city_id" type="hidden" id="cust_city_id" value="<%=cust_city_id%>" size="6" maxlength="50" />  -->
                             <select name="cust_city_id" id="cust_city_id" style="width:150px">
                             <option value="<%=job_cust_city_id%>"></option>
                                    <%
                                          sql = "SELECT distinct city_id, post_office FROM tblpostcode where postcode = '" & cust_postcode & "'"	
                                          set rs1 = server.CreateObject("adodb.recordset")
				                          rs1.Open sql,strconnect,3,3,&H0001                                      
                                          While Not rs1.EOF		                                
                                                if cstr((cust_city_id)) = cstr((rs1("city_id"))) then
					                               response.write "<option value='" & rs1("city_id") & "' selected>" & rs1("post_office") & "</option>"
					                            else
					                               response.write "<option value='" & rs1("city_id") & "'>" & rs1("post_office") & "</option>"
					                            end if 	                                
                                          rs1.movenext 
                                          wend                                                                       
                                          rs1.close 
                                    %>
                                        </select> 
                            <%else%>
                                    <input name="cust_city_id" type="hidden" id="cust_city_id" value="<%=cust_city_id%>" size="6" maxlength="50" /> 
                                    <input name="cust_city" type="text" id="cust_city" value="<%=cust_city%>" size="30" maxlength="50" />
                            <%end if%>
                       </td>
                      </tr>                       
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><input name="cust_email" type="text" id="textfield2" value="<%=cust_email%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top">
                          <input name="cust_tel1" type="text" id="cust_tel1" value="<%=cust_tel1%>" maxlength="50" />
                          e.g 012-1234657</td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><input name="cust_tel2" type="text" id="cust_tel2" value="<%=cust_tel2%>" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Fax</strong></font></td>
                        <td valign="top"><input name="cust_fax" type="text" id="cust_fax" value="<%=cust_fax%>" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                        <td valign="top"><%=cust_createdby%><br />
                        <%=chkdatetime(cust_cretateddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Status  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Type</strong></font></td>
                        <td align="left">
                            <%= Server.HTMLEncode(cust_type) %><input type="hidden" name="cust_type" value="<%= Server.HTMLEncode(cust_type) %>">
                        <!--<select name="cust_type" id="cust_type">
                            <option value="customer" <%if cust_type="Customer" then response.write " selected"%>>Customer</option>
                            <option value="dealer" <%if cust_type="Dealer" then response.write " selected"%>>Dealer</option>
                         </select>-->
                        </td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status  *</strong></font></td>
                        <td align="left">
                        <select name="cust_status" id="cust_status">
                            <option value="Y" <%if cust_status="Y" then response.write " selected"%>>Y</option>
                            <option value="N" <%if cust_status="N" then response.write " selected"%>>N</option>
                         </select>
                        </td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong> Company Name</strong></font></td>
                        <td align="left">
                          <input name="cust_company" type="text" id="cust_company" value="<%=cust_company%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Company Reg No</strong></font></td>
                        <td align="left"><strong>
                          <input name="cust_reg_no" type="text" id="cust_reg_no" value="<%=cust_reg_no%>" size="20" maxlength="50" />
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Website</strong></font></td>
                        <td align="left"><input name="cust_website" type="text" id="textfield6" value="<%=cust_website%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Attention</strong></font></td>
                        <td valign="top"><label for="textfield4"></label>
                        <input name="cust_attention" type="text" id="textfield4" value="<%=cust_attention%>" size="50" maxlength="100" /></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>PIC</strong></font></td>
                        <td valign="top"><input name="cust_pic" type="text" id="textfield7" value="<%=cust_pic%>" size="50" maxlength="100" /></td>
                      </tr>
                        <tr><td colspan="2" bgcolor="#FFFFFF" scope="col">&nbsp;</td></tr>
                        <tr>
                          <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Autocount Codes</font></strong></td>
                      </tr>
                        <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Debtor Code</strong></font></td>
                        <td valign="top"><input name="cust_pic" type="text" id="textfield9" value="<%=cust_debtor_code%>" style="background-color: #cccccc;" readonly maxlength="50" /></td>
                      </tr>
                        <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Branch Code</strong></font></td>
                        <td valign="top"><input name="cust_pic" type="text" id="textfield10" value="<%=cust_branch_code%>" style="background-color: #cccccc;" readonly maxlength="50" /></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF"><input type="submit" name="button2" id="button" value="<%=actionname%>" /></td>
                </tr>
                </form>
                
                
                <%if cust_code <> "" then %>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col">&nbsp;</td>
                    </tr>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#333333">
                          <td colspan="10" bgcolor="#E8E8E8"><strong><font size="2">Wrty Card No.<a name="spareparts" id="spareparts"></a></font></strong></td>
                        </tr>

 <%

if request("refer_id") <> "" then
		sql = "SELECT refer_id, warrantyno, productmodel, othermodel, serialno, dealername, purchase_date, invoiceno, deliveryno, customername, " & _
			  "customericno, customeremail, customeraddress, customerpostcode, customerstate, customercity, customertel1, customertel2, customerfax " & _
			  "FROM tblonlinewarranty where refer_id = '" & request("refer_id") & "'"	
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
        end if
		rs.close
		sbutton = "Update"
		stype="editOnlineWrty"	
else
		sbutton = "Add"
		stype="addOnlineWrty"
end if

%>

                       <form name="formwarrnaty" method="post" action="action.asp?type=<%=stype%>">
                        <tr bgcolor="#475387">
                          <td colspan="3" align="left"><input name="warrantyno" type="text" id="warrantyno" value="<%=warrantyno%>" size="22" maxlength="50" />
                            <input type="hidden" name="cust_code" id="cust_code" value="<%=cust_code%>" />
                            <input type="hidden" name="refer_id" id="refer_id" value="<%=refer_id%>" /></td>
                          <td align="center"><input name="serialno" type="text" id="serialno" value="<%=serialno%>" size="15" maxlength="50" /></td>
                          <td align="center" nowrap="nowrap"><font color="#000000"><strong>
                            <input name="purchase_date" type="text" id="purchase_date" value="<%=purchase_date%>" size="10" maxlength="20" />
                            <a href="javascript:void(null)" onclick="window.dateField = document.formwarrnaty.purchase_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                          <td align="center" nowrap="nowrap"><input name="productmodel" type="text" id="productmodel" value="<%=productmodel%>" size="15" maxlength="50" />
                            <a href="javascript:popup('rm_customer_new_model.asp?searchitem=md_code&searchvalue=','cb17','scrollbars=yes,resizable=yes,width=500,height=500')"><font color="#FFFFFF">[Sel]</font></a></td>
                          <td align="center"><input name="dealername" type="text" id="dealername" value="<%=dealername%>" size="15" maxlength="100" /></td>
                          <td align="center"><input name="invoiceno" type="text" id="invoiceno" value="<%=invoiceno%>" size="10" maxlength="30" /></td>
                          <td align="center"><input name="deliveryno" type="text" id="deliveryno" value="<%=deliveryno%>" size="10" maxlength="30" /></td>
                          <td align="center"><input type="submit" name="button" id="button3" value="Submit" /></td>
                        </tr>
                        </form>
                        
                        <tr bgcolor="#475387">
                          <td align="center"><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Wrty Card No.</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Wrty Status</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>SN</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Purchased Date</strong></font></td>
                          <td width="10%" align="center"><font color="#FFFFFF"><strong>Model</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Dealer</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Dealer Invoice</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Dealer DO No.</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Action</strong></font></td>
                        </tr>
                        <%
				i = 1
				sql1 = "SELECT top 50 refer_id, warrantyno, productmodel, othermodel, serialno, dealername, purchase_date, invoiceno, deliveryno, customername, " & _
				       "customericno, customeremail, customeraddress, customerpostcode, customerstate, customercity, customertel1, customertel2, customerfax " & _
	                   "FROM tblonlinewarranty where customeremail = '" & cust_email & "' or cust_code='" & cust_code & "' order by warrantyno "
					   
					   'response.write sql1
			    set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
				
				if isdate(rs1("purchase_date")) then 
				if dateadd("m",12,rs1("purchase_date")) > date() then 
				   warrantystatus = "Under"
				else
				   warrantystatus = "Over"
				end if
				end if
%>
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="center"><strong><a href="#"><%=rs1("warrantyno")%></a></strong></td>
                          <td align="center"><%=warrantystatus%></td>
                          <td align="center"><%=rs1("serialno")%></td>
                          <td align="center"><%=chkdate(rs1("purchase_date"))%></td>
                          <td align="center"><%=rs1("productmodel")%></td>
                          <td align="center"><%=rs1("dealername")%></td>
                          <td align="center"><%=rs1("invoiceno")%></td>
                          <td align="center"><%=rs1("deliveryno")%></td>
                          <td align="center" nowrap="nowrap">
                          <input type="button" name="button9" id="button22" value="Edit" onclick="document.location.href='rm_customer_new.asp?cust_code=<%=cust_code%>&refer_id=<%=rs1("refer_id")%>#spareparts'" />
                            <input type="button" name="button9" id="button22" value="Del" onclick="javascript:confirmAction('<%=rs1("warrantyno")%>','action.asp?type=delOnlineWrty&refer_id=<%=rs1("refer_id")%>&cust_code=<%=cust_code%>')" />
</td>
                          </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                      </table></td>
                    </tr>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col">&nbsp;</td>
                    </tr>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td colspan="8" bgcolor="#E8E8E8"><strong><font size="2">Related Jobs</font></strong></td>
                        </tr>
                        <tr bgcolor="#475387">
                          <td align="center"><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Job No.</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Model</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>SN</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Job Date</strong></font></td>
                          <td width="10%" align="center"><font color="#FFFFFF"><strong> Status</strong></font></td>
                          <td align="center"><font color="#FFFFFF"><strong>Technician</strong></font></td>
                          <td align="center"><strong><font color="#FFFFFF">City</font></strong></td>
                        </tr>
                        <%				i = 1
				sql1 = "SELECT top 50 tbljob.job_id, tbljob.job_code, tbljob.job_count, tbljob.job_date, tbljob.job_cust_code, tbljob.job_cust_name, tbljob.job_cust_address, " & _
				"tbljob.job_cust_postcode, tbljob.job_cust_state, tbljob.job_cust_state_id, tbljob.job_cust_city, tbljob.job_cust_city_id, tbljob.job_cust_email,  " & _
				"tbljob.job_cust_tel1, tbljob.job_cust_tel2, tbljob.job_createddate, tbljob.job_createdby, tbljob.job_JS_receiveddate, tbljob.job_JS_receivedby,  " & _
				"tbljob.job_status, tbljob.job_purchase_date, tbljob.job_onlineWrtyNo, tbljob.job_onlineWrtyStatus, tbljob.job_type, tbljob.job_SN_no,  " & _
				"tbljob.job_Model, tbljob.job_faulty_desc, tbljob.job_reportedby, tbljob.job_appointment_date, tbljob.job_appointment_time,  " & _
				"tbljob.job_tech_code, tbljob.job_appointment_remark, tbljob.job_emailsentdate, tbljob.job_emailsent, tbljob.job_smssentdate,  " & _
				"tbljob.job_smssent, tbljob.job_tech_type, tbljob.job_tech_model, tbljob.job_tech_tax_invoice, tbljob.job_tech_SN,  " & _
				"tbljob.job_tech_faulty_reason, tbljob.job_tech_faulty_action, tbljob.job_tech_status, tbljob.job_tech_product_collectdate,  " & _
				"tbljob.job_tech_returntoCustDate, tbljob.job_actual_wrty_status, tbljob.job_wrty_photo, tbljob.job_hq_remark,  " & _
				"tbljob.job_hq_category_code, tbljob.job_hq_received_date, tbljob.job_totalPartsAmt, tbljob.job_totallabourAmt, tbljob.job_totaltransportAmt,  " & _
				"tbljob.job_totalAmt, tbljob.job_repair_date, tbljob.job_return_tech_date, tbljob.job_office_issueRemark, tbljob.job_office_supervisor,  " & _
				"tbljob.job_office_taxinvoice, tbljob.job_rcn_no, tbljob.job_rcn_Date, tbljob.job_inv_no, tbljob.job_do_no, tbltechnician.tech_name, tbltechnician.tech_tel1 " & _
				"FROM tbljob inner join tbltechnician on tbljob.job_tech_code = tbltechnician.tech_code where tbljob.job_id is not null and " & _
				"tbljob.job_cust_code = '" & cust_code & "' "
			    set rs1 = server.CreateObject("adodb.recordset")
				rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
%>
                        <tr>
                          <td align="center"><%=i%>.</td>
                          <td align="center"><strong><a href="rm_jobsheet.asp?job_code=<%=rs1("job_code")%>" target="_blank"><%=rs1("job_code")%></a></strong></td>
                          <td align="center"><%=rs1("job_Model")%></td>
                          <td align="center">&nbsp;</td>
                          <td align="center"><%=rs1("job_date")%></td>
                          <td align="center"><%=rs1("job_status")%></td>
                          <td align="center"><%=rs1("job_tech_code") & "-" & rs1("tech_name") %></td>
                          <td align="center"><%=rs1("job_cust_city")%></td>
                        </tr>
                        <%	
				i = i + 1
				rs1.movenext
				wend
				rs1.close
	
%>
                      </table></td>
                    </tr>
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
<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style3 {
            width: 220px;
        }
    </style>
</head>
<%
dim stype,groupid,groupname,actionname

set rs = server.CreateObject("adodb.recordset")
  
if request("type") = "editusers" then
	sql = "SELECT     user_id, createddate, user_type, user_name, password, fullname, staff_id, department, email, contactno, country, address1, address2, " & _
		"city, state, zipcode, user_active, accesslevel, OutletID, OutletArea, OutletRegion, Outletselection, lastlogindate, lastloginIP, log_by, log_date, log_ip, SupplierID, TransID, job_tech_code, view_cost, verify_claim,approve_stk " & _
		"FROM         tblusers where user_id = " & request("user_id")	
		'response.Write(strconnect)	 
  
   rs.Open sql,strconnect,0,1
   if not rs.EOF then 
	  user_id = rs("user_id")
	  createddate = ChkDate(rs("createddate"))
	  user_type = rs("user_type")
	  user_name = rs("user_name")
	  password = LenDecrypt(rs("user_name"),rs("password"))
	  fullname = rs("fullname")
	  staff_id = rs("staff_id")
	  department = rs("department")
	  email = rs("email")
	  contactno = rs("contactno")
	  country = rs("country")	 
	  address1 = rs("address1")
	  address2 = rs("address2")	  
	  city = rs("city")	  
	  state = rs("state")	  
	  zipcode = rs("zipcode")	  
	  user_active = rs("user_active")	
	  accesslevel = rs("accesslevel")
	  OutletID = rs("OutletID") & ",0"
	  OutletArea = rs("OutletArea") 
	  OutletRegion = rs("OutletRegion")	  
	  Outletselection = rs("Outletselection")	  
	  lastlogindate = rs("lastlogindate")
	  lastloginIP = rs("lastloginIP")
	  lastlogindate = rs("lastlogindate")
	  log_by = rs("log_by") 			
	  log_date = ChkDateTime(rs("log_date"))
	  log_ip = rs("log_ip") 
	  SupplierID = rs("SupplierID") 
	  TransID = rs("TransID") 
	  job_tech_code = rs("job_tech_code")
	  view_cost = rs("view_cost") 
      verify_claim=rs("verify_claim")
      approve_stk=rs("approve_stk")
	  stype = "editusers"	
	  actionname = "Save" 
   end if 
   rs.Close
 else
	  stype = "addusers"
	  country = "MYS"
	  user_type = request("user_type")
	  actionname = "Save" 
	  user_active = "Y"
	  OutletID = "0,0"
	  OutletArea = 0
 end if
 
 if user_type = "" then
    user_type = "Riegen"
 end if
    listOutletID = split(OutletID, ",")
	
	
function CheckOutletID(itm)
	for i = 0 to ubound(listOutletID)
		if cstr(trim(listOutletID(i))) = cstr(itm) then		      
		   CheckOutletID = " checked " 
		end if
	next
end function
%>

<script type="text/javascript">
    function confirmDel(id, del_link) {
        if (confirm("Are you sure you want to DELETE \n ID: " + id))
            location.href = del_link
    }
    function isEmpty(s) {
        return ((s == null) || (s.length == 0));
    }
    function validateUser() {

        if (isEmpty(document.forms["form1"].user_name.value)) {
            alert("Please Enter User Name.");
            document.forms["form1"].user_name.focus();
            return false;
        }

        if (isEmpty(document.forms["form1"].password.value)) {
            alert("Please Enter Password.");
            document.forms["form1"].password.focus();
            return false;
        }
		
		selectAll(document.getElementById('target_fromoutletID'),true);

    }
</script>

<script type="text/javascript">
    checked = false;
    function checkedAll(form1) {
        var aa = document.getElementById('form1');
        if (checked == false) {
            checked = true
        }
        else {
            checked = false
        }
        for (var i = 0; i < aa.elements.length; i++) {
            aa.elements[i].checked = checked;
        }
    }
	
	function swapElement(directiontype)
	{

	if (directiontype == "right")
	   {
		 fromList = "fromoutletID";
		 toList = "target_fromoutletID";
	   }
	else
	   {
		 fromList = "target_fromoutletID";
		 toList = "fromoutletID";  
	   }   
			
		
	var selectOptions = document.getElementById(fromList);
	for (var i = 0; i<selectOptions.length; i++)
	 {
		var opt = selectOptions[i];
		if (opt.selected) 
		{
			document.getElementById(fromList).removeChild(opt);
			document.getElementById(toList).appendChild(opt);
			i--;
		}
	}
	
	}
	
	
function selectAll(selectBox,selectAll)
 {   
  // have we been passed an ID    
   if (typeof selectBox == "string")
    {         
    selectBox = document.getElementById(selectBox);  
    }    
       // is the select box a multiple select box?     
       if (selectBox.type == "select-multiple") 
       { 
               for (var i = 0; i < selectBox.options.length; i++) 
               {           
                  selectBox.options[i].selected = selectAll;   
               }  
       }
    }	
</script>

<tr>
    <td>
        <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
                <td colspan="2">&nbsp;
                    
                </td>
            </tr>
            <tr>
                <td colspan="2">&nbsp;
                    
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="25%" class="titleblue1">
                                <div align="left">
                                    User Management</div>
                            </td>
                            <td width="75%">
                                <div align="right">
                                    <!-- #include file="printemail.asp" -->
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <strong></strong><font color="#FF0000"><strong>
                        <%=request("loginerr")%>
                    </strong></font>
                </td>
            </tr>
            <tr>
                <td valign="top">
                    <table border="0" cellpadding="2" cellspacing="0">
                        <tr>
                            <td colspan="2" class="bodycopy">
                                <font color="#FF0000">*</font> Required Fields
                            </td>
                        </tr>
                        <form name="form2" id="form2" method="post" action="mis_user_edit.asp?type=<%=stype%>&user_id=<%=user_id%>">
                        <tr>
                            <td colspan="2" class="bodycopy">
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>User Type:</strong>
                            </td>
                            <td>
                                <select name="user_type" onchange="document.form1.user_type.value = this.options[this.selectedIndex].value">
                                    <option value="Riegen" <%if user_type = "Riegen" then response.write " selected"%>>Riegen</option>
                                </select>
                                <input type="submit" name="Submit2" value="Submit" />
                            </td>
                        </tr>
                        </form>
                        <form action="mis_user_action.asp?act=<%=stype%>" method="post" name="form1" id="form1"
                        onsubmit="return validateUser();">
                        <tr>
                            <td class="bodycopy">
                                <strong>ID:</strong>
                            </td>
                            <td>
                                <%=user_id%>
                                (System Auto Generate)
                                <input type="hidden" name="user_type" value="<%=user_type%>" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>Created Date:</strong>
                            </td>
                            <td>
                                <%=createddate%>
                                (System Auto Generate)
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>User Active: <font color="#FF0000">*</font></strong>
                            </td>
                            <td>
                                <select name="user_active">
                                    <option value="Y" <%if user_active = "Y" then response.Write(" selected")%>>Yes</option>
                                    <option value="N" <%if user_active = "N" then response.Write(" selected")%>>No</option>
                                </select>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="16%" class="bodycopy">
                                <strong>User Name: <font color="#FF0000">*</font> </strong>
                            </td>
                            <td width="84%">
                                <%if stype = "addusers" then %>
                                <input name="user_name" type="text" class="text" id="user_name" value="<%=user_name%>"
                                    size="25" maxlength="30" />
                                <%else%>
                                <input name="user_name" type="hidden" value="<%=user_name%>" />
                                <b>
                                    <%=user_name%>
                                </b>
                                <%end if%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>Password: <font color="#FF0000">*</font></strong>
                            </td>
                            <td>
                                <input name="password" type="password" class="text" id="password2" value="<%=password%>"
                                    size="25" maxlength="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>Staff ID:</strong>
                            </td>
                            <td>
                                <input name="staff_id" type="text" class="text" id="staff_id" value="<%=staff_id%>"
                                    size="25" maxlength="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>Level:</strong>
                            </td>
                            <td>
                                <select name="accesslevel" class="text" id="select">
                                    <option value="mis" <%if accesslevel = "mis" then response.Write(" selected")%>>MIS</option>
                                    <option value="admin" <%if accesslevel = "admin" then response.Write(" selected")%>>Admin</option>
                                    <option value="finance" <%if accesslevel = "finance" then response.Write(" selected")%>>Finance</option>
                                    <option value="cs" <%if accesslevel = "cs" then response.Write(" selected")%>>Customer Service</option>
                                    <option value="technician" <%if accesslevel = "technician" then response.Write(" selected")%>>technician</option>
                                    <option value="sc" <%if accesslevel = "sc" then response.Write(" selected")%>>Service Coordinator</option>
                                    <option value="sales" <%if accesslevel = "sales" then response.Write(" selected")%>>Sales</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                          <td bgcolor="#B9B9FF" class="bodycopy"><strong>Linked to Technician</strong></td>
                          <td bgcolor="#B9B9FF"><select name="job_tech_code" id="job_tech_code">
                            <option value=""></option>
                            <%			
				sql = "SELECT tech_id, tech_code, tech_name FROM tbltechnician where tech_status = 'Y'"	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (job_tech_code) = (rs("tech_code")) then
					  response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                          </select></td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>Full Name:</strong>
                            </td>
                            <td>
                                <input name="fullname" type="text" class="auto-style3" id="fullname" value="<%=fullname%>"
                                    size="25" maxlength="50" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>Company/Department:</strong>
                            </td>
                            <td>
                                <input name="department" type="text" class="auto-style3" id="department" value="<%=department%>"
                                    size="25" maxlength="50" />
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" class="bodycopy">
                                <strong>Email:</strong>
                            </td>
                            <td nowrap="nowrap">
                                <input name="email" type="text" class="auto-style3" id="email" value="<%=email%>" size="25"
                                    maxlength="100" />
                                <span class="bodycopy">example: user@rigen.com.my</span>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" nowrap="nowrap" class="bodycopy">
                                <strong>Contact No.:</strong>
                            </td>
                            <td nowrap="nowrap">
                                <input name="contactno" type="text" class="text" id="contactno" value="<%=contactno%>"
                                    size="25" maxlength="30" />
                                <span class="bodycopy">example: 60121234567</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bodycopy">
                                <strong>Country: <font color="#FF0000">*</font></strong>
                            </td>
                            <td>
                                <select name="country">
                                <option value="Malaysia">Malaysia</option>
                                    <%
				set rs1 = server.CreateObject("adodb.recordset")
				sql1 = "SELECT id, cnty_name, cnty_id FROM  tblcountry"	
                rs1.Open sql1,strconnect,3,3,&H0001
                while Not rs1.EOF
				if country = cstr(rs1("cnty_name")) then 
				   response.write "<option value='" & rs1("cnty_name") & "' selected>" & rs1("cnty_name") & "</option>"				 							
				else
				   response.write "<option value='" & rs1("cnty_name") & "'>" & rs1("cnty_name") & "</option>"				 							
				end if
				rs1.movenext
				wend
				rs1.close
				set rs1 = nothing					
                                    %>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" class="bodycopy">
                                <strong>Address:</strong>
                            </td>
                            <td>
                                <textarea name="address1" rows="4" wrap="VIRTUAL" class="auto-style3" id="textarea"><%=address1%></textarea>
                            </td>
                        </tr>
                        <tr>
                            <td class="text">&nbsp;
                                
                            </td>
                            <td class="text">&nbsp;
                                
                            </td>
                        </tr>
                        <tr>
                            <td class="text">
                                <strong>Update By:<br />
                                </strong>
                            </td>
                            <td class="text">
                                <%=log_by%>
                                (IP Address:
                                <%=log_ip%>
                                )
                            </td>
                        </tr>
                        <tr>
                            <td class="text">
                                <strong>Update Date:<br />
                                </strong>
                            </td>
                            <td class="text">
                                <%=log_date%>
                            </td>
                        </tr>
                        <tr>
                            <td>&nbsp;
                                
                            </td>
                            <td>
                                <input type="hidden" name="oldpassword" value="<%=password%>" />
                                <input name="user_id" type="hidden" class="user_id" value="<%=user_id%>" />
                                <input name="Submit" type="submit" class="button" value="<%=actionname%>" />
                            </td>
                        </tr>
                        </form>
                    </table>
                </td>
                <td valign="top"><table border="0" cellpadding="2" cellspacing="0">
                  <tr>
                    <td colspan="2" class="bodycopy"></td>
                  </tr>
                  <form action="mis_user_action.asp?act=Settingusers" method="post" name="form5" id="form5">
                    <tr>
                      <td colspan="2" bgcolor="#475387" align="center"><strong><font color="#FFFFFF">Access Setting for Master</font></strong></td>
                    </tr>
                    <tr>
                      <td colspan="2" nowrap="nowrap" bgcolor="#F3F3F3" class="bodycopy">&nbsp;</td>
                    </tr>
                     <tr>
                      <td width="16%" nowrap="nowrap" bgcolor="#F3F3F3" class="bodycopy"><strong>Verify Claim:</strong></td>
                      <td width="84%" bgcolor="#F3F3F3"><select name="verify_claim" id="verify_claim">
                        <option value="N" <%if verify_claim="N" then response.write " selected"%>>N</option>
                        <option value="Y" <%if verify_claim="Y" then response.write " selected"%>>Y</option>                       
                        </select></td>
                    </tr> 
                    <tr>
                      <td width="16%" nowrap="nowrap" bgcolor="#F3F3F3" class="bodycopy"><strong>View Stock Cost:</strong></td>
                      <td width="84%" bgcolor="#F3F3F3"><select name="view_cost" id="view_cost">
                        <option value="N" <%if view_cost="N" then response.write " selected"%>>N</option>
                        <option value="Y" <%if view_cost="Y" then response.write " selected"%>>Y</option>                       
                        </select></td>
                    </tr>
                     <tr>
                      <td width="16%" nowrap="nowrap" bgcolor="#F3F3F3" class="bodycopy"><strong>Stock Approver:</strong></td>
                      <td width="84%" bgcolor="#F3F3F3"><select name="approve_stk" id="approve_stk">
                        <option value="N" <%if approve_stk="N" then response.write " selected"%>>N</option>
                        <option value="Y" <%if approve_stk="Y" then response.write " selected"%>>Y</option>                       
                        </select> (Stock-In/Adjustment/Transfer)</td>
                    </tr> 
                  <tr>
                      <td bgcolor="#F4F4F4"><input name="user_id" type="hidden" class="user_id" value="<%=user_id%>" /></td>
                      <td align="right" bgcolor="#F4F4F4"><input name="Submit3" type="submit" class="button" value="Update Setting" /></td>
                    </tr>
                  </form>
                </table></td>
            </tr>
            <tr>
                <td colspan="2" align="right">&nbsp;
                    
                </td>
            </tr>
            <tr>
                <td colspan="2">&nbsp;
                    
                </td>
            </tr>
        </table>
    </td>
</tr>
<!-- #include file="footer.asp" -->

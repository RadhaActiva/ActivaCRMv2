<!-- #include file="header.asp" -->
 <%
dim stype,groupid,groupname,actionname

set rs = server.CreateObject("adodb.recordset")
  
if request("type") = "editcity" then
   sql = "SELECT     ct_id, ct_cnty_id, ct_state_id, ct_state_code, ct_name FROM tblcity where ct_id = " & request("ct_id")
   rs.Open sql,strconnect,0,1
   if not rs.EOF then 
	  ct_id = rs("ct_id")
	  ct_state_id = rs("ct_state_id")
	  ct_name = rs("ct_name")
	  stype = "editcity"	
	  actionname = "Edit" 
   end if 
   rs.Close
 else
	  stype = "addcity"
	  actionname = "Add" 
 end if
%>  
<script type="text/javascript">
function confirmDel(id,del_link){
  if (confirm("Are you sure you want to DELETE \n ID: " + id))
    location.href=del_link
}
function isEmpty(s) {
  return ((s == null) || (s.length == 0));
}
function validateUser(){
if (isEmpty(document.forms["form1"].username.value)) {
    alert("Please Enter Username.");    
    document.forms["form1"].username.focus();
   return false;
   }   
   
if (isEmpty(document.forms["form1"].password.value)) {
    alert("Please Enter Password.");    
    document.forms["form1"].password.focus();
   return false;
   }   
   
if (isEmpty(document.forms["form1"].name.value)) {
    alert("Please Enter Name.");    
    document.forms["form1"].name.focus();
   return false;
   }  
   
if (isEmpty(document.forms["form1"].email.value)) {
    alert("Please Enter Email.");    
    document.forms["form1"].email.focus();
   return false;
   } 
   
if (isEmpty(document.forms["form1"].contactno.value)) {
    alert("Please Enter Contact Number.");    
    document.forms["form1"].contactno.focus();
   return false;
   }    
  
}
</script>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="981" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View</font> 
                        Master</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" cellspacing="0" cellpadding="3">
                    <tr>
                      <td width="46%" align="left" bgcolor="#CCCCCC" scope="row"><strong> Master Setting &gt; City</strong></td>
                      <td align="left" bgcolor="#CCCCCC">&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="2" align="left" valign="top" scope="row"><table width="100%" border="0" cellpadding="5" cellspacing="0">
                        <tr>
                          <td class="bodycopy"><input type="button" name="Submit3" value="Create New City" onclick="document.location.href='mis_master_city_view.asp'" />
                            <font size="4"><strong> <font color="#FF0000" size="2"><%=request("loginerr")%></font></strong></font> <font color="#FF0000"><strong> </strong></font></td>
                        </tr>
                        <tr>
                          <td valign="top" class="bodycopy"><table width="766" border="0">
                            <tr>
                              <td width="257" valign="top"><form action="mis_user_action.asp?act=<%=stype%>" method="post" name="form1" id="form1" onsubmit="return validateUser();">
                                <table width="99%" border="0">
                                  <tr>
                                    <td valign="top"><strong>State</strong></td>
                                    <td><select name="ct_state_id" id="ct_state_id"  style="width:200px">
                          <option value="0"></option>
                          <%			
				sql = "SELECT state_id, state_cnty_id, state_code, state_name FROM tblstate order by state_name"	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (ct_state_id) = (rs("state_id")) then
					  response.write "<option value='" & rs("state_id") & "' selected>" & rs("state_name") & "</option>"
					  else
					  response.write "<option value='" & rs("state_id") & "'>" & rs("state_name") & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                        </select></td>
                                  </tr>
                                  <tr>
                                    <td width="29%" valign="top"><strong>City</strong></td>
                                    <td width="71%">
                                      <input name="ct_name" type="text" id="ct_name" value="<%=ct_name%>" size="30" maxlength="100" /></td>
                                  </tr>
                                  <tr>
                                    <td>&nbsp;</td>
                                    <td align="right"><input name="ct_id" type="hidden" class="ct_id" value="<%=ct_id%>" />
                                      <input name="Submit" type="submit" class="button" value="<%=actionname%>" /></td>
                                  </tr>
                                </table>
                              </form></td>
                              <td align="center" valign="top"><table border="1" cellspacing="0" cellpadding="3">
                                <tr>
                                  <td><strong>No</strong></td>
                                  <td class="whitecopy"><strong>City</strong></td>
                                  <td class="whitecopy"><strong>State</strong></td>
                                  <td><strong>Action</strong></td>
                                </tr>
                                <%
				i = 1				
				sql = "SELECT     ct_id, ct_cnty_id, ct_state_id, ct_state_code, ct_name FROM tblcity ORDER BY ct_state_code, ct_name"	
                rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
				%>
                                <tr>
                                  <td nowrap="nowrap" class="bodycopy"><%=i%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("ct_name")%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("ct_state_code")%></td>
                                  <td nowrap="nowrap"><input name="edit" type="button" class="button" id="edit" value="Edit" onclick="document.location.href='mis_master_city_view.asp?type=editcity&ct_id=<%=rs("ct_id")%>'" />
                                    <input name="edit2" type="button" class="button" id="edit2" value="Del" onclick="javascript:confirmDel('<%=rs.Fields("ct_name")%>','mis_user_action.asp?act=delcity&ct_id=<%=rs("ct_id")%>')" /></td>
                                </tr>
                                <%
				i = i + 1
				rs.movenext
				wend
				rs.close
				set rs = nothing
				%>
                              </table></td>
                            </tr>
                          </table></td>
                        </tr>
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
<!-- #include file="header.asp" -->
<SCRIPT LANGUAGE="JavaScript">
function isEmpty(s) {
  return ((s == null) || (s.length == 0));
}

function isDigit(c) {
var valid = "0123456789."
 var temp;
 for (var i=0; i<c.length; i++) {
 temp = "" + c.substring(i, i+1);
 if (valid.indexOf(temp) == "-1") { 
  return false;
  }
 }
 return true;  
}

function validatePassword(){

var invalid = " "; // Invalid character is a space

if (isEmpty(document.forms["form1"].current_password.value)) {
    alert("Please Enter Current Password.");    
    document.forms["form1"].current_password.focus();
   return false;
   }   
   
if (isEmpty(document.forms["form1"].new_password.value)) {
    alert("Please Enter New Password.");    
    document.forms["form1"].new_password.focus();
   return false;
   }   

// check for minimum length
if (document.forms["form1"].new_password.value.length < 8) {
    alert("Password must be at least 8 characters long. Please try again.");    
    document.forms["form1"].new_password.focus();
   return false;
   } 

// check for spaces
if (document.forms["form1"].new_password.value.indexOf(invalid) > -1) {
	alert("New Password is not allowed to has spaces. Please try again.");
	document.forms["form1"].new_password.focus();
	return false;
	}

if (isDigit(document.forms["form1"].new_password.value)) {
	    alert("Password must be in alphanumeric.");    
	    document.forms["form1"].new_password.focus();    
	   return false;     
	 }
	 	 
if (isEmpty(document.forms["form1"].renew_password.value)) {
    alert("Please Re-Enter New Password.");    
    document.forms["form1"].renew_password.focus();
   return false;
   }  
   
if (document.forms["form1"].new_password.value != document.forms["form1"].renew_password.value) {
    alert("New Password doesn't match with Re-Enter Password, please try again.");    
    document.forms["form1"].new_password.focus();
   return false;
   }    

if (document.forms["form1"].current_password.value == document.forms["form1"].renew_password.value) {
    alert("Current Password is not allowed to use in New Password, please try again.");    
    document.forms["form1"].new_password.focus();
   return false;
   } 
  
}
</script>
<%
sql = "select password from tblusers where user_name = '" & Request.Cookies("GAPS")("sloginid") & "'"
set rs = server.CreateObject("adodb.recordset")
usr=Request.Cookies("GAPS")("user_type")
rs.Open sql,strconnect,0,1
if not rs.eof then
    PW_Password = rs.Fields(0)
end if
rs.Close
set rs = nothing

%>
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="mis_user_action.asp?act=UpdatePAS" onSubmit="return validatePassword();">
                    <table border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                      <tbody>
                        <tr bgcolor="#E9E9E9"> 
                          <td bgcolor="#FFFFFF" 
          scope="col"><table border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                              <tbody>
                                <tr> 
                                  <td colspan="4" bgcolor="#E8E8E8" scope="col"><strong><font size="3"> 
                                    Change Password</font></strong></td>
                                </tr>
                                <tr > 
                                  <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Current 
                                    Password </strong></font></td>
                                  <td colspan="3" align="middle"><div align="left"> 
                                      <input name="current_password" type="password" id="current_password" size="20" maxlength="50" />
                                    </div>
                                    <div align="left"> </div></td>
                                </tr>
                                <tr > 
                                  <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>New 
                                    Password </strong></font></td>
                                  <td colspan="3" align="middle"><input name="new_password" type="password" id="new_password" size="20" maxlength="50" /></td>
                                </tr>
                                <tr > 
                                  <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Re-Enter 
                                    New Password</strong></font></td>
                                  <td colspan="3" align="middle"><input name="renew_password" type="password" id="renew_password" size="20" maxlength="50" /></td>
                                </tr>
                                <tr>
                                  <td colspan="4"><strong><font color="#0000FF"><%=request("loginerr")%></font></strong></td>
                                </tr>
                                <tr> 
                                  <td colspan="4" align="right"> <input type="hidden" name="hiddenField" value="<%=PW_Password%>" /> 
                                    <input type="submit" name="Submit3" value="Update My Password" /> 
                                  </td>
                                </tr>
                              </tbody>
                          </table></td>
                        </tr>
                      </tbody>
                    </table>
                  </form>
                  <br />
                    <br />
                  </td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->

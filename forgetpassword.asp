<!-- #include file="core/database/dbconnect.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Riegen Marketing CRM</title>
<style type="text/css">
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #666;
}
body {
	background-color: #f7f7f7;
	background-image: url(images/login-page-bkg.gif);
	background-repeat: repeat-x;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}

a:link {color: #ed1d24;}
a:visited {color: #ed1d24;}
a:active {color: #ed1d24;}
a:hover {text-decoration: underline; color: #ff0000;}
a {text-decoration: none;}

.textLINK01 a:link {text-decoration: underline; color: #272727;}
.textLINK01 a:visited {text-decoration: underline; color: #272727;}
.textLINK01 a:active {text-decoration: underline; color: #272727;}
.textLINK01 a:hover {text-decoration: underline; color: #ff0000;}


#apDiv1 {
	margin-top: 50px;
	z-index:1;
}
#apDiv2 {
	margin-top: 100px;
	z-index:1;
}
#apDiv3 {
	margin-top: 20px;
	z-index:1;
}
.txt01 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 38px;
	color: #666;}
	
.txt02 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #666;}

</style>
</head>

<body>
    <div style="height:20px;"></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><table width="764" border="0" cellspacing="0" cellpadding="0">
        <tr>
     <td style="padding-left:10px; width:150px;">
   <!--   <img src="images/topleft.png" width="100" height="75" style="display:block;" alt="" />-->
    </td>

    <td style="padding:0;">
      <!--<img src="images/centrelogo.png" style="display:block; height:60px; width:100%; " alt="" />-->
    </td>

  </tr>
      <tr>
        <td><table width="764" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="12" valign="top" bgcolor="#FFFFFF"><img src="images/login-box-l.jpg" width="12" height="326" /></td>
            <td align="left" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50" align="center" valign="top">&nbsp;</td>
                <td width="450" align="left" valign="top"><form name="form2" id="form2" method="post" action="validateuser.asp?type=forgetpassword"><div id="apDiv1"><span class="txt01">Lost Your Password?</span><br />
                  <br />
                  We'll help you reset your password. Please enter the login associated with your account.<br />
                  <br />
                            (Enter your username):<br />
  <input name="username" type="password" id="username" size="20" />
  <br />
                            <strong><font color="#FF0000"> 
                            <%
							response.write ChkStringLogin(request("loginerr"))%>
                            </font></strong><br />
  <input type="submit" name="button" id="button" value="Submit" />
  <br />
  <br />
  <a href="login.asp">Back to login page</a><br />
  <br />
<br />
  If you've forgotten your username and password, <a href="mailto:cwchu@riegen.com.my">contact your admin</a> . 
                  
                  
                </div>   </form>               
                  <h3>&nbsp;</h3></td>
                </tr>
            </table></td>
            <td width="12" valign="top" bgcolor="#FFFFFF"><img src="images/login-box-r.jpg" width="12" height="326" /></td>
          </tr>
            <tr>
            <td colspan="3" valign="top">&nbsp;</td>
            </tr>
          <tr>
            <td colspan="3" align="right" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="12" align="left" scope="row">&nbsp;</td>
                <td width="348" align="left" scope="row"></td>                
              </tr>
            </table></td>
          </tr>
          <tr>
            <td colspan="3" align="right" valign="top">&nbsp;</td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

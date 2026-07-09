<!-- #include file="core/database/dbconnect.asp" -->
<html>
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
	background-color: #FFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-image: url(images/login-page-bkg.gif);
	background-repeat: repeat-x;
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
	margin-top: 70px;
	z-index:1;
}
#apDiv3 {
	margin-top: 20px;
	z-index:1;
}

#apDiv4 {
    color: dimgray;
    font-size: 30px;
    font-weight: 600;
    font-family: 'Segoe UI', Arial, sans-serif;
    letter-spacing: 0.5px;
}
.txt01 {
	font-family: Cambria, Helvetica, sans-serif;
	font-size: 85px;
	color: #666;}
	
.txt02 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #666;}

    .auto-style1 {
        width: 187px;
    }

</style>

<noscript>
 For full functionality of this site it is necessary to enable JavaScript.
 Here are the <a href="http://www.enable-javascript.com/" target="_blank">
 instructions how to enable JavaScript in your web browser</a>.
</noscript>

</head>

<body>
    <div style="height:20px;"></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td align="center"> 
             <table width="764" border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse; background:#2A52BE;">
  <tr>
     <td style="padding-left:10px; width:150px;">
      <img src="images/topleft.png" width="100" height="75" style="display:block;" alt="" />
    </td>

    <td style="padding:0;">
      <img src="images/centrelogo.png" style="display:block; height:60px; width:100%; " alt="" />
    </td>

  </tr>

      <tr>
     
        <td><table width="764" border="0" cellspacing="0" cellpadding="0">
          
          <tr>
            <td width="12" valign="top" bgcolor="#FFFFFF"><img src="images/login-box-l.jpg" width="12" height="326" /></td>
            <td align="left" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="10" align="center" valign="top">&nbsp;</td>
                <td width="450" align="left" valign="bottom" style="padding-right:20px;padding-top:35px;">
                    <img src="images/crm1.png"  width="400" height="140" style="display:block;" alt="" />
                  </td>
                <td align="left" valign="top"><div id="apDiv2">
                 <form name="login" id="form2" method="post" action="validateuser.asp?type=login">
                    <div id="apDiv4">Welcome Back<br/></div>
                    <input name="user_type" type="hidden" value="Anakku">
                    <br>
                    <br>
                    Username<br />
                    <input name="txtID" type="text" id="txtID" size="22" />
                    <br />
                    <br>
                    Password<br />
<input name="txtPASSWORD" type="password" id="txtPASSWORD" size="22"/>
<br />
                            <font color="#CC0000">
                            <%Dim err,msg

if isnumeric(Request("err")) then 
   err = CInt(Request("err"))
else
   err = 0
end if


If Not err = "" Then
  Select Case err
    Case 1
      msg = "The username or password you entered is incorrect."
    Case 2
      msg = "Your username has been disabled due to three unsuccessful login attempts"
  End Select
  Response.Write "<font face=Verdana color=#CC0000 size=1>" & msg & "</font>"
End If%>
                            </font> <br />     
                     <table border="0" cellspacing="0" cellpadding="0" class="auto-style1">
                         <tr>
                             <td> <input type="image" name="submit" src="images/login-btn.png" align="left" border="0" alt="Submit" style="width: 160px;height: 60px"/></td>
                         </tr>
                         <tr>
                             <td valign="top" align="left"> <a href="forgetpassword.asp">Forgot Password? </a></td>
                         </tr>
                         </table>
                  </form>
                </div></td>
              </tr>
            </table>
             
            <td width="12" valign="top" bgcolor="#FFFFFF"><img src="images/login-box-r.jpg" width="12" height="326" /></td>
          </tr>
          <tr>
            <td colspan="3" valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="12" align="left" scope="row">&nbsp;</td>
                <td width="348" align="left" scope="row">
                <span class="txt02" style="color:#ffffff;">
                ©2024 Riegen Trading Sdn Bhd. The information contained herein is subject to change without notice.
                </span>
                </td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td colspan="2" align="right" valign="top">&nbsp;</td>
            <td align="right" valign="top">&nbsp;</td>
            </tr>
        </table></td>
      </tr>
    </table>
    <br />
<script type="text/javascript" src="//seal.alphassl.com/SiteSeal/alpha_image_115-55_en.js"></script>
    <br /></td>
  </tr>
</table>
</body>
</html>
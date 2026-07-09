<!-- #include file="database/datastore.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<link rel="stylesheet" type="text/css" href="scripts/style/general.css">
</head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Riegen CRM</title>
<%
fromDate = now()
%>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="inc/gaps.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="inc/popup.js"></script>

<style>		
.leftMenu {	text-align: left; }		
.centerMenu { text-align: center;}		
.rightMenu { text-align: right;	}
		
a.MenuLabelLink	{ COLOR: #2E86C1;	FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuLabelLink:link { COLOR: #FFFFFF;	FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLink:visited	{ COLOR: #FFFFFF; FONT-FAMILY: Tahoma; TEXT-DECORATION: None;	}
a.MenuLabelLink:hover{ COLOR: #FFFFFF; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuLabelLinkOn {	COLOR: #2E86C1; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuLabelLinkOn:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLinkOn:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLinkOn:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuItemLink { COLOR: #2E86C1; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuItemLink:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLink:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLink:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuItemLinkOn { COLOR: #2E86C1; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuItemLinkOn:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLinkOn:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLinkOn:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
.myMenu { position: absolute; visibility: hidden; z-index: 5; }				
.myMenuLabelleft { padding: 15px 0px 15px 0px; text-align: center; }		
.myMenuLabelcenter { padding: 15px 0px 0px 0px; text-align: center; }		
.myMenuLabelright { padding: 0px 0px 0px 0px; text-align: right; }		
.myMenuItemleft { padding: 0px 0px 0px 0px; text-align: left; }		
.myMenuItemcenter { padding: 0px 0px 0px 0px; text-align: center; }		
.myMenuItemright { padding: 0px 0px 0px 0px; text-align: right; }		
		
#myTest { 
width: 900px;
padding: 0px 0px 0px 0px;
z-index: 1;
}
</style>
<script language="JavaScript1.2" src="inc/api.js" type="text/javascript">
<!-- 
-->
</script>
<script language="JavaScript1.2" src="inc/menucode.js" type="text/javascript">
<!-- 
-->
</script>

<script>
function popup(theURL,winName,features) {
  window.open(theURL,winName,'top=0,left=0,toolbar=yes,location=yes,status=yes,menubar=no,' + features);
}

function confirmAction(id,url) {
if (confirm("Are you sure you want to Confirm this: "+id)) {
    location.href=url;
  }
}

function confirmDel(id,url) {
if (confirm("Are you sure you want to delete ID: "+id)) {
    location.href=url;
  }
}

</script>
<script>
var screenwidth = 850;
var screenHeight = 650;
</script>

<%
	  if Request.Cookies("GAPS")("slevel") = "mis" then 
%>
	  <!-- #include file="header_jsmis.asp" --> 
<%	  
	  elseif Request.Cookies("GAPS")("slevel") = "cs" or Request.Cookies("GAPS")("slevel") = "sc" then
%>
	  <!-- #include file="header_jscs.asp" --> 
<%	  
	  elseif Request.Cookies("GAPS")("slevel") = "technician" then 
%>
	  <!-- #include file="header_jstechnician.asp" -->  
<%	  
	  elseif Request.Cookies("GAPS")("slevel") = "sales" then 
%>
    <!-- #include file="header_jssales.asp" -->  
<% 	   
	  end if
%>



</head>

<body>
  <table width="1004" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td width="25" valign="top" background="images/bgleft.gif"><img src="images/blank.gif" width="29" height="28" /></td>
      <td width="953"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="37%" rowspan="2"><img src="images/topleft.png" width="170" height="86" /></td>
              <td width="63%"><table width="647" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="37%" align="right" background="images/boderleftback.gif"><img src="images/boderleft.gif" width="173" height="39" /></td>
                  <td width="63%" align="right" nowrap="nowrap" background="images/bgtop1.gif" bgcolor="#b91117" class="welcome_text"><div align="right">User 
                            <font color="#666666"><%=request.Cookies("GAPS")("sloginid")%></font> | Level : <font color="#666666"><%=request.Cookies("GAPS")("slevel")%></font> | <a href="changepassword.asp">Change 
                      Pwd</a> | <a href="../validateuser.asp?type=logout"><font color="#666666">Logout</font></a>&nbsp;&nbsp; 
                    </div></td>
                  <td width="2%" background="images/bgtop1.gif" bgcolor="#b91117">&nbsp;</td>
                </tr>
              </table></td>
            </tr>
            <tr>
              <td align="right" valign="middle" background="images/topright.gif"><font size="+2" color="#0087CF"><strong>CRM One</strong></font></td>
              <td align="center" valign="middle" background="images/topright.gif"><img src="../images/topright.gif" width="10" height="97" /></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td height="38" align="left" bgcolor="#515A5A">
<script language="JavaScript1.2" type="text/javascript">
<!--
myTest.writeMenuBar();
//-->
</script></td>
        </tr>
</html>
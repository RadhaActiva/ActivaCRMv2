<!-- #include file="database/datastore.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>News Title <%=gapstype%></title>
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
		
a.MenuLabelLink	{ COLOR: #FFFFFF;	FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuLabelLink:link { COLOR: #FFFFFF;	FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLink:visited	{ COLOR: #FFFFFF; FONT-FAMILY: Tahoma; TEXT-DECORATION: None;	}
a.MenuLabelLink:hover{ COLOR: #FFFFFF; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuLabelLinkOn {	COLOR: #ffffff; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuLabelLinkOn:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLinkOn:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuLabelLinkOn:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuItemLink { COLOR: #ffffff; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuItemLink:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLink:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLink:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
a.MenuItemLinkOn { COLOR: #ffffff; FONT-SIZE: 12px;
FONT-FAMILY: Tahoma; TEXT-DECORATION: None;
margin: 0px; padding: 0px; font-weight: bold; }
a.MenuItemLinkOn:link { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLinkOn:visited { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
a.MenuItemLinkOn:hover { COLOR: #ffffff; FONT-FAMILY: Tahoma; TEXT-DECORATION: None; }
		
.myMenu { position: absolute; visibility: hidden; z-index: 5; }		
		
.myMenuLabelleft { padding: 15px 0px 0px 0px; text-align: center; }		
.myMenuLabelcenter { padding: 15px 0px 0px 0px; text-align: center; }		
.myMenuLabelright { padding: 0px 0px 0px 0px; text-align: right; }		
.myMenuItemleft { padding: 0px 0px 0px 0px; text-align: left; }		
.myMenuItemcenter { padding: 0px 0px 0px 0px; text-align: center; }		
.myMenuItemright { padding: 0px 0px 0px 0px; text-align: right; }		
		
#myTest { 
width: 800px;
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


</head>

<body>
  <table width="1004" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td valign="top"></td>
      <td width="953">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td></td>
        </tr>

<!-- #include file="database/datastore.asp" -->

<%
sql = "SELECT     news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews " & _
      "where news_id = " & Request("news_id")
set rs = server.CreateObject("adodb.recordset")	
rs.Open sql,strconnect,3,3,&H0001
if Not rs.EOF then
   news_date = chkdate(rs("news_date"))
   news_title = rs("news_title")
   news_description = rs("news_description")
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>:: News Title</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript" src="inc/popup.js"></script>
<link href="inc/gaps.css" rel="stylesheet" type="text/css" />
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<TABLE width="450" border=1 cellPadding=3 cellSpacing=0 bordercolor="#E5E5E5" id=ctl00_ContentPlaceHolder1_GV>
  <TBODY>
    <TR bgcolor="#E9E9E9"> 
      <td bgcolor="#FFFFFF" 
          scope=col>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
          <tr> 
            <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td><img src="images/arrow_btn.gif" width="18" height="17" /></td>
                  <td width="93%" class="contentheader1"><strong> <font size="3"> 
                    <%=news_date%> - <%=news_title%> </font></strong></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign="top"><strong> <%=news_description%> </strong></td>
          </tr>
        </table></td>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>

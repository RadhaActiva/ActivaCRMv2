<!-- #include file="header.asp" -->
<head>
    <style type="text/css">
        .auto-style1 {
            width: 789px;
        }
    </style>
</head>
<%
set rs = server.CreateObject("adodb.recordset")  
if request("news_id") <> "" then
	sql = "SELECT     news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews " & _
	      "where news_id = " & request("news_id")		  
else
	sql = "SELECT   top 1 news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  " & _
	      "FROM  tblNews where news_active = 'Y' and news_type = '" & gapstype & "' order by news_date desc "
end if   

   rs.Open sql,strconnect,0,1
   if not rs.EOF then 
	  news_id = rs("news_id")
	  news_type = rs("news_type")
	  news_date = ChkDate(rs("news_date"))  
	  news_grouplevel = rs("news_grouplevel")
	  news_active = rs("news_active")
	  news_title = rs("news_title")
	  news_desc_header = rs("news_desc_header")
	  news_description = rs("news_description")	 
	  log_by = rs("log_by") 			
	  log_date = ChkDateTime(rs("log_date"))
   end if 
   rs.Close 
%> 
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left">Welcome 
                          to <b>CRM One System</b></div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td width="250" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td align="left" valign="top" bgcolor="#FFFFFF">
                    <div align="left"><br />
                      <br />
                    </div>
                    </td>
                  
                <td align="left" valign="top"> <table width="95%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td><table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td><img src="images/mainpic.jpg" height="354" class="auto-style1" /></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table>
                  <br /> </td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
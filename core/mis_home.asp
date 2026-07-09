<!-- #include file="header.asp" -->
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
                        <td class="titleblue1"><div align="left"><font color="#CC0000">Reigen Trading CRM One System</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td width="250" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td align="left" valign="top" bgcolor="#FFFFFF"><table width="206" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td><img src="images/announcement_top.gif" width="206" height="13" /></td>
                      </tr>
                      <tr> 
                        <td width="206" valign="top" background="images/announcement_body.jpg"><table width="94%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <td><img src="images/arrow_btn.gif" width="18" height="17" /></td>
                                  <td class="contentheader1"><strong>Admin Page </strong></td>
                                </tr>
                                <tr> 
                                  <td>&nbsp;</td>
                                  <td class="contentheader1">&nbsp;</td>
                                </tr>
								
<%
sql = "SELECT     news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews " & _
      "where news_type = '" & gapstype & "' and news_active = 'Y' and (news_grouplevel = 'All' or news_grouplevel = '" & request.Cookies("GAPS")("slevel") & "') order by news_date desc"
set rs = server.CreateObject("adodb.recordset")	
rs.Open sql,strconnect,3,3,&H0001
While Not rs.EOF
%>
                                <tr> 
                                  <td>&nbsp;</td>
                                  <td><span class="content1"><strong><a href="javascript:popup('news_detail.asp?news_id=<%=rs("news_id")%>','news','scrollbars=yes,resizable=yes,width=650,height=500')">
                                    <%=chkdate(rs("news_date"))%>
                                    - 
                                    <%=rs("news_title")%>
                                    </a></strong></span></td>
                                </tr>
                                <tr> 
                                  <td colspan="2"><div align="center"><img src="images/announcement_line.jpg" width="162" height="3" /></div></td>
                                </tr>								
                                <tr> 
                                  <td>&nbsp;</td>
                                  <td class="contentheader1">&nbsp;</td>
                                </tr>
<%
rs.movenext
wend
rs.close
%>
                                
                              </table>
                            </td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td><img src="images/announcement_bottom.gif" width="206" height="14" /></td>
                      </tr>
                    </table>
                    <div align="left"><br />
                      <br />
                    </div>
                    </td>
                  
                <td align="left" valign="top"> <table width="95%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td><table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td><img src="images/gaps_mainpic.jpg" width="517" height="354" /></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table>
                  <br /> </td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
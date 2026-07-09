<!-- #include file="header.asp" -->

<%
dim stype,groupid,groupname,actionname

set rs = server.CreateObject("adodb.recordset")	
	sql = "SELECT     news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews "

	  if request("searchvalue") <> "all" and request("searchvalue") <> "" then
		 sql = sql & " and " & request("searchitem") & " LIKE '%" & request("searchvalue") & "%' "
	  end if		
	 
	 if request("sortby") <> "" then		  
	 sql = sql & " Order by " & request("sortby")
	 else
	 sql = sql & " Order by news_date desc"
	 end if	 			

'response.write sql
'response.End()

Response.Cookies("GAPS")("sqlexcel") = sql
            rs.Open sql,strconnect,3,3,&H0001
			If Not rs.EOF Then

			if request("rowno") <> "" then
				  row = cint(request("rowno"))
			else
				  row = 50
			end if
			
Showed = Request.QueryString("num")
If Showed = "" Then Showed = 0
TotalRecord = rs.RecordCount
Remain = TotalRecord - Showed

If Remain > row Then
  LoopMax = Showed + row
Else
  LoopMax = Showed + Remain
End If

	If Int(TotalRecord/row) <> TotalRecord/row Then
	  pgCount = Int(TotalRecord/row) + 1
	Else
	  pgCount = TotalRecord/row
	End If

	if LoopMax mod row = 0 then
		pagestartno = LoopMax/row
	else
		pagestartno = pgCount
	end if		
end if

count = count + Showed
link = "&type=searchdata&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby")
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
if (isEmpty(document.forms["form1"].user_name.value)) {
    alert("Please Enter user_name.");    
    document.forms["form1"].user_name.focus();
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
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%" class="titleblue1"><div align="left">News 
                          Management</div></td>
                        <td width="75%"><div align="right"> <!-- #include file="printemail.asp" -->
                        </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  
                <td><strong> </strong></td>
                </tr>
                <tr> 
                  <td valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                      <tbody>
                        <tr bgcolor="#E9E9E9"> 
                          <td width="792" bgcolor="#FFFFFF"><div align="left"> 
                            <form action="mis_news_view.asp?type=searchdata" method="post" name="form2" id="form2">
                              <strong>Search: 
                              <input name="searchvalue" type="text" class="txtbox" id="searchvalue" value="<%=request("searchvalue")%>" />
                              <select name="searchitem" class="txtbox" id="select">
                                <option value="news_title" <%if request("searchitem") = "news_title" then response.write " selected"%>>News Title</option>
                                <option value="news_date" <%if request("searchitem") = "news_date" then response.write " selected"%>>News Date</option>
                                <option value="news_description" <%if request("searchitem") = "news_description" then response.write " selected"%>>News Description</option>                               
                              </select>
                              <input name="searchwhat" type="submit" class="Button" id="searchwhat" value="Display" />
                              </strong><strong> 
                              <input type="button" name="Button" value="Create News" onclick="document.location.href='mis_news_edit.asp?type=addnews'"/>
                              <br />
                              Sort by: - 
                              <select name="sortby" class="txtbox" id="select">
                                <option value="news_title" <%if request("sortby") = "news_title" then response.write " selected"%>>News Title</option>
                                <option value="news_date" <%if request("sortby") = "news_date" then response.write " selected"%>>News Date</option>
                                <option value="news_description" <%if request("sortby") = "news_description" then response.write " selected"%>>News Description</option>                               
                              </select>
                              </strong> 
                            </form>
                          </div></td>
                        </tr>
                        <tr bgcolor="#E9E9E9"> 
                          
                        <td valign="top" bgcolor="#FFFFFF" 
          > <table width="100%" border="0" cellpadding="4" cellspacing="0">
                            <tr bgcolor="#666666"> 
                              <td width="6%" height="30"> <div align="center" class="whitecopy"><font color="#FFFFFF"><strong>No</strong></font></div></td>
                              <td width="14%" class="whitecopy"> <font color="#FFFFFF"><strong>News 
                                Date </strong></font></td>
                              <td width="14%" class="whitecopy"> <font color="#FFFFFF"><strong>News 
                                Type</strong></font></td>
                              <td width="15%" class="whitecopy"> <font color="#FFFFFF"><strong>News 
                                Target </strong></font></td>
                              <td width="47%" class="whitecopy"><font color="#FFFFFF"><strong>Title</strong></font></td>
                              <td width="9%" align="center" class="bodycopy"><font color="#FFFFFF"><strong>Active</strong></font></td>
                              <td width="9%" align="center"> <div class="whitecopy"><font color="#FFFFFF"><strong>Action</strong></font></div></td>
                            </tr>
                            <%
if not rs.eof then
rs.Move Showed
end if
For j = Showed + 1 To LoopMax

if count mod 2 = 0 then 
   bgcolor = "#FFFFFF"
else
   bgcolor = "#E5E5E5"
end if				
				%>
                            <tr bgcolor="<%=bgcolor%>"> 
                              <td nowrap="nowrap" class="bodycopy"> <div align="center"> 
                                  <%=j%> </div></td>
                              <td nowrap="nowrap" class="bodycopy"> 
                                <%=chkdate(rs("news_date"))%>
                              </td>
                              <td nowrap="nowrap" class="bodycopy"> <%=rs("news_type")%> </td>
                              <td nowrap="nowrap" class="bodycopy"> <%=rs("news_grouplevel")%> </td>
                              <td nowrap="nowrap" class="bodycopy"> <%=rs("news_title")%> </td>
                              <td align="center" nowrap="nowrap" class="bodycopy"> 
                                <%=rs("news_active")%> </td>
                              <td align="center" nowrap="nowrap"><strong> <a href="mis_news_edit.asp?type=editnews&news_id=<%=rs.Fields("news_id")%>">Edit</a> 
                                | <a href="javascript:confirmDel('<%=rs.Fields("news_id")%>','mis_user_action.asp?act=deletenews&news_id=<%=rs.Fields("news_id")%>')">Del</a></strong> 
                              </td>
                            </tr>
                            <%
count = count + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
				%>
                          </table></td>
                        </tr>
                      </tbody>
                    </table></td>
                </tr>
                <tr> 
                  
                <td align="right">Page <font color="3366ff"> 
                  <%=pagestartno%>
                  </font>of <font color="3366ff"> 
                  <%=pgCount%>
                  </font>: 
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='mis_news_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='mis_news_view.asp?ur_active="& Request("ur_active") &"&num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
                </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
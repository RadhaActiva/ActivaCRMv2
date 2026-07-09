<!-- #include file="header.asp" -->

<%

set rs = server.CreateObject("adodb.recordset")	
	sql = "SELECT     user_id, user_name, user_type, password, fullname, staff_id, department, email, contactno, country, address1, user_active, accesslevel, lastlogindate, " & _
        "           lastloginIP, createddate, address2, city, state, zipcode " & _
		"FROM         tblusers where user_id is not null "

	  if request("searchvalue") <> "all" and request("searchvalue") <> "" then
		 sql = sql & " and " & request("searchitem") & " LIKE '%" & request("searchvalue") & "%' "
	  end if		
	 
	 if request("sortby") <> "" then		  
	 sql = sql & " Order by " & request("sortby") & ";"
	 else
	 sql = sql & " Order by user_name;"
	 end if	 			

'response.write sql
'response.End()

Response.Cookies("GAPS")("sqlexcel") = sql
            rs.ActiveConnection = strconnect
            rs.Source = sql
            rs.CursorLocation  = 3
			rs.Open
			
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
                        <td width="25%" class="titleblue1"><div align="left">User 
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
                            <form action="mis_user_view.asp?type=searchdata" method="post" name="form2" id="form2">
                              <strong>Search: 
                              <input name="searchvalue" type="text" class="txtbox" id="searchvalue" value="<%=request("searchvalue")%>" />
                              <select name="searchitem" class="txtbox" id="select">
                                <option value="user_name" <%if request("searchitem") = "user_name" then response.write " selected"%>>User 
                                Name</option>								
								<option value="user_type" <%if request("searchitem") = "user_type" then response.write " selected"%>>User Type</option>
                                <option value="fullname" <%if request("searchitem") = "fullname" then response.write " selected"%>>Full 
                                Name</option>
                                <option value="staff_id" <%if request("searchitem") = "staff_id" then response.write " selected"%>>Staff 
                                ID</option>
                                <option value="email" <%if request("searchitem") = "email" then response.write " selected"%>>Email</option>
                                <option value="accesslevel" <%if request("searchitem") = "accesslevel" then response.write " selected"%>>Level</option>
								<option value="department" <%if request("searchitem") = "department" then response.write " selected"%>>Company/Department</option>
                              </select>
                              <input name="searchwhat" type="submit" class="Button" id="searchwhat" value="Display" />
                              </strong><strong> 
                              <input type="button" name="Button" value="Create New User" onclick="document.location.href='mis_user_edit.asp?type=addusers'"/>
                              <br />
                              Sort by: - 
                              <select name="sortby" class="txtbox" id="select">
                                <option value="user_name" <%if request("sortby") = "user_name" then response.write " selected"%>>User 
                                Name</option>
                                <option value="fullname" <%if request("sortby") = "fullname" then response.write " selected"%>>Full 
                                Name</option>
                                <option value="staff_id" <%if request("sortby") = "staff_id" then response.write " selected"%>>Staff 
                                ID</option>
                                <option value="email" <%if request("sortby") = "email" then response.write " selected"%>>Email</option>
                                <option value="accesslevel" <%if request("sortby") = "accesslevel" then response.write " selected"%>>Level</option>
								<option value="department" <%if request("sortby") = "department" then response.write " selected"%>>Company/Department</option>
                              </select>
                              </strong> 
                            </form>
                          </div></td>
                        </tr>
                        <tr bgcolor="#E9E9E9"> 
                          
                        <td valign="top" bgcolor="#FFFFFF" 
          > <table width="100%" border="0" cellpadding="4" cellspacing="0">
                            <tr bgcolor="#475387"> 
                              <td height="30"><font color="#FFFFFF"><strong>No</strong></font></td>
                              <td class="whitecopy"> <font color="#FFFFFF"><strong>User 
                                Name</strong></font></td>
                              <td class="whitecopy"> <font color="#FFFFFF"><strong>Full 
                                Name </strong></font></td>
                              <td class="whitecopy"><font color="#FFFFFF"><strong>Email</strong></font></td>
                              <td><font color="#FFFFFF"><strong>Level</strong></font></td>
                              <td><font color="#FFFFFF"><strong>Contact No</strong></font></td>
                              <td><font color="#FFFFFF"><strong>Company/Dept</strong></font></td>
                              <td class="bodycopy"><font color="#FFFFFF"><strong>User 
                                Type </strong></font></td>
                              <td align="center" class="bodycopy"><font color="#FFFFFF"><strong>Active</strong></font></td>
                              <td align="right"><font color="#FFFFFF"><strong>Action</strong></font></td>
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
                              <td nowrap="nowrap" class="bodycopy"><%=j%></td>
                              <td nowrap="nowrap" class="bodycopy"> <%=rs("user_name")%> </td>
                              <td nowrap="nowrap" class="bodycopy"> <%=rs("fullname")%> </td>
                              <td nowrap="nowrap" class="bodycopy"> <%=rs("email")%> </td>
                              <td nowrap="nowrap"><%=rs("accesslevel")%></td>
                              <td> <%=rs("contactno")%> </td>
                              <td><%=rs("department")%></td>
                              <td nowrap="nowrap" class="bodycopy"> 
                              <%=rs("user_type")%> </td>
                              <td align="center" nowrap="nowrap" class="bodycopy">
                                <%=rs("user_active")%>
                              </td>
                              <td align="right" nowrap="nowrap"><strong> <a href="mis_user_edit.asp?type=editusers&user_id=<%=rs.Fields("user_id")%>">Edit</a> 
                                | <a href="javascript:confirmDel('<%=rs.Fields("user_id")%>','mis_user_action.asp?act=deleteusers&user_id=<%=rs.Fields("user_id")%>')">Del</a></strong> 
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
					Response.Write " <a href='mis_user_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='mis_user_view.asp?ur_active="& Request("ur_active") &"&num=" & Showed+row & link & "'> Next >></a>"
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
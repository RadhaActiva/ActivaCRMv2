<!-- #include file="header.asp" -->
 <%
dim stype,groupid,groupname,actionname

set rs = server.CreateObject("adodb.recordset")
  
if request("type") = "editFaultyCode" then
   sql = "SELECT    fr_id, fr_code, fr_description, fr_type, fr_status FROM tblfaultyreason where fr_id = " & request("fr_id")
   rs.Open sql,strconnect,0,1
   if not rs.EOF then 
	  fr_id = rs("fr_id")
	  fr_code = rs("fr_code")
	  fr_description = rs("fr_description")
	  fr_type = rs("fr_type")
	  fr_status = rs("fr_status")
	  stype = "editFaultyCode"	
	  actionname = "Edit" 
   end if 
   rs.Close
 else
	  stype = "addFaultyCode"
	  actionname = "Add" 
 end if
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
if (isEmpty(document.forms["form1"].username.value)) {
    alert("Please Enter Username.");    
    document.forms["form1"].username.focus();
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
                  <td width="981" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View</font> 
                        Master</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" cellspacing="0" cellpadding="3">
                    <tr>
                      <td width="46%" align="left" bgcolor="#CCCCCC" scope="row"><strong> Master Setting &gt; Faulty Code View</strong></td>
                      <td align="left" bgcolor="#CCCCCC">&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="2" align="left" valign="top" scope="row"><table width="100%" border="0" cellpadding="5" cellspacing="0">
                        <tr>
                          <td class="bodycopy"><input type="button" name="Submit3" value="Create New Faulty Code" onclick="document.location.href='mis_master_FaultyCode_view.asp'" />
                            <font size="4"><strong> <font color="#FF0000" size="2"><%=request("loginerr")%></font></strong></font> <font color="#FF0000"><strong> </strong></font></td>
                        </tr>
                        <tr>
                          <td valign="top" class="bodycopy"><table width="766" border="0">
                            <tr>
                              <td width="257" valign="top"><form action="mis_user_action.asp?act=<%=stype%>" method="post" name="form1" id="form1" onsubmit="return validateUser();">
                                <table width="99%" border="0">
                                  <tr>
                                    <td width="29%" valign="top"><strong>Faulty Code</strong></td>
                                    <td width="71%"><%=fr_code%></td>
                                  </tr>
                                  <tr>
                                    <td valign="top"><strong>Description</strong></td>
                                    <td><input name="fr_description" type="text" id="fr_description" value="<%=fr_description%>" size="30" maxlength="50" /></td>
                                  </tr>
                                  <tr>
                                    <td valign="top"><strong>Type</strong></td>
                                    <td><select name="fr_type" id="fr_type">
                                      <option value=""></option>
                                      <option value="CF" <%if fr_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                                      <option value="WH" <%if fr_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                                    </select></td>
                                  </tr>
                                  <tr>
                                    <td valign="top"><strong>Status</strong></td>
                                    <td><label for="fr_description"><strong>
                                      <select name="fr_status" id="fr_status">
                                        <option value=""></option>
                                        <option value="Y" <%if fr_status="Y" then response.write " selected"%>>Y</option>
                                        <option value="N" <%if fr_status="N" then response.write " selected"%>>N</option>
                                      </select>
                                    </strong></label></td>
                                  </tr>
                                  <tr>
                                    <td>&nbsp;</td>
                                    <td align="right"><input name="fr_id" type="hidden" class="fr_id" value="<%=fr_id%>" />
                                      <input name="Submit" type="submit" class="button" value="<%=actionname%>" /></td>
                                  </tr>
                                </table>
                              </form></td>
                              <td align="center" valign="top"><table border="1" cellspacing="0" cellpadding="3">
                                <tr>
                                  <td><strong>No</strong></td>
                                  <td class="whitecopy"><strong>Faulty Code</strong></td>
                                  <td class="whitecopy"><strong>Description </strong></td>
                                  <td class="whitecopy"><strong>Type</strong></td>
                                  <td class="whitecopy"><strong>Status</strong></td>
                                  <td><strong>Action</strong></td>
                                </tr>
                                <%
				i = 1				
				sql = "SELECT     fr_id, fr_code, fr_description, fr_type, fr_status FROM tblfaultyreason order by fr_id"	
                rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
				%>
                                <tr>
                                  <td nowrap="nowrap" class="bodycopy"><%=i%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("fr_code")%></td>
                                  <td class="bodycopy"><%=rs("fr_description")%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("fr_type")%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("fr_status")%></td>
                                  <td nowrap="nowrap"><input name="edit" type="button" class="button" id="edit" value="Edit" onclick="document.location.href='mis_master_FaultyCode_view.asp?type=editFaultyCode&fr_id=<%=rs("fr_id")%>'" />
                                    <input name="edit2" type="button" class="button" id="edit2" value="Del" onclick="javascript:confirmDel('<%=rs.Fields("fr_id")%>','mis_user_action.asp?act=delFaultyCode&fr_id=<%=rs("fr_id")%>')" /></td>
                                </tr>
                                <%
				i = i + 1
				rs.movenext
				wend
				rs.close
				set rs = nothing
				%>
                              </table></td>
                            </tr>
                          </table></td>
                        </tr>
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
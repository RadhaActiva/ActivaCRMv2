<!-- #include file="header.asp" -->
 <%
dim stype,groupid,groupname,actionname

set rs = server.CreateObject("adodb.recordset")
  
if request("type") = "editmodeltype" then
   sql = "SELECT   catid, catcode, catname, log_by, log_date  FROM tblcategory where catid = " & request("catid")
   rs.Open sql,strconnect,0,1
   if not rs.EOF then 
	  catid = rs("catid")
	  catcode = rs("catcode")
	  catname = rs("catname")
	  stype = "editmodeltype"	
	  actionname = "Edit" 
   end if 
   rs.Close
 else
	  stype = "addmodeltype"
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
                      <td width="46%" align="left" bgcolor="#CCCCCC" scope="row"><strong> Master Setting &gt; Model Type</strong></td>
                      <td align="left" bgcolor="#CCCCCC">&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="2" align="left" valign="top" scope="row"><table width="100%" border="0" cellpadding="5" cellspacing="0">
                        <tr>
                          <td class="bodycopy"><input type="button" name="Submit3" value="Create New Model Type" onclick="document.location.href='mis_master_modeltype_view.asp'" />
                            <font size="4"><strong> <font color="#FF0000" size="2"><%=request("loginerr")%></font></strong></font> <font color="#FF0000"><strong> </strong></font></td>
                        </tr>
                        <tr>
                          <td valign="top" class="bodycopy"><table width="766" border="0">
                            <tr>
                              <td width="257" valign="top"><form action="mis_user_action.asp?act=<%=stype%>" method="post" name="form1" id="form1" onsubmit="return validateUser();">
                                <table width="99%" border="0">
                                  <tr>
                                    <td valign="top"><strong>Model Code</strong></td>
                                    <td><input name="catcode" type="text" id="catcode" value="<%=catcode%>" size="30" maxlength="50" /></td>
                                  </tr>
                                  <tr>
                                    <td width="29%" valign="top"><strong>Model Type</strong></td>
                                    <td width="71%">
                                      <input name="catname" type="text" id="catname" value="<%=catname%>" size="30" maxlength="50" /></td>
                                  </tr>
                                  <tr>
                                    <td>&nbsp;</td>
                                    <td align="right"><input name="catid" type="hidden" class="catid" value="<%=catid%>" />
                                      <input name="Submit" type="submit" class="button" value="<%=actionname%>" /></td>
                                  </tr>
                                </table>
                              </form></td>
                              <td align="center" valign="top"><table border="1" cellspacing="0" cellpadding="3">
                                <tr>
                                  <td><strong>No</strong></td>
                                  <td class="whitecopy"><strong>Model Code</strong></td>
                                  <td class="whitecopy"><strong>Model Type</strong></td>
                                  <td><strong>Action</strong></td>
                                </tr>
                                <%
				i = 1				
				sql = "SELECT   catid, catcode, catname, log_by, log_date  FROM tblcategory order by catname"	
                rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
				%>
                                <tr>
                                  <td nowrap="nowrap" class="bodycopy"><%=i%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("catcode")%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("catname")%></td>
                                  <td nowrap="nowrap"><input name="edit" type="button" class="button" id="edit" value="Edit" onclick="document.location.href='mis_master_modeltype_view.asp?type=editmodeltype&catid=<%=rs("catid")%>'" />
                                    <input name="edit2" type="button" class="button" id="edit2" value="Del" onclick="javascript:confirmDel('<%=rs.Fields("catid")%>','mis_user_action.asp?act=delmodeltype&catid=<%=rs("catid")%>')" /></td>
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
<!-- #include file="header.asp" -->
 <%
dim stype,groupid,groupname,actionname

set rs = server.CreateObject("adodb.recordset")
  
if request("type") = "editFaultyRepair" then
   sql = "SELECT ra_id, ra_type, ra_category, ra_repairaction FROM tbljob_repairaction where ra_id = " & request("ra_id")
   rs.Open sql,strconnect,0,1
   if not rs.EOF then 
	  ra_id = rs("ra_id")
	  ra_type = rs("ra_type")
	  ra_category = rs("ra_category")
	  ra_repairaction = rs("ra_repairaction")
	  stype = "editFaultyRepair"	
	  actionname = "Edit" 
   end if 
   rs.Close
 else
	  stype = "addFaultyRepair"
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
                      <td width="46%" align="left" bgcolor="#CCCCCC" scope="row"><strong> Master Setting &gt; Faulty Reason View</strong></td>
                      <td align="left" bgcolor="#CCCCCC">&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="2" align="left" valign="top" scope="row"><table width="100%" border="0" cellpadding="5" cellspacing="0">
                        <tr>
                          <td class="bodycopy"><input type="button" name="Submit3" value="Create New Faulty Reason" onclick="document.location.href='mis_master_faultyreason_view.asp'" />
                            <font size="4"><strong> <font color="#FF0000" size="2"><%=request("loginerr")%></font></strong></font> <font color="#FF0000"><strong> </strong></font></td>
                        </tr>
                        <tr>
                          <td valign="top" class="bodycopy"><table width="766" border="0">
                            <tr>
                              <td width="257" valign="top"><form action="mis_user_action.asp?act=<%=stype%>" method="post" name="form1" id="form1" onsubmit="return validateUser();">
                                <table width="99%" border="0">
                                  <tr>
                                    <td width="29%" valign="top"><strong>Faulty Reason ID</strong></td>
                                    <td width="71%"><%=fr_code%></td>
                                  </tr>
                                  <tr>
                                    <td valign="top"><strong>Repair Type</strong></td>
                                    <td><select name="ra_type" id="ra_type">
                                      <option value=""></option>
                                      <option value="CF" <%if ra_type="CF" then response.write " selected"%>>CF-Ceiling Fan</option>
                                      <option value="WH" <%if ra_type="WH" then response.write " selected"%>>WH-Water Heater</option>
                                    </select></td>
                                  </tr>
                                  <tr>
                                    <td valign="top"><strong>Repair Category</strong></td>
                                    <td><input name="ra_category" type="text" id="ra_category" value="<%=ra_category%>" size="30" maxlength="50" /></td>
                                  </tr>
                                  <tr>
                                    <td valign="top"><strong>Repair Description</strong></td>
                                    <td><textarea name="ra_repairaction" cols="30" rows="3" id="ra_repairaction"><%=ra_repairaction%></textarea></td>
                                  </tr>
                                  <tr>
                                    <td>&nbsp;</td>
                                    <td align="right"><input name="ra_id" type="hidden" class="ra_id" value="<%=ra_id%>" />
                                      <input name="Submit" type="submit" class="button" value="<%=actionname%>" /></td>
                                  </tr>
                                </table>
                              </form></td>
                              <td align="center" valign="top"><table border="1" cellspacing="0" cellpadding="3">
                                <tr>
                                  <td><strong>No</strong></td>
                                  <td class="whitecopy"><strong>Faulty Reason ID</strong></td>
                                  <td class="whitecopy"><strong>Repair Category</strong></td>
                                  <td class="whitecopy"><strong>Repair Type</strong></td>
                                  <td class="whitecopy"><strong>Repair Description</strong></td>
                                  <td><strong>Action</strong></td>
                                </tr>
                                <%
				i = 1				
				sql = "SELECT ra_id, ra_type, ra_category, ra_repairaction FROM tbljob_repairaction order by ra_type, ra_category"	
                rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
				%>
                                <tr>
                                  <td nowrap="nowrap" class="bodycopy"><%=i%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("ra_id")%></td>
                                  <td class="bodycopy"><%=rs("ra_category")%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("ra_type")%></td>
                                  <td nowrap="nowrap" class="bodycopy"><%=rs("ra_repairaction")%></td>
                                  <td nowrap="nowrap"><input name="edit" type="button" class="button" id="edit" value="Edit" onclick="document.location.href='mis_master_faultyreason_view.asp?type=editFaultyRepair&ra_id=<%=rs("ra_id")%>'" />
                                    <input name="edit2" type="button" class="button" id="edit2" value="Del" onclick="javascript:confirmDel('<%=rs.Fields("ra_id")%>','mis_user_action.asp?act=delFaultyRepair&ra_id=<%=rs("ra_id")%>')" /></td>
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
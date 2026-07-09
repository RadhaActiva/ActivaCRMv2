<!-- #include file="header.asp" -->
<!-- #include file="newsletterfunction.asp" -->
<script language="JavaScript" src="scripts/innovaeditor.js"></script>

<%
dim stype,groupid,groupname,actionname

set rs = server.CreateObject("adodb.recordset")
  
if request("type") = "editnews" then
	sql = "SELECT     news_type, news_grouplevel, news_title, news_date, news_desc_header, news_description, news_active, log_by, log_date, news_id  FROM  tblNews " & _
	      "where news_id = " & request("news_id")		  
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
	  	  
	  stype = "editnews"	
	  actionname = "Update News" 
   end if 
   rs.Close
 else
	  stype = "addnews"	 
	  news_date = ChkDate(date()) 
	  actionname = "Add News" 
	  news_active = "Y"
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
 
 if (isEmpty(document.forms["form1"].user_name.value)) {
    alert("Please Enter User Name.");    
    document.forms["form1"].user_name.focus();
   return false;
   }    
     
if (isEmpty(document.forms["form1"].password.value)) {
    alert("Please Enter Password.");    
    document.forms["form1"].password.focus();
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
                        <td width="75%"><div align="right"><!-- #include file="printemail.asp" --></div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  
                <td><strong></strong><font color="#FF0000"><strong> 
                        <%=request("loginerr")%> </strong> </font></td>
                </tr>
                <tr> 
                  
                <td valign="top"><form action="mis_user_action.asp?act=<%=stype%>" method="post" name="form1" id="form1" onsubmit="return validateUser();">
                    <table width="99%" border="0" cellpadding="2" cellspacing="0">
                      <tr> 
                        <td colspan="2" class="bodycopy"><font color="#FF0000">*</font> 
                          Required Fields</td>
                      </tr>
                      <tr> 
                        <td width="16%" class="bodycopy"><strong>News ID:</strong></td>
                        <td width="84%"> <%=news_id%> (System Auto Generate)</td>
                      </tr>
                      <tr> 
                        <td class="bodycopy"><strong>Created Date:</strong></td>
                        <td> <font color="#000000"><strong> 
                          <input name="news_date" type="text" id="news_date" value="<%=news_date%>" size="12" />
                          <a href="javascript:void(null)" onclick="window.dateField = document.form1.news_date;
             calendar = window.open('date.asp',null,'WIDTH=185,HEIGHT=203,status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=1,top=0,right=0')"><img src="images/calender.gif" width="20" height="19" border="0" align="absmiddle" /></a></strong></font></td>
                      </tr>
                      <tr> 
                        <td class="bodycopy"><strong>Publish: <font color="#FF0000">*</font></strong></td>
                        <td><input name="news_active" type="radio" class="bodycopy" value="Y" <%if news_active = "Y" then response.Write(" checked")%> /> 
                          <span class="bodycopy">Yes 
                          <input name="news_active" type="radio" class="bodycopy" value="N" <%if news_active = "N" then response.Write(" checked")%> />
                          No </span></td>
                      </tr>
                      <tr> 
                        <td class="bodycopy"><strong>Type:</strong></td>
                        <td><input name="news_type" type="radio" class="bodycopy" value="baby" <%if news_type = "baby" then response.Write(" checked")%> />
                          <span class="bodycopy">Baby 
                          <input name="news_type" type="radio" class="bodycopy" value="lingerie" <%if news_type = "lingerie" then response.Write(" checked")%> />
                          Lingerie</span></td>
                      </tr>
                      <tr> 
                        <td class="bodycopy"><strong>Level:</strong></td>
                        <td><select name="news_grouplevel" class="text" id="select">
                            <option value="all" <%if news_grouplevel = "alll" then response.Write(" selected")%>>All 
                            Level</option>
                            <option value="mis" <%if news_grouplevel = "mis" then response.Write(" selected")%>>MIS</option>
                            <option value="global1" <%if news_grouplevel = "global1" then response.Write(" selected")%>>Global1</option>
                            <option value="global2" <%if news_grouplevel = "global2" then response.Write(" selected")%>>Global2</option>
                            <option value="outlet1" <%if news_grouplevel = "outlet1" then response.Write(" selected")%>>Outlet1</option>
                            <option value="outlet2" <%if news_grouplevel = "outlet2" then response.Write(" selected")%>>Outlet2</option>
                            <option value="supplier" <%if news_grouplevel = "supplier" then response.Write(" selected")%>>Supplier</option>
                            <option value="transporter" <%if news_grouplevel = "transporter" then response.Write(" selected")%>>Transporter</option>
                            <option value="finance" <%if news_grouplevel = "finance" then response.Write(" selected")%>>Finance</option>
                            <option value="management" <%if news_grouplevel = "management" then response.Write(" selected")%>>Management</option>
                          </select></td>
                      </tr>
                      <tr> 
                        <td class="bodycopy"><strong>News Title:</strong></td>
                        <td><input name="news_title" type="text" class="text" id="news_title" value="<%=news_title%>" size="70"></td>
                      </tr>
                      <tr> 
                        <td colspan="2" valign="top" class="bodycopy"><strong>Description:</strong></td>
                      </tr>
                      <tr> 
                        <td colspan="2" class="text"> <textarea name="news_description" cols="50" rows="5" wrap="VIRTUAL" class="text" id="textarea2"><%=news_description%></textarea> 
                          <script>
		var oEdit1 = new InnovaEditor("oEdit1");
		oEdit1.width=775;//You can also use %, for example: oEdit1.width="100%"
		oEdit1.height=350;
		
		oEdit1.btnPrint=true;		
		oEdit1.btnPasteText=false;		
		oEdit1.btnPasteWord = false;
		oEdit1.btnBookmark = false;		
        oEdit1.btnForm = false;
		oEdit1.btnRemoveFormat = false;
		oEdit1.btnStyles = false;		
		oEdit1.btnFlash=false;
		oEdit1.btnMedia=false;
		oEdit1.btnLTR=false;
		oEdit1.btnRTL=false;
		oEdit1.btnSpellCheck=false;
		oEdit1.btnStrikethrough=false;
		oEdit1.btnSuperscript=false;
		oEdit1.btnSubscript=false;
		oEdit1.btnClearAll=true;
		oEdit1.btnSave=true;
		oEdit1.btnStyles=false; //Show "Styles/Style Selection" button
	
		oEdit1.css="inc/gaps.css"; //Specify external css file here

		oEdit1.cmdAssetManager = "modalDialogShow('/gaps/baby/assetmanager/assetmanager.asp',640,465)"; //Command to open the Asset Manager add-on.
		
		oEdit1.customColors=["#ff4500","#ffa500","#808000","#4682b4","#1e90ff","#9400d3","#ff1493","#a9a9a9"];//predefined custom colors
				
		oEdit1.mode="XHTMLBody";		
		
		oEdit1.REPLACE("news_description");
	</script> </td>
                      </tr>
                      <tr> 
                        <td colspan="2" class="text">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="text"><strong>Update By:<br />
                          </strong></td>
                        <td class="text"> <%=log_by%> </td>
                      </tr>
                      <tr> 
                        <td class="text"><strong>Update Date:<br />
                          </strong></td>
                        <td class="text"> <%=log_date%> </td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td> <input name="news_id" type="hidden" class="news_id" value="<%=news_id%>" /> 
                          <input name="Submit" type="submit" class="button" value="<%=actionname%>" /> 
                        </td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr> 
                  
                <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td>
				  
				  
				  </td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
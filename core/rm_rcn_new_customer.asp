<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
%>
<html>
<head>
<!-- #include file="meta.asp" -->
    <style type="text/css">
        .auto-style1 {
            height: 35px;
        }
    </style>
</head>

<body>

<%
i = 1
sql = "SELECT cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, " & _
		"cust_type, cust_status, cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id,tblcountry.cnty_name, cust_email,  " & _
		"cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic " & _
		"FROM tblcustomer inner join tblcountry on tblcustomer.cust_cnty_id = tblcountry.cnty_id "  & _
        "WHERE tblcustomer.cust_id is not null and " & searchitem & " like '%" & searchvalue& "%' " & _
		"order by tblcustomer.cust_name"

'response.write sql

set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if rs.eof then
   norecord = "There is no record found."
end if

If Not rs.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby")
%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Customer List</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_rcn_new_customer.asp">
        <select name="searchitem">                   
		  <option value="tblcustomer.cust_code" <%if searchitem = "tblcustomer.cust_code" then response.write " selected"%>>Customer 
          Code</option>
		   <option value="tblcustomer.cust_name" <%if searchitem = "tblcustomer.cust_name" then response.write " selected"%>>Customer 
          Name</option>
          <!--<option value="tblcustomer.cust_status" <%if searchitem = "tblcustomer.cust_status" then response.write " selected"%>>Status</option>-->
          <option value="tblcustomer.cust_reg_no" <%if searchitem = "tblcustomer.cust_reg_no" then response.write " selected"%>>Company 
          Registration</option>
          <option value="tblcustomer.cust_address" <%if searchitem = "tblcustomer.cust_address" then response.write " selected"%>>Address</option>
          <!--<option value="tblcustomer.cust_city" <%if searchitem = "tblcustomer.cust_city" then response.write " selected"%>>City</option>-->
          <option value="tblcustomer.cust_tel1" <%if searchitem = "tblcustomer.cust_tel1" then response.write " selected"%>>Tel 1</option>
       <!--   <option value="tblcustomer.cust_tel2" <%if searchitem = "tblcustomer.cust_tel2" then response.write " selected"%>>Tel 2</option>
          <option value="tblcustomer.cust_email" <%if searchitem = "tblcustomer.cust_email" then response.write " selected"%>>Email</option>-->
        </select>
        <input type="text" name="searchvalue" value="<%=searchvalue%>">
        <input type="submit" name="Submit" value="Submit">
      </form></td>
  </tr>
  <tr> 
    <td align="right" valign="top"><strong>Page</strong> <font color="3366ff"> 
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
					Response.Write " <a href='rm_rcn_new_customer.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rcn_new_customer.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#2B7FFF"> 
          <td width="3%" style="color:#F1F5F9;"><strong>No.</strong></td>
          <td width="12%" style="color:#F1F5F9;"><strong>Code</strong></td>
          <td width="12%" style="color:#F1F5F9;"><strong>Name</strong></td>
          <td style="color:#F1F5F9;"><strong>Address</strong></td>
          <td width="7%" style="color:#F1F5F9;"><strong>Contact</strong></td>
          <td width="7%" style="color:#F1F5F9;"><strong>Remark/PIC</strong></td>
        </tr>
        <% 
Set rs1 = Server.CreateObject("ADODB.Recordset")

if not rs.eof then
rs.Move Showed
count = Showed + 1
end if

For j = Showed + 1 To LoopMax

if i mod 2 = 0 then
	nbgcolor = "#F3F3F3"
else
	nbgcolor = "#FFFFFF"
end if

%>
        <tr valign="top" bgcolor="<%=nbgcolor%>"> 
          <td><%=count%>.</td>
          <td>
          <a href="javascript:parent.opener.document.forms['formorder'].rcn_cust_code.value='<%=rs("cust_code")%>';parent.opener.document.forms['formorder'].rcn_cust_name.value='<%=rs("cust_name")%>';parent.opener.document.forms['formorder'].rcn_cust_address.value='<%=rs("cust_address")%>';parent.opener.document.forms['formorder'].rcn_cust_postcode.value='<%=rs("cust_postcode")%>';parent.opener.document.forms['formorder'].rcn_cust_state_id.value='<%=rs("cust_state_id")%>';parent.opener.document.forms['formorder'].rcn_cust_city_id.value='<%=rs("cust_city_id")%>';parent.opener.document.forms['formorder'].rcn_cust_city.value='<%=rs("cust_city")%>';parent.opener.document.forms['formorder'].rcn_cust_state.value='<%=rs("cust_state")%>';parent.opener.document.forms['formorder'].rcn_cust_email.value='<%=rs("cust_email")%>';parent.opener.document.forms['formorder'].rcn_cust_tel1.value='<%=rs("cust_tel1")%>';parent.opener.document.forms['formorder'].rcn_cust_tel2.value='<%=rs("cust_tel2")%>';parent.opener.document.forms['formorder'].rcn_remark.value='<%=rs("cust_attention")%> <%=rs("cust_pic")%>';parent.opener.document.forms['formorder'].rcn_cust_cnty_id.value='<%=rs("cust_cnty_id")%>';parent.window.close();"><%=rs("cust_code")%></a><br>
          </td>
          <td><%=rs("cust_name")%><br>            <br></td>
          <td><%=rs("cust_address")%>&nbsp; <%=rs("cust_postcode")%>&nbsp; <%=rs("cust_city")%>&nbsp; <%=rs("cust_state")%>&nbsp; <%=rs("cnty_name")%></td>
          <td><%=rs("cust_tel1")%><br>
            <%=rs("cust_tel2")%><br>
          <%=rs("cust_email")%></td>
          <td><%=rs("cust_attention")%> <%=rs("cust_pic")%></td>
        </tr>
        <%
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
      </table></td>
  </tr>
  <tr valign="top"> 
    <td colspan="9" align="right"><strong>Page</strong> <font color="3366ff"> 
      <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>: 
      <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_rcn_new_customer.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_rcn_new_customer.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

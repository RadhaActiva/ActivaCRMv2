<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
i = 1

sql = "SELECT tblcustomer.cust_id,  " & _
		"tblcustomer.cust_code, tblcustomer.cust_name, " & _
		"tblcustomer.cust_type, tblcustomer.cust_status, tblcustomer.cust_reg_no, tblcustomer.cust_company, tblcustomer.cust_address, tblcustomer.cust_postcode,  " & _
		"tblcustomer.cust_state, tblcustomer.cust_state_id, tblcustomer.cust_city, tblcustomer.cust_city_id,tblcustomer.cust_cnty_id, tblcustomer.cust_email,  " & _
		"tblcustomer.cust_tel1, tblcustomer.cust_tel2, tblcustomer.cust_fax, tblcustomer.cust_website, tblcustomer.cust_password, tblcustomer.cust_gstregno,  " & _
        "tblcustomer.cust_attention, " & _
		"tblonlinewarranty.warrantyno, tblonlinewarranty.serialno, tblonlinewarranty.productmodel, tblonlinewarranty.dealername, tblonlinewarranty.purchase_date, " & _
		"job_total_job = ( select count(job_id) from tbljob where tbljob.job_cust_code = tblcustomer.cust_code) " & _
		"FROM tblcustomer left join tblonlinewarranty on tblcustomer.cust_email=tblonlinewarranty.customeremail " & _
	   "WHERE tblcustomer.cust_id is not null and " & searchitem & " like '%" & searchvalue& "%' " & _
		"order by tblcustomer.cust_name"

set rs1 = server.CreateObject("adodb.recordset")
rs1.ActiveConnection = strconnect
rs1.Source = sql
rs1.CursorLocation  = 3
rs1.Open
if rs1.eof then
   norecord = "There is no record found."
end if

If Not rs1.EOF Then

if request("rowno") <> "" then
	  row = cint(request("rowno"))
else
	  row = 50
end if
			
Showed = Request("num")
If Showed = "" Then Showed = 0
TotalRecord = rs1.RecordCount
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
    <td class="style21"><form name="form1" method="post" action="rm_job_new_customer.asp">
        <select name="searchitem">   
		   <option value="tblcustomer.cust_name" <%if searchitem = "tblcustomer.cust_name" then response.write " selected"%>>Customer 
          Name</option>
           <option value="tblcustomer.cust_code" <%if searchitem = "tblcustomer.cust_code" then response.write " selected"%>>Customer 
          Code</option>
       <!--   <option value="tblcustomer.cust_status" <%if searchitem = "tblcustomer.cust_status" then response.write " selected"%>>Status</option>
          <option value="tblcustomer.cust_reg_no" <%if searchitem = "tblcustomer.cust_reg_no" then response.write " selected"%>>Company -->
          Registration</option>
          <option value="tblcustomer.cust_address" <%if searchitem = "tblcustomer.cust_address" then response.write " selected"%>>Address</option>
          <!--<option value="tblcustomer.cust_city" <%if searchitem = "tblcustomer.cust_city" then response.write " selected"%>>City</option>-->
          <option value="tblcustomer.cust_tel1" <%if searchitem = "tblcustomer.cust_tel1" then response.write " selected"%>>Telephone</option>
          <!--<option value="tblcustomer.cust_tel2" <%if searchitem = "tblcustomer.cust_tel2" then response.write " selected"%>>Tel 2</option>
          <option value="tblcustomer.cust_email" <%if searchitem = "tblcustomer.cust_email" then response.write " selected"%>>Dealer</option>-->
        </select>
        <input type="text" size="40" name="searchvalue" value="<%=searchvalue%>">
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
					Response.Write " <a href='rm_job_new_customer.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_new_customer.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If%>

    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#2B7FFF" style="color:#F1F5F9;"> 
          <td width="5%" style="color:#F1F5F9;"><strong>No.</strong></td>
          <td width="9%" style="color:#F1F5F9;"><strong> Code</strong></td>
          <td width="14%" style="color:#F1F5F9;"><strong>Name</strong></td>
          <td width="24%" style="color:#F1F5F9;"><strong>Address</strong></td>
          <td width="13%" class='tktTotals' style="color:#F1F5F9;"><strong>Contact</strong></td>
          <td width="16%" class='tktTotals' style="color:#F1F5F9;"><strong>Remark/PIC</strong></td>
          <td width="19%" class='tktTotals' style="color:#F1F5F9;"><strong>Wrty No. / SN / Model / Dealer / Purchased Date</strong></td>
        </tr>
        <% 
'Set rs1 = Server.CreateObject("ADODB.Recordset")

if not rs1.eof then
rs1.Move Showed
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
               <% sql1 = "select cnty_name from tblcountry where cnty_id =" & rs1("cust_cnty_id") 
		         country=selectid(sql1) 
                %>
           <a href="javascript:parent.opener.document.forms['formorder'].job_cust_code.value='<%=rs1("cust_code")%>';parent.opener.document.getElementById('job_total_job').innerHTML='<%=rs1("job_total_job")%>';parent.opener.document.forms['formorder'].job_cust_name.value='<%=rs1("cust_name")%>';parent.opener.document.forms['formorder'].job_cust_email.value='<%=rs1("cust_email")%>';parent.opener.document.forms['formorder'].job_cust_address.value='<%=rs1("cust_address")%>';parent.opener.document.forms['formorder'].job_cust_state.value='<%=rs1("cust_state")%>';parent.opener.document.forms['formorder'].job_cust_state_id.value='<%=rs1("cust_state_id")%>';parent.opener.document.forms['formorder'].job_cust_postcode.value='<%=rs1("cust_postcode")%>';parent.opener.document.forms['formorder'].job_cust_city_code.value='<%=rs1("cust_city_id")%>';parent.opener.document.forms['formorder'].job_cust_city.value='<%=rs1("cust_city")%>';parent.opener.document.forms['formorder'].job_cust_tel1.value='<%=rs1("cust_tel1")%>';parent.opener.document.forms['formorder'].job_cust_tel2.value='<%=rs1("cust_tel2")%>';parent.opener.document.forms['formorder'].job_cust_cnty_id.value='<%=rs1("cust_cnty_id")%>';parent.window.close();">
           <%=rs1("cust_code")%></a><br>
        <!--  <%=rs1("cust_status")%>  -->          
          </td>
          <td><%=rs1("cust_name")%><br><br></td>
          <td><%=rs1("cust_address")%>&nbsp; <%=rs1("cust_postcode")%>&nbsp; <%=rs1("cust_city")%>&nbsp; <%=rs1("cust_state")%>&nbsp;<%=country%></td>
          <td><%=rs1("cust_tel1")%><br>
            <%=rs1("cust_tel2")%><br>
          <%=rs1("cust_email")%></td>
          <td><%=rs1("cust_attention")%> </td>
          <td><%=rs1("warrantyno")%>/<br>
            <%=rs1("serialno")%>/<br>
            <%=rs1("productmodel")%>/<br>
            <%=rs1("dealername")%>/<br>
            <%=chkdate(rs1("purchase_date"))%>
            </td>
        </tr>
        <%
count = count + 1 
i = i + 1
rs1.MoveNext
Next
rs1.Close
Set rs1 = Nothing
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
					Response.Write " <a href='rm_job_new_customer.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_new_customer.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

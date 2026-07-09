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
sql = "SELECT do_id, do_no, do_status, do_date, do_inv_no, do_inv_date, do_cust_code, do_cust_name, do_cust_address, do_cust_postcode, do_cust_state, " & _ 
		"do_cust_state_id, do_cust_city, do_cust_city_id, do_cust_email, do_cust_tel1, do_cust_tel2, do_createddate, do_createdby, do_job_code, do_tech_code,  " & _
		"do_totalqty, do_totalPartsAmt, do_remark, do_labourAmt, do_transportAmt, do_gstAmt, do_totalAmt, do_emailsent, do_emailsentdate, do_deliveredby,  " & _
		"do_delivereddate, do_doneby, do_donedate, do_postedby, do_posteddate, do_cancelledby, do_cancelleddate, do_purchase_date, do_onlineWrtyNo,  " & _
		"do_onlineWrtyStatus, do_SN_no, do_type, do_Model, do_model_desc, do_appointment_date, do_appointment_time, do_appointment_remark " & _
		"FROM tbldo where do_id is not null " 

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

sql = sql & " order by do_id"

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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&formname=" & formname & "&fieldname=" & fieldname & "&fieldname1=" & fieldname1

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>DO List</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_cn_dono.asp">
      <select name="searchitem">                   
        <option value="tbldo.do_no"  <% if searchitem = "tbldo.do_no" then response.write " selected" %>>DO No</option>
        <option value="tbldo.do_cust_name" <% if searchitem = "tbldo.do_cust_name" then response.write " selected" %>>Customer Name</option>
        <option value="tbldo.do_cust_tel1" <% if searchitem = "tbldo.do_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
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
					Response.Write " <a href='rm_cn_dono.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cn_dono.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>DO No</strong></td>
          <td width="12%"><strong> DO Status</strong></td>
          <td width="12%"><strong>DO Date</strong></td>
          <td width="12%"><strong>DO Qty</strong></td>
          <td width="14%" class='tktTotals'><strong>Customer Name</strong></td>
          <td width="14%" class='tktTotals'><strong>Customer Mobile</strong></td>
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
          <td nowrap><%=count%>.</td>
          <td><a href="javascript:parent.opener.document.forms['formorder'].cn_job_code.value='<%=rs("do_job_code")%>';parent.opener.document.forms['formorder'].cn_do_no.value='<%=rs("do_no")%>';parent.opener.document.forms['formorder'].cn_invoice_no.value='<%=rs("do_inv_no")%>';parent.window.close();"><%=rs("do_no")%></a></td>
          <td><%=rs("do_status")%></td>
          <td><%=chkdate(rs("do_date"))%></td>
          <td><%=rs("do_totalqty")%></td>
          <td><%=rs("do_cust_name")%></td>
          <td><%=rs("do_cust_tel1")%></td>
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
					Response.Write " <a href='rm_cn_dono.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cn_dono.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

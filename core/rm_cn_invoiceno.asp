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
sql = "SELECT tblinvoice.inv_id, tblinvoice.inv_no, tblinvoice.inv_date, tblinvoice.inv_cust_code, tblinvoice.inv_cust_name, tblinvoice.inv_cust_address, tblinvoice.inv_cust_postcode, " & _
"tblinvoice.inv_cust_state, tblinvoice.inv_cust_state_id, tblinvoice.inv_cust_city, tblinvoice.inv_cust_city_id, tblinvoice.inv_cust_email, tblinvoice.inv_cust_tel1, tblinvoice.inv_cust_tel2, " & _
"tblinvoice.inv_createddate, tblinvoice.inv_createdby, tblinvoice.inv_job_code, tblinvoice.inv_tech_code, tblinvoice.inv_totalqty, tblinvoice.inv_totalPartsAmt, tblinvoice.inv_labourAmt, " & _
"tblinvoice.inv_transportAmt, tblinvoice.inv_gstAmt, tblinvoice.inv_gstRate, tblinvoice.inv_gstCode, tblinvoice.inv_totalAmt, tblinvoice.inv_emailsent, tblinvoice.inv_emailsentdate, tblinvoice.inv_status, inv_approvedby, inv_approveddate, " & _
"tbltechnician.tech_name, tbltechnician.tech_tel1 FROM tblinvoice left join tbltechnician on tblinvoice.inv_tech_code = tbltechnician.tech_code where tblinvoice.inv_id is not null " 


if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

sql = sql & " order by tblinvoice.inv_id"

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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue 

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Invoice List</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_cn_invoiceno.asp">
      <select name="searchitem">                   
        <option value="tblinvoice.inv_no"  <% if searchitem = "tblinvoice.inv_no" then response.write " selected" %>>Invoice No</option>
        <option value="tblinvoice.inv_cust_name" <% if searchitem = "tblinvoice.inv_cust_name" then response.write " selected" %>>Customer Name</option>
        <option value="tblinvoice.inv_cust_tel1" <% if searchitem = "tblinvoice.inv_cust_tel1" then response.write " selected" %>>Customer Tel 1</option>
        <option value="tblinvoice.inv_cust_email" <% if searchitem = "tblinvoice.inv_cust_email" then response.write " selected" %>>Customer Email</option>
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
					Response.Write " <a href='rm_cn_invoiceno.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cn_invoiceno.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>Invoice No</strong></td>
          <td width="12%"><strong> Invoice Status</strong></td>
          <td width="12%"><strong>Invoice Date</strong></td>
          <td width="12%"><strong>Invoice Amt</strong></td>
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
          <td><a href="javascript:parent.opener.document.forms['forminvoice'].cn_invoice_no.value='<%=rs("inv_no")%>';parent.window.close();"><%=rs("inv_no")%></a></td>
          <td><%=rs("inv_status")%></td>
          <td><%=chkdate(rs("inv_date"))%></td>
          <td><%=rs("inv_totalAmt")%></td>
          <td><%=rs("inv_cust_name")%></td>
          <td><%=rs("inv_cust_tel1")%></td>
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
					Response.Write " <a href='rm_cn_invoiceno.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_cn_invoiceno.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

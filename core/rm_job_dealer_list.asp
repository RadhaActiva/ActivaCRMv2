<!-- #include file="database/datastore.asp" -->

<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
md_type=request("md_type") 
formname=request("formname") 
fieldname=request("fieldname")
fieldname1=request("fieldname1")
tech_wh=request("tech_wh")   
faulty=request("faulty")   
%>
<html>
<head>
<!-- #include file="meta.asp" -->
</head>

<body>

<%
  
'modified by  110922 to exclude qty lesser than 1
i = 1	

       sql2 = "select cust_debtor_code, cust_name,cust_branch_code, cust_address from tblcustomer "

        'if request("md_type") <> "All" then 
         '   sql2 = sql2 & " where tblmodel.md_type='" & request("md_type") & "' "
        'end if

        sql2 = sql2 & " where cust_type = 'Dealer' order by cust_name"
  
    
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql2
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
link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&tech_wh=" & tech_wh & "&formname=" & formname & "&fieldname=" & fieldname & "&fieldname1=" & fieldname1
'link = "&searchitem=" & searchitem & "&searchvalue=" & searchvalue & "&md_type=" & md_type & "&formname=" & formname & "&fieldname=" & fieldname & "&sf_fromwarehouse=" & sf_fromwarehouse

%>
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td class="style21"><font size="4"><strong>Nationwide Dealer List</strong></font></td>
  </tr>
  <tr> 
    <td class="style21"><form name="form1" method="post" action="rm_job_dealer_list.asp?tech_wh=<%=tech_wh%>">
<strong>Dealer Name</strong>
        <input type="text" name="searchvalue" value="<%=searchvalue%>">
        <input type="hidden" name="formname" value="<%=formname%>">
        <input type="hidden" name="fieldname" value="<%=fieldname%>">
        <input type="hidden" name="fieldname1" value="<%=fieldname1%>">
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
					Response.Write " <a href='rm_job_dealer_list.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_dealer_list.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %>
    </td>
  </tr>
  <tr> 
    <td valign="top"><table border="1" cellpadding="5" cellspacing="0" bordercolor="#E8E8E8">
        <tr valign="top" bgcolor="#88c0a7"> 
          <td width="3%"><strong>No.</strong></td>
          <td width="12%"><strong>Debtor Code</strong></td>
          <td width="12%"><strong>Branch Code</strong></td>
          <td><strong>Dealer Name</strong></td>
          <td width="40%" class='tktTotals'><strong>Address</strong></td>          
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
          <td>
         <!-- <a href="javascript:parent.opener.document.forms['<%=formname%>'].<%=fieldname%>.value='<%=rs("cust_code")%>';parent.opener.document.forms['<%=formname%>'].<%=fieldname1%>.value='<%=ChkString(rs("cust_branch_code"))%>';parent.opener.document.forms['<%=formname%>'].jobp_unitcost.value='<%=ChkNumber2(rs("md_rcpprice"))%>';parent.opener.document.forms['<%=formname%>'].jobp_qty.value='1';parent.opener.document.forms['<%=formname%>'].jobp_subtotal.value='<%=ChkNumber2(rs("md_rcpprice"))%>';parent.window.close();"><%=rs("cust_code")%></a>-->
              <a href="javascript:parent.opener.document.forms['<%=formname%>'].<%=fieldname%>.value='<%=rs("cust_code")%>';parent.opener.document.forms['<%=formname%>'].<%=fieldname1%>.value='<%=ChkString(rs("cust_branch_code"))%>';parent.window.close();"><%=rs("cust_code")%></a>
     </td>
          <td><%=rs("cust_branch_code")%></td>
          <td><%=rs("cust_name")%></td>
          <td><%=rs("cust_address")%></td>
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
					Response.Write " <a href='rm_job_dealer_list.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_job_dealer_list.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %> </td>
  </tr>
</table>
</body>
</html>

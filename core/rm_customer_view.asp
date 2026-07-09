<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")
cust_type = request("cust_type")

if ordertype = "" then 
   ordertype = "desc"
end if

i = 1
sql = "SELECT cust_id, cust_createddate, cust_createdby, cust_JS_receivedby, cust_JS_receiveddate, cust_code, cust_name, cust_type, " & _
      "cust_status, cust_reg_no, cust_company, cust_address, cust_postcode, cust_state, cust_state_id, cust_city, cust_city_id, cust_cnty_id, cust_email,  " & _
      "cust_tel1, cust_tel2, cust_fax, cust_website, cust_password, cust_gstregno, cust_lastjob_code, cust_source, cust_attention, cust_pic " & _
	  "FROM tblcustomer  where cust_id is not null "
	  
if cust_type <> "" then 
   sql = sql & " and cust_type = '" & cust_type& "' "
end if

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

'if orderby <> "" then
'sql = sql & " order by " & orderby & " " & ordertype
'else
sql = sql & " order by cust_name"
'end if

'response.write request.Cookies("GAPS")("slevel") & "<br>"

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
link = "&searchitem=" & request("searchitem") & "&searchvalue=" & request("searchvalue") & "&sortby=" & request("sortby") & "&cust_type=" & request("cust_type")

%>  
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Customer</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_customer_view.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" valign="top" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td width="84%">
                          <select name="searchitem" id="searchitem">
                            <option value="tblcustomer.cust_name" <% if searchitem = "tblcustomer.cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblcustomer.cust_code" <% if searchitem = "tblcustomer.cust_code" then response.write " selected" %>>Customer Code </option>
                            <option value="tblcustomer.cust_tel1" <% if searchitem = "tblcustomer.cust_tel1" then response.write " selected" %>>Customer Mobile No</option>
                            <option value="tblcustomer.cust_city" <% if searchitem = "tblcustomer.cust_city" then response.write " selected" %>>Customer City</option>
                            <option value="tblcustomer.cust_address" <% if searchitem = "tblcustomer.cust_address" then response.write " selected" %>>Customer Address</option>
                          </select>
                          
                          <input name="searchvalue" type="text" id="searchvalue" size="40" value="<%=searchvalue%>"/>&nbsp;&nbsp;Customer Type
                          <select name="cust_type" id="cust_type">
                            <option value="">All</option>
                            <option value="customer" <%if cust_type="customer" then response.write " selected"%>>Customer</option>
                            <option value="dealer" <%if cust_type="dealer" then response.write " selected"%>>Dealer</option>
                          </select>
                            <input type="submit" name="button" id="button3" value="Submit" />
                          <br />
                       <!--    <select name="orderby" id="orderby">
                            <option value="tblcustomer.cust_name" <% if orderby = "tblcustomer.cust_name" then response.write " selected" %>>Customer Name</option>
                            <option value="tblcustomer.cust_code" <% if orderby = "tblcustomer.cust_code" then response.write " selected" %>>Customer Code </option>
                            <option value="tblcustomer.cust_tel1" <% if orderby = "tblcustomer.cust_tel1" then response.write " selected" %>>Customer Mobile No</option>
                            <option value="tblcustomer.cust_city" <% if orderby = "tblcustomer.cust_city" then response.write " selected" %>>Customer City</option>
                          <option value="tblcustomer.cust_email" <% if orderby = "tblcustomer.cust_email" then response.write " selected" %>>Customer Email</option>
                          </select>-->
                        <!--  <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>-->
                          
                          </td>
                      </tr>
                    </table>
                  </form></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CLng(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_customer_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_customer_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Customer Code<br />
                      </strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Name</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Mobile</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Address </strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> State</strong></font></td>
                      <td align="left" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> City</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Type</strong></font></td>
                      <!--<td align="right" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Last Job No</strong></font></td>-->
                    </tr>
                    
<% 
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
                    
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40"><%=j%> </td>
                      <td nowrap="nowrap"><strong><a href="rm_customer_new.asp?cust_code=<%=rs("cust_code")%>"><%=rs("cust_code")%></a></strong></td>
                      <td nowrap="nowrap"><%=rs("cust_name")%></td>
                      <td> <%=rs("cust_tel1")%> / <%=rs("cust_tel2")%> </td>
                      <td> <%=rs("cust_address")%></td>
                      <td align="left"> <%=rs("cust_state")%> </td>
                      <td align="left"> <%=rs("cust_city")%> </td>
                      <td> <%=rs("cust_type")%></td>                     
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
          
                <tr>
                  <td height="30" align="right" bgcolor="#FFFFFF"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                  <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_customer_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_customer_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
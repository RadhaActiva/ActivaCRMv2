<!-- #include file="header.asp" -->
<%
searchitem = request("searchitem")
searchvalue = request("searchvalue")
Searchor_date = request("Searchor_date")
orderby = request("orderby")
ordertype = request("ordertype")

if ordertype = "" then 
   ordertype = "desc"
end if

i = 1
sql = "SELECT tech_id, tech_code, tech_type, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_city, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_wh_code " & _
	  "FROM tbltechnician  where tech_id is not null "

if searchvalue <> "" then 
   sql = sql & " and " & searchitem & " like '%" & searchvalue& "%' "
end if

if orderby <> "" then
sql = sql & " order by " & orderby & " " & ordertype
else
sql = sql & " order by tech_code desc"
end if


'response.write request.Cookies("GAPS")("slevel") & "<br>"
'response.write sql

set rs = server.CreateObject("adodb.recordset")
set rs1 = server.CreateObject("adodb.recordset")
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
        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td class="titleblue1"><div align="left"><font color="#CC0000">View </font>Technician</div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><form name="form1" id="form1" method="post" action="rm_contractor_view.asp?type=searchdata">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="16%" class="titlegrey1"><div align="left"> Filtered by</div></td>
                        <td width="84%">
                          <select name="searchitem" id="searchitem">
                            <option value="tbltechnician.tech_name" <% if searchitem = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tbltechnician.tech_code" <% if searchitem = "tbltechnician.tech_code" then response.write " selected" %>>Technician Code </option>
                            <option value="tbltechnician.tech_tel1" <% if searchitem = "tbltechnician.tech_tel1" then response.write " selected" %>>Technician Mobile No</option>
                            <option value="tbltechnician.tech_area" <% if searchitem = "tbltechnician.tech_area" then response.write " selected" %>>Technician Location</option>
                          </select>
                          <input name="searchvalue" type="text" id="searchvalue" value="<%=searchvalue%>" />
                          <select name="orderby" id="orderby">
                           <option value="tbltechnician.tech_name" <% if orderby = "tbltechnician.tech_name" then response.write " selected" %>>Technician Name</option>
                            <option value="tbltechnician.tech_code" <% if orderby = "tbltechnician.tech_code" then response.write " selected" %>>Technician Code </option>
                            <option value="tbltechnician.tech_tel1" <% if orderby = "tbltechnician.tech_tel1" then response.write " selected" %>>Technician Mobile No</option>
                            <option value="tbltechnician.tech_area" <% if orderby = "tbltechnician.tech_area" then response.write " selected" %>>Technician Location</option>
                          </select>
                           <select name="ordertype" id="ordertype">
                            <option value="asc" <% if ordertype = "asc" then response.write " selected"%>>A-Z</option>
                            <option value="desc" <% if ordertype = "desc" then response.write " selected"%>>Z-A</option>
                          </select>
                          <input type="submit" name="button" id="button3" value="Submit" />
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
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_contractor_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_contractor_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="4" cellspacing="0">
                    <tr>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span>No</span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Technician Code
                      </strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> <span> Type </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong><span> </span><span> Name
                      </span></strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Mobile</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Email </strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong> Location</strong></font></td>
                      <td bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Store</strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Submitted </strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Accepted </strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Done </strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Posted </strong></font></td>
                      <td align="center" bgcolor="#475387" class="style1"><font color="#FFFFFF"><strong>Cancel </strong></font></td>
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

job_Submitted = 0
job_Accepted = 0
job_Done = 0
job_Posted = 0
job_Cancel = 0

sql1="select count(job_id) as totaljob, job_status, job_tech_code from tbljob " & _
     "where job_tech_code='" & rs("tech_code") & "' group by job_status, job_tech_code"
	 
	 'response.write sql1
	 'response.End()
	 
rs1.ActiveConnection = strconnect
rs1.Source = sql1
rs1.CursorLocation  = 3
rs1.Open
while not rs1.eof  
   if rs1("job_status") = "Submitted" then 
      job_Submitted = rs1("totaljob")
   elseif rs1("job_status") = "Accepted" then 
      job_Accepted = rs1("totaljob")
   elseif rs1("job_status") = "Done" then 
      job_Done = rs1("totaljob")
   elseif rs1("job_status") = "Posted" then 
      job_Posted = rs1("totaljob")
   elseif rs1("job_status") = "Cancel" then 
      job_Cancel = rs1("totaljob")
   end if
rs1.movenext
wend
rs1.close


%>
                    <tr bgcolor="<%=nbgcolor%>">
                      <td height="40"><%=j%></td>
                      <td nowrap="nowrap"><strong><a href="rm_contractor_new.asp?tech_code=<%=rs("tech_code")%>"><font color="#0000FF"><%=rs("tech_code")%></font></a></strong></td>
                      <td align="left"><%=rs("tech_type")%></td>
                      <td align="left"><%=rs("tech_name")%></td>
                      <td><%=rs("tech_tel1")%></td>
                      <td><%=rs("tech_email")%></td>
                      <td><%=rs("tech_area")%></td>
                      <td><%=rs("tech_wh_code")%></td>
                      <td align="center"><strong><%=job_Submitted%></strong></td>
                      <td align="center"><strong><%=job_Accepted%></strong></td>
                      <td align="center"><strong><%=job_Done%></strong></td>
                      <td align="center"><strong><%=job_Posted%></strong></td>
                      <td align="center"><strong><%=job_Cancel%></strong></td>
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
					Response.Write " <a href='rm_contractor_view.asp?num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_contractor_view.asp?num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
              </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
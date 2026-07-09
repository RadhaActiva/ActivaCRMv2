<!-- #include file="header.asp" -->
<%

set rs = server.CreateObject("adodb.recordset") 
sql = "SELECT tech_id, tech_code, tech_type, tech_name, tech_icno, tech_address, tech_postcode, tech_state, tech_state_id,  tech_city, tech_city_id, tech_email, tech_tel1, tech_tel2, " & _
      "tech_createdby, tech_cretateddate, tech_carmodel, tech_carplateno, tech_carcolour, tech_password, tech_status, tech_area, tech_area_id, tech_wh_code " & _
	  "FROM tbltechnician WHERE tech_code = '" & Request.Cookies("GAPS")("job_tech_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			tech_id = rs("tech_id") 
			tech_code = rs("tech_code") 
			tech_type = rs("tech_type") 
			tech_name = rs("tech_name") 
			tech_icno = rs("tech_icno")
			tech_address = rs("tech_address")
			tech_postcode = rs("tech_postcode") 
			tech_state = rs("tech_state") 
			tech_state_id = rs("tech_state_id") 
			tech_city = rs("tech_city") 
			tech_city_id = rs("tech_city_id") 
			tech_email = rs("tech_email") 
			tech_tel1 = rs("tech_tel1") 
			tech_tel2 = rs("tech_tel2") 
			tech_createdby = rs("tech_createdby") 
			tech_cretateddate = rs("tech_cretateddate") 
			tech_carmodel = rs("tech_carmodel")  
			tech_carplateno = rs("tech_carplateno") 
			tech_carcolour = rs("tech_carcolour") 
			tech_password = rs("tech_password") 
			tech_status = rs("tech_status") 
			tech_area = rs("tech_area")
			tech_area_id = rs("tech_area_id")
			tech_wh_code = rs("tech_wh_code")
		End If
		rs.Close    
%>


        <tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td colspan="2" align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><div align="left">View Spare Part Stock Balance</div></td>
                        <td width="23%" align="right" class="titleblue1">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>

                <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><strong><font color="#FF0000"><%=request("loginerr")%></font></strong></td>
                </tr>
                <tr>
                  <td width="49%" valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV4">
                    <tbody>
                      <tr>
                        <td colspan="2" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Technician  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician Code *</strong></font></td>
                        <td align="left"><%=tech_code%></td>
                      </tr>
                      <tr>
                        <td width="22%" align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician Name *</strong></font></td>
                        <td align="left"><%=tech_name%></td>
                      </tr>
                      <tr>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Technician ICNO *</strong></font></td>
                        <td align="left"><%=tech_icno%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Address *</strong></font></td>
                        <td align="left"><%=tech_address%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode*</strong></font></td>
                        <td align="left"><%=tech_postcode%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State*</strong></font></td>
                        <td align="left"><%=tech_state%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>City*</strong></font></td>
                        <td align="left"><%=tech_city%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td valign="top"><%=tech_email%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 1*</strong></font></td>
                        <td valign="top"><%=tech_tel1%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. 2</strong></font></td>
                        <td valign="top"><%=tech_tel2%></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Prepared by</strong></font></td>
                        <td valign="top"><%=tech_createdby%> @ <%=chkdatetime(tech_cretateddate)%></td>
                      </tr>
                    </tbody>
                  </table></td>
                  <td width="51%" valign="top" bgcolor="#FFFFFF"><table width="99%" border="1" align="right" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV3">
                    <tbody>
                      <tr bgcolor="#E8E8E8">
                        <td colspan="4" scope="col"><strong>Store Details</strong></td>
                      </tr>
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Location</strong></font></td>
                        <td align="left"><%=tech_area%></td>
                        <td align="left" nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Car Model</strong></font></td>
                        <td align="left"><%=tech_carmodel%></td>
                      </tr>
                     
                      <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Car Plate No. </strong></font></td>
                        <td align="left"><%=tech_carplateno%></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Car colour</strong></font></td>
                        <td><%=tech_carcolour%></td>
                      </tr>
                       <tr >
                         <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF">Password</font></strong></td>
                         <td align="left"><%=tech_password%></td>
                         <td align="left" bgcolor="#CD6155"><strong><font color="#FFFFFF">Status</font></strong></td>
                         <td><%=tech_status%></td>
                       </tr>
                       <tr >
                        <td nowrap="nowrap" bgcolor="#CD6155"><strong><font color="#FFFFFF">Store Code</font></strong></td>
                        <td align="left"><%=tech_wh_code%></td>
                        <td align="left" bgcolor="#CD6155">&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                </form>
                
                
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <form name="form1" id="form1" method="post" action="global_ma_repeating_new.asp?type=editRepeat">
                    </form>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#475387">
                          <td colspan="8" bgcolor="#E8E8E8"><strong><font size="2">Stock Items List</font></strong></td>
                        </tr>
                        <tr bgcolor="#475387">
                          <td width="4%" align="center"><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td width="21%" align="left"><font color="#FFFFFF"><strong>Item Code.</strong></font></td>
                          <td width="17%" align="left"><font color="#FFFFFF"><strong>Model Name</strong></font></td>
                          <td width="17%" align="center"><font color="#FFFFFF"><strong>Current Stock</strong></font></td>
                          <td width="16%" align="left"><font color="#FFFFFF"><strong>Category</strong></font></td>
                          <td width="10%" align="center"><font color="#FFFFFF"><strong> Type</strong></font></td>
                          <td width="8%" align="center"><font color="#FFFFFF"><strong>Brand</strong></font></td>
                          <td width="7%" align="center"><font color="#FFFFFF"><strong>RCP</strong></font></td>
                        </tr>
                        <%
i = 1
sql = "SELECT tblwarehouse_stock.wst_id, tblwarehouse_stock.wst_wh_code, tblwarehouse_stock.wst_itm_code, " & _
		"tblwarehouse_stock.wst_itm_current_qty, tblwarehouse_stock.wst_itm_min_qty,  " & _
		"tblwarehouse_stock.wst_itm_remarks, tblwarehouse_stock.wst_lastupdateby, tblwarehouse_stock.wst_lastupdatedate, " & _
		"tblmodel.md_code, tblmodel.md_category, tblmodel.md_desc, tblmodel.md_unitprice, tblmodel.md_type, tblmodel.md_brands, tblmodel.md_unitprice1, tblmodel.md_rcpprice " & _
		"FROM tblwarehouse_stock inner join tblmodel on tblwarehouse_stock.wst_itm_code = tblmodel.md_code " & _
		"where tblwarehouse_stock.wst_id is not null and tblwarehouse_stock.wst_wh_code='" & tech_wh_code & "' "
sql = sql & " order by tblwarehouse_stock.wst_itm_code"

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
link = ""

%>
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
                        <tr bgcolor="<%=nbgcolor%>">
                          <td align="center"><%=count%>.</td>
                          <td align="left"><%=rs("md_code")%></td>
                          <td align="left"><%=rs("md_desc")%></td>
                          <td align="center"><%=rs("wst_itm_current_qty")%></td>
                          <td align="left"><%=rs("md_category")%></td>
                          <td align="center"><%=rs("md_type")%></td>
                          <td align="center"><%=rs("md_brands")%></td>
                          <% if  tech_type = "IC" then%>
                                 <td align="center"><%=rs("md_unitprice1")%>
                          <%else%>
                                 <td align="center"><%=rs("md_rcpprice")%>
                          <%end if%>
                          </td>
                        </tr>
                        <%
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
                        <tr>
                          <td colspan="8" align="right"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%></font>of <font color="3366ff"> <%=pgCount%></font>:
                            <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_contractor_stockbalance.asp?wh_code=" & wh_code & "&num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_contractor_stockbalance.asp?wh_code=" & wh_code & "&num=" & Showed+row & link & "'> Next >></a>"
	End If
	
                    %></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td colspan="2" align="right" bgcolor="#FFFFFF" 
          scope="col">&nbsp;</td>
                    </tr>
                    <tr align="right">
                      <td colspan="2" bgcolor="#FFFFFF"></td>
                    </tr>
                    <tr>
                      <td></tbody></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="2" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->
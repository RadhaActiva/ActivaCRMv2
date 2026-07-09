<!-- #include file="header.asp" -->
<%


wh_name=request.querystring("wh_name")
wh_tel=request.querystring("wh_tel")
wh_fax=request.querystring("wh_fax")
wh_address=request.querystring("wh_address")
wh_postcode = request.form("wh_postcode")

if wh_postcode = "" then
    wh_postcode = request.QueryString("wh_postcode")
end if

set rs = server.CreateObject("adodb.recordset")

job_date_to = dateadd("d",1,date())

if request("wh_code") <> "" then	  
sql = "SELECT wh_id, wh_code, wh_name, wh_address, wh_postcode, wh_state_id, wh_state, wh_city_id, wh_city, wh_tel, wh_fax, wh_remark, wh_contact_person, wh_email, wh_status " & _
	  "FROM tblwarehouse WHERE wh_code = '" & request("wh_code") & "' "
		rs.Open sql,strconnect,0,1,&H0001
		If Not rs.EOF Then
			wh_id = rs("wh_id") 
			wh_code = rs("wh_code") 
			wh_name = rs("wh_name") 
			wh_address = rs("wh_address") 
			wh_postcode = rs("wh_postcode") 
			wh_state = rs("wh_state") 
			wh_state_id = rs("wh_state_id") 
			wh_city = rs("wh_city") 
			wh_city_id = rs("wh_city_id") 
			wh_tel = rs("wh_tel") 
			wh_fax = rs("wh_fax") 
			wh_remark = rs("wh_remark") 
			wh_contact_person = rs("wh_contact_person") 
			wh_email = rs("wh_email")
			wh_status = rs("wh_status")
		End If
		rs.Close
	  stype = "editWarehouse"	
	  actionname = "Save" 
 else    
	  stype = "addWarehouse"
	  actionname = "Save" 
end if

if wh_postcode <> "" then
    set rs1 = server.CreateObject("adodb.recordset")
     sql1 = "SELECT city_id, post_office, state_id, state_name from tblpostcode WHERE postcode = '" & wh_postcode & "' "
		rs1.Open sql1,strconnect,0,1,&H0001   
		If Not rs1.EOF Then
             wh_state_id = rs1("state_id") 'will auto populate state
             wh_state =  rs1("state_name")
             wh_city_id = rs1("city_id") 'will auto populate city
             wh_city = rs1("post_office")    
        end if
    rs1.close
end if

%> 
<script language="javascript">

    function getPostcode(p) {  
        //document.getElementById('cust_name').value = s;
        document.getElementById('wh_postcode').value = p;
        document.formorder.submit();
    }

</script>
<tr>
          <td><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="2" class="titleblue1">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="77%" class="titleblue1"><div align="left"><font color="#CC0000">Create </font>Store</div></td>
                        <td width="23%" align="right" class="titleblue1">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                
                 <form name="formorder" method="post" action="action.asp?type=<%=stype%>">
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV2">
                    <tbody>
                      <tr>
                        <td colspan="4" bgcolor="#E8E8E8" scope="col"><strong><font size="2">Store  
                          Information </font></strong></td>
                      </tr>
                      <tr>
                        <td width="21%" align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Store Code<br />
                        <font size="1">(System Generate) </font>                        </strong></font></td>
                        <td align="left"><label for="textfield9"><strong><%=wh_code%>
                          <input type="hidden" name="wh_code" id="wh_code" value="<%=wh_code%>" />
                        </strong></label></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Status</strong></font></td>
                        <td align="left"><select name="wh_status" id="wh_status">
                          <option value="Y" <%if wh_status="Y" then response.write " selected"%>>Y</option>
                          <option value="N" <%if wh_status="N" then response.write " selected"%>>N</option>
                        </select></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Store Name</strong></font></td>
                        <td align="left"><input name="wh_name" type="text" id="wh_name" value="<%=wh_name%>" size="40" maxlength="100" /></td>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Person in Charge/<br />
                          Technician
                        </strong></font></td>
                        <td align="left">
                          <select name="wh_contact_person" id="wh_contact_person">
                            <option value=""></option>
                            <%			
				sql = "SELECT tech_id, tech_code, tech_name FROM tbltechnician where tech_status = 'Y' "	
                set rs = server.CreateObject("adodb.recordset")
				rs.Open sql,strconnect,3,3,&H0001
                while Not rs.EOF
					  if (wh_contact_person) = (rs("tech_code")) then
					  response.write "<option value='" & rs("tech_code") & "' selected>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  else
					  response.write "<option value='" & rs("tech_code") & "'>" & rs("tech_code") & " - " & rs("tech_name")  & "</option>"
					  end if 					  
				rs.movenext
				wend
				rs.close					
				%>
                          </select></td>
                      </tr>
                      <tr>
                        <td align="left" valign="top" bgcolor="#CD6155"><font color="#FFFFFF"><strong> Address </strong></font></td>
                        <td colspan="3" align="left"><strong>
                          <textarea name="wh_address" cols="50" rows="5" wrap="virtual" id="wh_address"><%=wh_address%></textarea>
                        </strong></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Postcode</strong></font></td>
                        <td><label for="textfield10"></label>
                          <input name="wh_postcode" type="text" id="wh_postcode" onchange="getPostcode(this.value)" value="<%=wh_postcode%>" size="30" maxlength="50" />
                          <label for="textfield11"></label></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>Remark</strong></font></td>
                        <td><input name="wh_remark" type="text" id="wh_remark" value="<%=wh_remark%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>State</strong></font></td>
                        <td>
                        <input name="wh_state" type="text" id="wh_state" value="<%=wh_state%>" size="30" readonly maxlength="50" />
                        <input name="wh_state_id" type="hidden" id="wh_state_id" value="<%=wh_state_id%>" size="30" maxlength="50" /></td>
                        <td bgcolor="#CD6155"><font color="#FFFFFF"><strong>City</strong></font></td>
                        <td>
                        <input name="wh_city" type="text" id="wh_city" value="<%=wh_city%>" size="30" readonly maxlength="50" />
                        <input name="wh_city_id" type="hidden" id="wh_city_id" value="<%=wh_city_id%>" size="30" maxlength="50" /></td>
                      </tr>
                      <tr>
                        <td align="left" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Tel. No. </strong></font></td>
                        <td width="26%"><label for="textfield12"></label>
                          <input name="wh_tel" type="text" id="textfield12" value="<%=wh_tel%>" size="30" maxlength="50" /></td>
                        <td width="13%" bgcolor="#CD6155"><font color="#FFFFFF"><strong>Email </strong></font></td>
                        <td width="40%"><input name="wh_email" type="text" id="textfield13" value="<%=wh_email%>" size="30" maxlength="150" /></td>
                      </tr>
                      <tr>
                        <td colspan="4" align="right" valign="top"><input type="submit" name="button" id="button" value="<%=actionname%>" /></td>
                      </tr>
                    </tbody>
                  </table></td>
                </tr>
                </form>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr>
                  <td valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="3" cellspacing="0" bordercolor="#E5E5E5" id="ctl00_ContentPlaceHolder1_GV">
                    <tbody>
                    </tbody>
                    <form name="form1" id="form1" method="post" action="global_ma_repeating_new.asp?type=editRepeat">
                    </form>
                    <tr valign="top">
                      <td colspan="2" bgcolor="#FFFFFF" 
          scope="col"><table width="100%" border="0" cellspacing="0" cellpadding="8">
                        <tr bgcolor="#333333">
                          <td colspan="10" bgcolor="#E8E8E8"><strong><font size="2">Stock Items List</font></strong></td>
                          </tr>
                        <tr bgcolor="#475387">
                          <td width="4%" align="center"><font color="#FFFFFF"><strong>No</strong></font></td>
                          <td width="13%" align="left"><font color="#FFFFFF"><strong>Item Code.</strong></font></td>
                          <td width="17%" align="left"><font color="#FFFFFF"><strong>Model Name</strong></font></td>
                          <td width="11%" align="center"><font color="#FFFFFF"><strong>Current Stock</strong></font></td>
                          <td width="14%" align="left"><font color="#FFFFFF"><strong>Category</strong></font></td>
                          <td width="9%" align="center"><font color="#FFFFFF"><strong> Type</strong></font></td>
                          <td width="8%" align="center"><font color="#FFFFFF"><strong>Brand</strong></font></td>
                          <td width="8%" align="center"><font color="#FFFFFF"><strong>RCP</strong></font></td>
                          <td width="6%" align="center"><font color="#FFFFFF"><strong>Avg Cost</strong></font></td>
                          <td width="10%" align="center"><font color="#FFFFFF"><strong>Total Cost</strong></font></td>
                          </tr>
                        
<%
i = 1
		
sql = "Select sum(tblstocktran.stk_qty) grandqty, " & _
		"round(sum(tblstocktran.stk_qty*tblmodel.md_averageecost),2) as grandtotal  " & _
		"from tblstocktran inner join tblmodel " & _
		"on tblstocktran.stk_itm_code=tblmodel.md_code " & _
		"where tblstocktran.stk_reference ='" & wh_code & "' and tblstocktran.stk_reference <> '0' " 
set rs = server.CreateObject("adodb.recordset")
rs.ActiveConnection = strconnect
rs.Source = sql
rs.CursorLocation  = 3
rs.Open
if not rs.eof then
   grandqty = rs("grandqty") 
   grandtotal = rs("grandtotal") 
end if
rs.close
		
sql = "Select  tblstocktran.stk_reference, (select top 1 tblwarehouse.wh_name from tblwarehouse where tblwarehouse.wh_code=tblstocktran.stk_reference ) as wh_name, " & _
		"sum(tblstocktran.stk_qty) Totalstockqty, " & _
		"round(sum(tblstocktran.stk_qty*tblmodel.md_averageecost),2) as totalvalue,  " & _
		"tblmodel.md_code, tblmodel.md_category, tblmodel.md_desc, tblmodel.md_unitprice, tblmodel.md_type, tblmodel.md_brands, tblmodel.md_rcpprice, tblmodel.md_averageecost " & _
		"from tblstocktran inner join tblmodel " & _
		"on tblstocktran.stk_itm_code=tblmodel.md_code " & _
		"where tblstocktran.stk_reference ='" & wh_code & "' and tblstocktran.stk_reference <> '0' " & _
		"group by tblstocktran.stk_reference, " & _
		"tblmodel.md_code, tblmodel.md_category, tblmodel.md_desc, tblmodel.md_unitprice, tblmodel.md_type, tblmodel.md_brands, tblmodel.md_rcpprice, tblmodel.md_averageecost order by md_type, md_code"
		
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
                          <td align="center" nowrap="nowrap"><%=count%>.</td>
                          <td align="left"><strong><a href="javascript:popup('rm_warehouse_stockmovement.asp?job_date_from=01-Jan-2015&job_date_to=<%=job_date_to%>&stk_reference=<%=wh_code%>&stk_itm_code=<%=rs("md_code")%>&stype=In&wh_code=<%=rs("md_code")%>','cb18','scrollbars=yes,resizable=yes,width=500,height=500')"><%=rs("md_code")%></a></strong></td>
                          <td align="left"><%=rs("md_desc")%></td>
                          <td align="center"><%=rs("Totalstockqty")%></td>
                          <td align="left"><%=rs("md_category")%></td>
                          <td align="center"><%=rs("md_type")%></td>
                          <td align="center"><%=rs("md_brands")%></td>
                          <td align="center"><%=chknumber2(rs("md_rcpprice"))%></td>
                          <td align="center"><%=chknumber2(rs("md_averageecost"))%></td>
                          <td align="center"><%=chknumber2(rs("totalvalue"))%></td>
                          </tr>
 <%
totalcost =  totalcost + rs("totalvalue")
wst_itm_current_qty =  wst_itm_current_qty + rs("Totalstockqty")
count = count + 1 
i = i + 1
rs.MoveNext
Next
rs.Close
Set rs = Nothing
%>
 
 <tr bgcolor="<%=nbgcolor%>">
                        <td colspan="3" align="right" nowrap="nowrap"><strong>Total</strong></td>
                        <td align="center" nowrap="nowrap"><strong><%=(wst_itm_current_qty)%></strong></td>
                        <td align="right" nowrap="nowrap">&nbsp;</td>
                        <td align="right" nowrap="nowrap">&nbsp;</td>
                        <td align="right" nowrap="nowrap">&nbsp;</td>
                        <td align="right" nowrap="nowrap">&nbsp;</td>
                        <td align="right" nowrap="nowrap">&nbsp;</td>
                        <td align="center"><strong><%=chknumber2(totalcost)%></strong></td>
                      </tr>
                      
<tr bgcolor="<%=nbgcolor%>">
   <td colspan="3" align="right" nowrap="nowrap" bgcolor="#999999"><strong>Grand Total</strong></td>
   <td align="center" nowrap="nowrap" bgcolor="#999999"><strong><%=(grandqty)%></strong></td>
   <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
   <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
   <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
   <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
   <td align="right" nowrap="nowrap" bgcolor="#999999">&nbsp;</td>
   <td align="center" bgcolor="#999999"><strong><%=chknumber2(grandtotal)%></strong></td>
 </tr>
                        <tr>
                          <td colspan="10" align="right"><strong>Page</strong> <font color="3366ff"> <%=pagestartno%> </font>of <font color="3366ff"> <%=pgCount%> </font>:
                            <%	
	i = 0
	For j = 1 To pgCount
				If CInt(Showed) = ((j-1) * row) Then
					Response.Write "<font color=#000000><b>"& j &"</b></font>"
				Else
					Response.Write " <a href='rm_warehouse_new.asp?wh_code=" & wh_code & "&num=" & (j-1) * row & link & "'>"& j &"</a>"
				End If
			If Not j = pgCount Then Response.Write " "
	i = i + 1
	Next
	
	If Remain > row Then
	  Response.Write "<a href='rm_warehouse_new.asp?wh_code=" & wh_code & "&num=" & Showed+row & link & "'> Next >></a>"
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
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
              </table></td>
        </tr>
<!-- #include file="footer.asp" -->